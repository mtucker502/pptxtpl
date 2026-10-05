"""PptxTemplate — core render engine for pptxtpl.

Provides the main API for loading a .pptx template, rendering it with a
Jinja2 context dict, and saving the result.
"""

import re
from xml.sax.saxutils import escape

from lxml import etree
from jinja2 import Environment, BaseLoader, TemplateSyntaxError, meta
from pptx import Presentation
from pptx.oxml.ns import qn

from pptxtpl.xml_utils import preprocess_xml, dedupe_table_ids
from pptxtpl.richtext import RichText, Listing
from pptxtpl.slide_ops import clone_slide, _drop_slide_owned_rels
from pptxtpl.slide_loops import (
    SLIDE_FOR_RE as _SLIDE_FOR_RE,
    SLIDE_ENDFOR_RE as _SLIDE_ENDFOR_RE,
    scan_slide_tags,
    parse_loop_tree,
    expand_loop_tree,
)
from pptxtpl.exceptions import TemplateRenderError, InvalidTemplateError


# Regex for Jinja tags used to discover undeclared variables
_JINJA_TAG_RE = re.compile(r"(\{\{.*?\}\}|\{%.*?%\}|\{#.*?#\})", re.DOTALL)

# Slide-level conditional tags: {%slide if EXPR %} and {%slide endif %}
_SLIDE_IF_RE = re.compile(
    r"\{%-?\s*slide\s+if\s+(.*?)\s*-?%\}", re.DOTALL
)
_SLIDE_ENDIF_RE = re.compile(r"\{%-?\s*slide\s+endif\s*-?%\}")


def _strip_slide_tags(xml: str) -> str:
    """Remove {%slide ...%} tags from XML."""
    xml = _SLIDE_FOR_RE.sub("", xml)
    xml = _SLIDE_ENDFOR_RE.sub("", xml)
    xml = _SLIDE_IF_RE.sub("", xml)
    xml = _SLIDE_ENDIF_RE.sub("", xml)
    return xml


def _escape_value(value):
    """Recursively XML-escape string values in a context value."""
    if isinstance(value, str):
        return escape(value)
    if isinstance(value, dict):
        return {k: _escape_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return type(value)(_escape_value(v) for v in value)
    if isinstance(value, (RichText, Listing)):
        return str(value)
    return value


class PptxTemplate:
    """Load a PowerPoint template, render with Jinja2, and save.

    Usage::

        tpl = PptxTemplate("template.pptx")
        tpl.render({"name": "World"})
        tpl.save("output.pptx")
    """

    def __init__(self, template_path: str):
        self._template_path = template_path
        try:
            self._prs = Presentation(template_path)
        except Exception as exc:
            raise InvalidTemplateError(f"Cannot load template: {exc}") from exc

    @property
    def slides(self):
        """Access the presentation's slides."""
        return self._prs.slides

    def render(self, context: dict | None = None, jinja_env: Environment | None = None) -> None:
        """Render all slides with the given context dict.

        Args:
            context: Template variables dict. RichText and Listing values are
                     automatically converted to their XML/text representations.
            jinja_env: Optional custom Jinja2 Environment. If not provided,
                       a default environment is created.
        """
        if context is None:
            context = {}

        if jinja_env is None:
            jinja_env = Environment(loader=BaseLoader(), autoescape=False)
            jinja_env.globals.update({"RichText": RichText, "Listing": Listing})

        # Convert context values for Jinja2:
        # - RichText/Listing → their XML string representation (already escaped)
        # - Plain strings → XML-escaped to prevent invalid XML after rendering
        render_context = {k: _escape_value(v) for k, v in context.items()}

        # Parse the slide-loop tree once; validation errors (unbalanced or
        # sibling-sharing tags) surface here before any slides are touched.
        loop_roots = self._parse_loop_roots()

        # Validate that no {%slide if%} appears inside a {%slide for%} range,
        # which has undefined semantics (the if is evaluated against the
        # original context, not per-iteration).  This must run before Phase 1
        # below: a {%slide if%} referencing a not-yet-bound loop variable
        # (e.g. an inner if using the outer loop's variable) would otherwise
        # blow up evaluating the condition instead of reporting the real,
        # actionable error.
        self._validate_no_if_inside_for(loop_roots)

        # Phase 1: Identify conditional slides to remove (defer actual removal
        # so that sldIdLst length stays stable for slide-loop partname generation)
        cond_removals = self._find_false_conditional_slides(render_context, jinja_env)

        # Phase 2: Expand slide-level loops (clones slides, modifies slide list)
        rid_contexts = self._expand_slide_loops(loop_roots, render_context, jinja_env)

        # Phase 3: Remove false conditional slides now that cloning is done
        sldIdLst = self._prs.slides._sldIdLst
        for sldId, rId in cond_removals:
            self._drop_slide_by_rid(sldIdLst, sldId, rId)

        # Phase 4: Render each slide (use rId-based context lookup)
        for i, slide in enumerate(self._prs.slides):
            ctx = render_context.copy()
            rId = sldIdLst[i].get(qn("r:id"))
            if rId in rid_contexts:
                ctx.update(rid_contexts[rId])
            self._render_slide(slide, ctx, jinja_env)

    def _drop_slide_by_rid(self, sldIdLst, sldId, rId) -> None:
        """Remove a slide and clean up parts it owned (e.g. notesSlide).

        Without this, removing a slide that has speaker notes leaves the
        notes part orphaned (its only reference was the deleted slide).
        """
        try:
            slide_part = self._prs.part.related_part(rId)
        except KeyError:
            slide_part = None
        if slide_part is not None:
            _drop_slide_owned_rels(slide_part)
        sldIdLst.remove(sldId)
        self._prs.part.drop_rel(rId)

    def _parse_loop_roots(self):
        """Scan all slides for {%slide for/endfor%} tags and parse the tree."""
        per_slide_events = []
        for slide in self._prs.slides:
            xml_str = etree.tostring(slide._element, encoding="unicode")
            xml_str = preprocess_xml(xml_str)
            per_slide_events.append(scan_slide_tags(xml_str))
        return parse_loop_tree(per_slide_events)

    def _validate_no_if_inside_for(self, loop_roots) -> None:
        """Raise if any {%slide if%} appears inside a {%slide for%} range."""
        if_indices: list[int] = []
        for i, slide in enumerate(self._prs.slides):
            xml_str = etree.tostring(slide._element, encoding="unicode")
            xml_str = preprocess_xml(xml_str)
            if _SLIDE_IF_RE.search(xml_str):
                if_indices.append(i)
        for idx in if_indices:
            for root in loop_roots:
                if root.start <= idx <= root.end:
                    raise InvalidTemplateError(
                        "{%slide if%} cannot appear on a slide inside a "
                        "{%slide for%} ... endfor range"
                    )

    def _find_false_conditional_slides(
        self, context: dict, jinja_env: Environment
    ) -> list[tuple]:
        """Identify slides with falsy {%slide if EXPR %} conditions.

        Returns a list of ``(sldId_element, rId)`` tuples for slides that
        should be removed.  The actual removal is deferred so that the
        sldIdLst length stays stable for slide-loop partname generation.
        """
        sldIdLst = self._prs.slides._sldIdLst
        removals: list[tuple] = []

        for i, slide in enumerate(self._prs.slides):
            xml_str = etree.tostring(slide._element, encoding="unicode")
            xml_str = preprocess_xml(xml_str)
            matches = _SLIDE_IF_RE.findall(xml_str)
            if not matches:
                continue
            if len(matches) > 1:
                raise InvalidTemplateError(
                    "Only one {%slide if%} per slide is supported"
                )
            expr = matches[0].strip()
            try:
                expr_fn = jinja_env.compile_expression(expr)
                result = expr_fn(**context)
            except Exception as exc:
                raise TemplateRenderError(
                    f"Cannot evaluate slide condition '{expr}': {exc}"
                ) from exc
            if not result:
                sldId = sldIdLst[i]
                rId = sldId.get(qn("r:id"))
                removals.append((sldId, rId))

        return removals

    def _expand_slide_loops(
        self, loop_roots, context: dict, jinja_env: Environment
    ) -> dict:
        """Expand {%slide for%} loop trees into cloned slides.

        Supports single-slide loops, multi-slide groups, and nested loops
        (a slide loop whose range sits inside another loop's range).  Each
        root tree expands to a flat clone sequence via expand_loop_tree;
        inner iterables are evaluated lazily against the accumulated path
        context, so ``{%slide for c in region.cities %}`` sees the current
        ``region``.

        Returns a dict mapping rId → per-slide context overrides (loop
        variables, `as`-named helpers, and a ``loop`` helper for the
        innermost loop with index/first/last/length).
        """
        if not loop_roots:
            return {}

        def eval_iterable(expr, path_ctx):
            # Items come from render_context which is already recursively
            # XML-escaped by render(); don't double-escape.
            merged = {**context, **path_ctx}
            try:
                expr_fn = jinja_env.compile_expression(expr)
                return list(expr_fn(**merged))
            except Exception as exc:
                raise TemplateRenderError(
                    f"Cannot evaluate slide loop iterable '{expr}': {exc}"
                ) from exc

        # Map context by rId so indices stay correct across multiple expansions
        rid_context: dict[str, dict] = {}
        sldIdLst = self._prs.slides._sldIdLst

        # Keep template sldIds in sldIdLst during expansion so that
        # add_slide (which uses len(sldIdLst) for partnames) never
        # generates duplicate names.  Remove them all at the end.
        deferred_removals: list[tuple] = []  # (template_sldId, rId)

        # Process in reverse order so that earlier indices remain valid
        for root in reversed(loop_roots):
            emissions = expand_loop_tree(root, eval_iterable)

            # Collect template sldIds for all slides in the root range
            template_sldIds: list[tuple] = []
            for idx in range(root.start, root.end + 1):
                sldId = sldIdLst[idx]
                rId = sldId.get(qn("r:id"))
                template_sldIds.append((sldId, rId))

            if not emissions:
                deferred_removals.extend(template_sldIds)
                continue

            # Clone slides in emission order.  Do NOT touch sldIdLst between
            # clones, because add_slide uses len(prs.slides) to generate
            # unique part names.
            n_before = len(list(sldIdLst))
            for slide_idx, _ctx in emissions:
                clone_slide(self._prs, self._prs.slides[slide_idx])

            # Collect the clone sldIds (they were appended at the end)
            clone_sldIds = [
                sldIdLst[n_before + i] for i in range(len(emissions))
            ]

            # Remove clones from the end of sldIdLst
            for sldId in reversed(clone_sldIds):
                sldIdLst.remove(sldId)

            # Insert clones just before the first template slide
            template_pos = list(sldIdLst).index(template_sldIds[0][0])
            for i, clone_sldId in enumerate(clone_sldIds):
                sldIdLst.insert(template_pos + i, clone_sldId)

            # Mark all template slides for deferred removal
            deferred_removals.extend(template_sldIds)

            # Store per-slide context keyed by rId.  expand_loop_tree already
            # returns an independent dict per emission.
            for (slide_idx, ctx), clone_sldId in zip(emissions, clone_sldIds):
                rid_context[clone_sldId.get(qn("r:id"))] = ctx

        # Now remove all template slides and drop their relationships
        for template_sldId, rId in deferred_removals:
            self._drop_slide_by_rid(sldIdLst, template_sldId, rId)

        return rid_context

    def _render_slide(self, slide, context: dict, jinja_env: Environment) -> None:
        """Render a single slide's XML through the Jinja2 pipeline."""
        # Get the slide's XML element
        slide_element = slide._element

        # Serialize to XML string
        xml_str = etree.tostring(slide_element, encoding="unicode")

        # Preprocess: fix fragmented delimiters, strip internal tags, etc.
        xml_str = preprocess_xml(xml_str)

        # Strip slide-level loop tags (already handled by _expand_slide_loops)
        stripped_xml_str = _strip_slide_tags(xml_str)

        # Bail out only if nothing changed AND no Jinja tags remain. A slide
        # holding only {%slide for/endfor/if/endif%} tags (no other Jinja
        # markup) still needs the stripped text committed below, or the
        # literal tag text leaks into the rendered output.
        if stripped_xml_str == xml_str and not _JINJA_TAG_RE.search(xml_str):
            return  # No templates on this slide
        xml_str = stripped_xml_str

        # Render with Jinja2
        try:
            template = jinja_env.from_string(xml_str)
            rendered_xml = template.render(context)
        except TemplateSyntaxError as exc:
            raise TemplateRenderError(
                f"Jinja2 syntax error on slide: {exc}"
            ) from exc
        except Exception as exc:
            raise TemplateRenderError(
                f"Rendering failed on slide: {exc}"
            ) from exc

        # Post-process: convert \n to line breaks, \a to paragraph breaks
        rendered_xml = self._post_process(rendered_xml)

        # Regenerate duplicate a16:rowId / a16:colId values introduced when
        # {%tr for ... %} loops cloned the template row.  Without this,
        # PowerPoint Online collapses rows that share an id.
        rendered_xml = dedupe_table_ids(rendered_xml)

        # Parse the rendered XML back into an element tree
        try:
            new_element = etree.fromstring(rendered_xml.encode("utf-8"))
        except etree.XMLSyntaxError as exc:
            raise TemplateRenderError(
                f"Rendered XML is invalid: {exc}"
            ) from exc

        # Replace the slide's element tree
        parent = slide_element.getparent()
        if parent is not None:
            parent.replace(slide_element, new_element)
            slide._element = new_element
        else:
            # Slide is the root element — replace children
            slide_element.clear()
            for attr_name, attr_val in new_element.attrib.items():
                slide_element.set(attr_name, attr_val)
            for child in new_element:
                slide_element.append(child)

    def _post_process(self, xml: str) -> str:
        """Convert escape sequences in rendered text to PowerPoint XML.

        - ``\\n`` inside <a:t> elements becomes ``<a:br/>``

        TODO: ``\\a`` paragraph break is not yet implemented; it would
        require splitting the enclosing <a:p> while cloning <a:pPr>.
        """
        # Handle \n → line break within <a:t> elements
        # We replace \n in text content with </a:t></a:r><a:br/><a:r><a:t>
        xml = self._replace_newlines_in_text(xml)
        return xml

    def _replace_newlines_in_text(self, xml: str) -> str:
        """Replace literal \\n characters inside <a:t> elements with <a:br/> elements.

        Each split run and the inserted <a:br/> carries the original run's
        <a:rPr> so PowerPoint Online preserves formatting (size, bold, font).
        Bare <a:r><a:t> with no rPr inherits from the layout/master in Online,
        which produces wrong sizes even when desktop PowerPoint renders fine.
        """

        run_re = re.compile(
            r"<a:r>(?:\s*(<a:rPr\b[^>]*(?:/>|>.*?</a:rPr>)))?\s*"
            r"(<a:t\b[^>]*>)(.*?)</a:t>\s*</a:r>",
            re.DOTALL,
        )

        def _replace_run(match: re.Match) -> str:
            rpr = match.group(1) or ""
            t_open = match.group(2)
            content = match.group(3)
            if "\n" not in content:
                return match.group(0)

            # Build a self-closing rPr to attach to <a:br>.  <a:br> only accepts
            # an empty rPr child, so collapse <a:rPr ...>...</a:rPr> if needed.
            br_rpr = ""
            if rpr:
                m = re.match(r"<a:rPr\b([^>]*?)/>", rpr)
                if m:
                    br_rpr = f"<a:rPr{m.group(1)}/>"
                else:
                    m = re.match(r"<a:rPr\b([^>]*?)>", rpr)
                    if m:
                        br_rpr = f"<a:rPr{m.group(1)}/>"
            br = f"</a:t></a:r><a:br>{br_rpr}</a:br><a:r>{rpr}{t_open}"

            parts = content.split("\n")
            return f"<a:r>{rpr}{t_open}{br.join(parts)}</a:t></a:r>"

        return run_re.sub(_replace_run, xml)

    def save(self, output_path: str, autofit: bool = False, **autofit_options) -> None:
        """Save the rendered presentation to a file or file-like object.

        Args:
            output_path: Destination path or file-like object.
            autofit: Shrink overflowing text to fit its shape. PowerPoint
                stores the shrink factor in the file and only recalculates it
                when text is edited, so rendered decks overflow until this is
                applied. Requires the ``autofit`` extra (Pillow).
            **autofit_options: Forwarded to
                :func:`pptxtpl.autofit.fit_slide` -- ``uniform`` to scale each
                slide by a single factor, ``grow`` to let shapes expand into
                free space before shrinking, plus ``default_size_pt``,
                ``font_path`` and ``min_scale``. Only valid with
                ``autofit=True``.
        """
        if autofit_options and not autofit:
            raise TypeError(
                "autofit options given without autofit=True: "
                + ", ".join(sorted(autofit_options))
            )
        if autofit:
            from pptxtpl.autofit import fit_presentation

            fit_presentation(self._prs, **autofit_options)
        self._prs.save(output_path)

    def get_undeclared_template_variables(
        self, jinja_env: Environment | None = None
    ) -> set[str]:
        """Find all undeclared variables across all slides.

        Returns a set of variable names that appear in Jinja2 expressions
        but are not defined in the environment's globals.
        """
        if jinja_env is None:
            jinja_env = Environment(loader=BaseLoader(), autoescape=False)

        all_vars: set[str] = set()

        for i, slide in enumerate(self._prs.slides):
            xml_str = etree.tostring(slide._element, encoding="unicode")
            xml_str = preprocess_xml(xml_str)
            xml_str = _strip_slide_tags(xml_str)

            try:
                ast = jinja_env.parse(xml_str)
            except TemplateSyntaxError as exc:
                raise TemplateRenderError(
                    f"Jinja2 syntax error on slide {i + 1}: {exc}"
                ) from exc
            variables = meta.find_undeclared_variables(ast)
            all_vars.update(variables)

        return all_vars

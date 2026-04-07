"""Slide-level operations: cloning and deletion."""

import copy
import re

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.slide import NotesSlidePart


def clone_slide(prs, source_slide):
    """Clone a slide including its XML content and relationships.

    The clone is appended at the end of the presentation.
    Returns the new slide object.
    """
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    # Collect source relationships (excluding the slide layout)
    src_rels = {}
    src_layout_rid = None
    for rel in source_slide.part.rels.values():
        if "slideLayout" in rel.reltype:
            src_layout_rid = rel.rId
        else:
            src_rels[rel.rId] = rel

    # Find the layout rId assigned to the new slide
    dst_layout_rid = None
    for rel in new_slide.part.rels.values():
        if "slideLayout" in rel.reltype:
            dst_layout_rid = rel.rId

    # Deep copy the source slide's XML element tree into the destination
    dst = new_slide._element
    for child in list(dst):
        dst.remove(child)
    for key in list(dst.attrib.keys()):
        del dst.attrib[key]

    for key, val in source_slide._element.attrib.items():
        dst.set(key, val)
    for child in source_slide._element:
        dst.append(copy.deepcopy(child))

    # Build rId remap table
    remap = {}
    if (
        src_layout_rid
        and dst_layout_rid
        and src_layout_rid != dst_layout_rid
    ):
        remap[src_layout_rid] = dst_layout_rid

    # Copy non-layout relationships and track rId changes
    for src_rid, rel in src_rels.items():
        if rel.is_external:
            new_rid = new_slide.part.relate_to(
                rel.target_ref, rel.reltype, is_external=True
            )
        elif rel.reltype == RT.NOTES_SLIDE:
            # NotesSlide parts are 1:1 with slides and contain a back-reference
            # to their owning slide. Reusing the source's notesSlide would both
            # share notes content across clones and leave a dangling back-ref
            # to the (later-deleted) template slide, producing an orphan part.
            # Create a fresh notesSlide for this clone and copy the source XML.
            src_notes_part = rel.target_part
            new_notes_part = _clone_notes_slide_part(
                new_slide.part, src_notes_part
            )
            new_rid = new_slide.part.relate_to(new_notes_part, RT.NOTES_SLIDE)
        else:
            new_rid = new_slide.part.relate_to(rel.target_part, rel.reltype)
        if new_rid != src_rid:
            remap[src_rid] = new_rid

    # Apply rId remapping to the copied XML
    if remap:
        _remap_rids(dst, remap)

    return new_slide


def _clone_notes_slide_part(new_slide_part, src_notes_part):
    """Create a new NotesSlidePart whose content mirrors src_notes_part.

    The new part is wired to the package's notes master and to
    `new_slide_part` (the back-reference). All non-master, non-slide
    relationships from the source notes part are copied. Any rId
    references inside the copied notes XML are remapped to the new rIds.
    """
    package = new_slide_part.package
    try:
        notes_master_part = package.presentation_part.notes_master_part
    except Exception as exc:
        # python-pptx normally creates a notes master on demand; if that
        # fails (e.g. an unusual presentation without one), surface a
        # clear error rather than a confusing AttributeError downstream.
        raise RuntimeError(
            "Cannot clone notes slide: presentation has no notes master "
            f"and one could not be created ({exc})"
        ) from exc

    new_notes_part = NotesSlidePart(
        package.next_partname("/ppt/notesSlides/notesSlide%d.xml"),
        src_notes_part.content_type,
        package,
        copy.deepcopy(src_notes_part._element),
    )
    new_notes_part.relate_to(notes_master_part, RT.NOTES_MASTER)
    new_notes_part.relate_to(new_slide_part, RT.SLIDE)

    # Copy any other relationships (images, hyperlinks, etc.) and remap rIds
    remap = {}
    for src_rid, rel in src_notes_part.rels.items():
        if rel.reltype in (RT.NOTES_MASTER, RT.SLIDE):
            continue
        if rel.is_external:
            new_rid = new_notes_part.relate_to(
                rel.target_ref, rel.reltype, is_external=True
            )
        else:
            new_rid = new_notes_part.relate_to(rel.target_part, rel.reltype)
        if new_rid != src_rid:
            remap[src_rid] = new_rid

    if remap:
        _remap_rids(new_notes_part._element, remap)

    return new_notes_part


def _drop_slide_owned_rels(slide_part):
    """Drop relationships owned by a slide that would otherwise orphan parts.

    Currently drops the notesSlide rel so the notes part isn't left as
    an orphan whose only reference was from a now-deleted slide.
    """
    notes_rids = [
        rid for rid, rel in slide_part.rels.items()
        if rel.reltype == RT.NOTES_SLIDE
    ]
    for rid in notes_rids:
        slide_part.drop_rel(rid)


def delete_slide(prs, slide_index):
    """Remove a slide from the presentation by its zero-based index."""
    sldIdLst = prs.slides._sldIdLst
    sldId = sldIdLst[slide_index]
    rId = sldId.get(qn("r:id"))
    slide_part = prs.part.related_part(rId)
    _drop_slide_owned_rels(slide_part)
    prs.part.drop_rel(rId)
    sldIdLst.remove(sldId)


def _remap_rids(element, remap):
    """Replace rId references in an XML element tree.

    Uses string replacement on the serialized XML to update all
    attribute values containing old rIds.
    """
    xml_str = etree.tostring(element, encoding="unicode")
    # Single-pass replacement to avoid chained-rename collisions
    # (e.g. {rId2: rId3, rId3: rId4} must not double-substitute).
    pattern = re.compile(
        r'"(' + "|".join(re.escape(k) for k in remap) + r')"'
    )
    xml_str = pattern.sub(lambda m: f'"{remap[m.group(1)]}"', xml_str)

    new_element = etree.fromstring(xml_str.encode("utf-8"))

    # Replace content in-place
    for child in list(element):
        element.remove(child)
    for key in list(element.attrib.keys()):
        del element.attrib[key]
    for key, val in new_element.attrib.items():
        element.set(key, val)
    for child in new_element:
        element.append(child)

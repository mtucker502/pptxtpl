# pptxtpl

Jinja2 templating for PowerPoint `.pptx` files. Like [docxtpl](https://github.com/elapouya/python-docxtpl) but for presentations.

- [Install](#install)
- [Quick start](#quick-start)
- [Template syntax](#template-syntax)
- [RichText](#richtext)
- [Slide loops](#slide-loops)
- [Conditional slides](#conditional-slides)
- [Table row loops](#table-row-loops)
- [Table cell conditionals](#table-cell-conditionals)
- [Inspecting templates](#inspecting-templates)
- [Sandboxed rendering](#sandboxed-rendering)

## Install

```bash
uv add git+https://github.com/mtucker502/pptxtpl.git
```

Optional extras:

```bash
uv add "pptxtpl[autofit] @ git+https://github.com/mtucker502/pptxtpl.git"
```

`autofit` pulls in Pillow, used to measure text so overflowing shapes can be shrunk to fit. See [Autofit](#autofit-shrink-text-on-overflow).

## Quick start

Create a `.pptx` template in PowerPoint (or with python-pptx) containing Jinja2 tags in text boxes, tables, or shapes. Then render it:

```python
from pptxtpl import PptxTemplate

tpl = PptxTemplate("template.pptx")
tpl.render({
    "title": "Q4 Review",
    "author": "Jane Smith",
    "items": ["Revenue up 18%", "NPS at 72", "3 new clients"],
})
tpl.save("output.pptx")
```

## Template syntax

Standard Jinja2 syntax works inside any text element.

### Variables

```
{{ title }}
{{ metrics.revenue }}
{{ team.0.name }}
```

### Conditionals

```
{% if executive_summary %}
{{ executive_summary }}
{% else %}
No summary provided.
{% endif %}
```

### For loops

```
{% for item in items %}
{{ item }}
{% endfor %}
```

```
{% for member in team %}
{{ member.name }} — {{ member.role }}
{% endfor %}
```

### Filters

```
{{ name|upper }}
{{ items|length }}
{{ description|default("N/A") }}
```

### Comments

```
{# This won't appear in the output #}
```

## RichText

Use `RichText` to inject styled inline text:

```python
from pptxtpl import PptxTemplate, RichText

rt = RichText("Revenue: ", bold=True)
rt.add("$4.2M", color="00B050", bold=True)
rt.add(" (target: $3.8M)")

tpl = PptxTemplate("template.pptx")
tpl.render({"summary": rt})
tpl.save("output.pptx")
```

The template just uses `{{ summary }}` — the formatting is applied at render time.

Supported styles: `bold`, `italic`, `underline`, `color` (hex), `size` (pt), `font`.

## Slide loops

Use `{%slide for %}` to duplicate an entire slide for each item in a list. Place the tags anywhere on the template slide — they're stripped before rendering.

### Single-slide loop

When both tags are on the same slide, that slide is duplicated once per item:

```
{%slide for project in projects %}
Name: {{ project.name }}
Status: {{ project.status }}
Tags: {{ project.tags | join(", ") }}
{%slide endfor %}
```

```python
from pptxtpl import PptxTemplate

tpl = PptxTemplate("template.pptx")
tpl.render({
    "projects": [
        {"name": "Atlas", "status": "On track", "tags": ["backend", "Q3"]},
        {"name": "Beacon", "status": "At risk", "tags": ["frontend", "Q3"]},
        {"name": "Comet", "status": "Complete", "tags": ["infra", "Q2"]},
    ],
})
tpl.save("output.pptx")
# → 3 slides, one per project
```

### Multi-slide loop

Place `{%slide for %}` on one slide and `{%slide endfor %}` on a later slide to duplicate the entire group as a unit. All slides between (and including) the two tags are cloned together for each iteration.

**In the template** (two slides per project):

```
Slide 1 (static):  Title

Slide 2 (loop start):
  {%slide for project in projects %}
  Summary: {{ project.name }}
  Status: {{ project.status }}

Slide 3 (loop end):
  Description: {{ project.description }}
  Tags: {{ project.tags | join(", ") }}
  {%slide endfor %}

Slide 4 (static):  Closing
```

**Render:**

```python
tpl = PptxTemplate("template.pptx")
tpl.render({
    "projects": [
        {"name": "Atlas", "status": "On track",
         "description": "Backend API platform.", "tags": ["backend", "Q1"]},
        {"name": "Beacon", "status": "At risk",
         "description": "Notification overhaul.", "tags": ["frontend", "Q1"]},
    ],
})
tpl.save("output.pptx")
# → 6 slides: Title, Atlas Summary, Atlas Details, Beacon Summary, Beacon Details, Closing
```

The group can span any number of slides. The loop variable and `loop` helper are available on every slide in the group.

### Loop helper

A `loop` helper is available on each cloned slide, mirroring Jinja2's loop variable:

```
Slide {{ loop.index }} of {{ loop.length }}
{% if loop.first %}(Introduction){% endif %}
{% if loop.last %}(Final){% endif %}
```

| Variable | Description |
|---|---|
| `loop.index` | 1-based iteration count |
| `loop.index0` | 0-based iteration count |
| `loop.first` | `True` on the first iteration |
| `loop.last` | `True` on the last iteration |
| `loop.length` | Total number of iterations |

If the list is empty, all template slides in the group are removed. Multiple slide loops (single or multi-slide) in one presentation work independently.

### Nested slide loops

Slide loops nest: place one `{%slide for %}` range inside another. Inner
iterables are evaluated per outer iteration, so they can reference the outer
loop variable. Each slide is cloned with the full context of every loop it
sits inside.

```
Slide 1: {%slide for region in regions %}   Region: {{ region.name }}
Slide 2: {%slide for city in region.cities %}
         City: {{ city.name }} in {{ region.name }}
         {%slide endfor %}
Slide 3: End of {{ region.name }}           {%slide endfor %}
```

```python
tpl.render({
    "regions": [
        {"name": "West", "cities": [{"name": "SF"}, {"name": "LA"}]},
        {"name": "East", "cities": [{"name": "NYC"}]},
    ],
})
# → 7 slides: West, SF, LA, End of West, East, NYC, End of East
```

`loop` always refers to the innermost enclosing slide loop. Empty inner
iterables skip only the inner slides for that iteration. Sibling loops
cannot share a slide (a slide is cloned as a unit).

### Named loop helpers

Add `as name` to bind a loop's helper under a stable name. Named helpers
are visible on every slide inside the loop — including nested loops — and,
unlike `loop`, are not shadowed by inline `{% for %}` loops.

```
{%slide for region in regions as regionloop %}
{%slide for city in region.cities as cityloop %}
City {{ cityloop.index }}/{{ cityloop.length }} of region {{ regionloop.index }}
{%slide endfor %}
{%slide endfor %}
```

Naming a helper `loop` or the same as the loop variable raises
`InvalidTemplateError`.

## Conditional slides

Use `{%slide if %}` to conditionally include or exclude entire slides based on the render context. Place the tags anywhere on the slide — they're stripped before rendering.

**In the template:**

```
{%slide if financials %}
Revenue: {{ financials.revenue }}
Profit: {{ financials.profit }}
{%slide endif %}
```

**Render:**

```python
from pptxtpl import PptxTemplate

tpl = PptxTemplate("template.pptx")

# Slide is included — financials is truthy
tpl.render({"financials": {"revenue": "$4.2M", "profit": "$1.1M"}})
tpl.save("with_financials.pptx")

# Slide is removed — financials is missing/falsy
tpl2 = PptxTemplate("template.pptx")
tpl2.render({})
tpl2.save("without_financials.pptx")
```

The condition is any valid Jinja2 expression:

```
{%slide if items|length > 0 %}
{%slide if show_section and has_data %}
{%slide if user.role == "admin" %}
```

Conditional slides are evaluated before slide loops, so a `{%slide if %}` can gate a section without needing the loop's iterable to exist when the condition is false.

## Table row loops

Use `{%tr for %}` to duplicate a table row for each item in a list. The opening and closing tags **must be in separate rows**: a marker row containing only `{%tr for ... %}`, one or more body rows that get repeated, and a marker row containing only `{%tr endfor %}`. The marker rows are removed; the body rows are cloned per iteration.

**In the template** (4 rows: header, for-marker, body, endfor-marker):

| Metric | Value | Status |
|---|---|---|
| `{%tr for m in metrics %}` | | |
| `{{ m.name }}` | `{{ m.value }}` | `{{ m.status }}` |
| `{%tr endfor %}` | | |

**Render:**

```python
tpl = PptxTemplate("template.pptx")
tpl.render({
    "metrics": [
        {"name": "Revenue", "value": "$4.2M", "status": "On track"},
        {"name": "NPS", "value": "72", "status": "Above target"},
        {"name": "Churn", "value": "3.1%", "status": "At risk"},
    ],
})
tpl.save("output.pptx")
# → Table has 4 rows: 1 header + 3 data rows
```

The `{%tr %}` prefix elevates the Jinja tag to the `<a:tr>` (table row) XML level, removing the marker row entirely. Placing both `{%tr for%}` and `{%tr endfor%}` inside the same row consumes both markers together and raises `InvalidTemplateError`.

Conditionals work the same way — `{%tr if condition %}` and `{%tr endif %}` in separate marker rows wrap the body row(s) between them.

## Table cell conditionals

Use `{%tc if %}` to conditionally include or exclude individual table cells. Place the opening tag in one cell and the closing tag in another — both cells are consumed by the directive, and the cells between them are conditionally rendered.

**In the template:**

| Name | `{%tc if show_detail %}` | Detail | `{%tc endif %}` | Score |
|---|---|---|---|---|

**Render:**

```python
tpl = PptxTemplate("template.pptx")

# Detail column is included
tpl.render({"show_detail": True, ...})

# Detail column is removed
tpl.render({"show_detail": False, ...})
```

The `{%tc %}` prefix elevates the Jinja tag to the `<a:tc>` (table cell) XML level. The cells containing the `{%tc %}` tags themselves are replaced by the bare Jinja directive, while the cells between them are conditionally included in the output.

**Note:** PowerPoint defines column widths in a fixed grid (`<a:tblGrid>`), so removing cells may affect the table layout. You may need to adjust column widths or use a merged cell to accommodate the conditional content.

## Autofit (shrink text on overflow)

Rendered text is usually longer than the placeholder text it replaced, so it overflows its shape. PowerPoint's "Shrink text on overflow" does not fix this on its own: the shrink factor is *stored in the file* as `<a:normAutofit fontScale="..."/>`, and PowerPoint only recalculates it when the text is edited in the UI. A generated deck therefore overflows until you click into the box and press a key.

Pass `autofit=True` to compute and store that factor at save time:

```bash
pip install "pptxtpl[autofit]"     # adds Pillow, used for text measurement
```

```python
tpl = PptxTemplate("template.pptx")
tpl.render(context)
tpl.save("output.pptx", autofit=True)
```

Each text frame is measured, a scale is found by binary search, and `<a:normAutofit>` is written with the result. Footer, slide-number, date, and header placeholders are skipped, since the master controls them.

### Keeping text proportional

By default each shape is scaled on its own, so a cramped title can end up much smaller than the body beneath it. Two options address that:

```python
tpl.save("output.pptx", autofit=True, grow=True, uniform=True)
```

`grow=True` lets a shape expand downward into empty space *before* any shrinking, so a tight box is not what forces the font down. A long title in a 33pt frame with 20pt of clear space under it keeps a much larger font this way. Only top-anchored frames are moved, a shape never grows past whatever sits below it, and the new height is kept only if it actually buys a larger font.

`uniform=True` scales every shape on a slide by the same factor — the smallest any one of them needs — so title and body keep their relative sizes.

They compose: `grow` raises the worst-case shape, and `uniform` then levels the slide at that higher factor.

Options are forwarded to `fit_shape`:

```python
tpl.save(
    "output.pptx",
    autofit=True,
    font_path="/path/to/Metric-Regular.ttf",  # measure with the deck's real font
    default_size_pt=18.0,                     # assumed when nothing declares a size
    min_scale=0.25,                           # PowerPoint's own floor
)
```

Measurement uses Pillow's bundled font unless `font_path` is given. That is accurate enough for a scale factor, but passing the deck's actual font is better when a shape is close to the boundary.

It can also be used directly on a `python-pptx` presentation:

```python
from pptx import Presentation
from pptxtpl import fit_presentation

prs = Presentation("deck.pptx")
for result in fit_presentation(prs):
    print(result.slide_index, result.shape_name, result.font_scale)
prs.save("deck.pptx")
```

## Inspecting templates

Find all undeclared variables across slides:

```python
tpl = PptxTemplate("template.pptx")
print(tpl.get_undeclared_template_variables())
# {'title', 'author', 'items', 'metrics', ...}
```

## Sandboxed rendering

By default, pptxtpl uses Jinja2's standard `Environment`, which does not restrict what template authors can access. If you render templates from untrusted sources (e.g., user-uploaded `.pptx` files), pass a `SandboxedEnvironment` to prevent access to private attributes and unsafe methods:

```python
from pptxtpl import PptxTemplate
from jinja2.sandbox import SandboxedEnvironment

tpl = PptxTemplate("untrusted_template.pptx")
tpl.render({"name": "World"}, jinja_env=SandboxedEnvironment())
tpl.save("output.pptx")
```

With sandboxing enabled, templates that attempt to access private attributes (e.g., `{{ obj.__class__ }}`) will raise a `jinja2.exceptions.SecurityError` instead of executing.

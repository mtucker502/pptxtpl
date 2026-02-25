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

## Install

```bash
uv add git+https://github.com/mtucker502/pptxtpl.git
```

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

**In the template** (a single slide):

```
{%slide for project in projects %}
Name: {{ project.name }}
Status: {{ project.status }}
Tags: {{ project.tags | join(", ") }}
{%slide endfor %}
```

**Render:**

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
| `loop.first` | `True` on the first slide |
| `loop.last` | `True` on the last slide |
| `loop.length` | Total number of slides |

If the list is empty, the template slide is removed entirely. Multiple slide loops in one presentation work independently.

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

Use `{%tr for %}` to duplicate a table row for each item in a list. Place the opening tag in the first cell and the closing tag in the last cell of the row you want to repeat.

**In the template** (a table with a header row and one template row):

| Metric | Value | Status |
|---|---|---|
| `{%tr for m in metrics %}{{ m.name }}` | `{{ m.value }}` | `{{ m.status }}{%tr endfor %}` |

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

The `{%tr %}` prefix elevates the Jinja tag to the `<a:tr>` (table row) XML level, so the loop wraps the entire row element.

Conditionals work the same way — `{%tr if condition %}...{%tr endif %}` to include or exclude a row.

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

## Inspecting templates

Find all undeclared variables across slides:

```python
tpl = PptxTemplate("template.pptx")
print(tpl.get_undeclared_template_variables())
# {'title', 'author', 'items', 'metrics', ...}
```

# pptxtpl

Jinja2 templating for PowerPoint `.pptx` files. Like [docxtpl](https://github.com/elapouya/python-docxtpl) but for presentations.

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

pptxtpl renders each slide's XML independently — there's no way for a Jinja2 `{% if %}` block to remove an entire slide, only to blank out its content.

Alternatively, to conditionally include slides, use python first to remove them from the presentation before rendering:

```python
import json
from pptx.oxml.ns import qn
from pptxtpl import PptxTemplate


def delete_slide(prs, slide_index):
    """Remove a slide by zero-based index."""
    sldIdLst = prs.slides._sldIdLst
    sldId = sldIdLst[slide_index]
    rId = sldId.get(qn("r:id"))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sldId)


# Map optional slide indices to required context keys
OPTIONAL_SLIDES = {
    1: "financials",
    2: "feedback",
    3: "roadmap",
}

with open("context.json") as f:
    context = json.load(f)

tpl = PptxTemplate("template.pptx")

# Remove slides for missing keys (reverse order preserves indices)
for idx in sorted(OPTIONAL_SLIDES, reverse=True):
    if OPTIONAL_SLIDES[idx] not in context:
        delete_slide(tpl._prs, idx)

tpl.render(context)
tpl.save("output.pptx")
```

## Inspecting templates

Find all undeclared variables across slides:

```python
tpl = PptxTemplate("template.pptx")
print(tpl.get_undeclared_template_variables())
# {'title', 'author', 'items', 'metrics', ...}
```

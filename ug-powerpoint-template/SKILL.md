---
name: ug-powerpoint-template
description: Build or restyle PowerPoint presentations in the University of Ghana house style using the bundled UG template. Use for UG slides, course presentations, research defences, project reports, group abstracts and requests to convert existing material into UG style.
---

# University of Ghana PowerPoint template

Create editable `.pptx` presentations based on the bundled University of Ghana template. The builder uses the supplied template and crest assets. It provides four content patterns: title, bullets, split and closing. It does not reproduce every possible official layout, and generated decks still require visual inspection.

## Workflow

1. **Understand the task.** Identify the audience, purpose, presentation length, required sections, citation style and whether the user supplied a source document or existing deck. Ask only for missing information that materially affects the result.
2. **Choose the starting template.** Use `assets/UG_Powerpoint_template.pptx` as the standard, clean starting point. If the user specifically prefers the image-rich version, use `assets/UG Powerpoint Template - With Images_new.pptx` as a visual reference or manual starting file. Keep the original photos and decorative elements only where relevant to the presentation. The current JSON builder uses the standard template automatically; it does not yet switch templates based on a JSON option.
3. **Ground the content.** Use the supplied source as the authority when the user asks to convert or summarise existing work. Preserve the meaning of claims, numbers, names, dates, quotations, citations and findings. Do not invent evidence or references. If content must be shortened, retain the central point and disclose material omissions or condensation.
4. **Plan the slides.** Use one main message per slide. Prefer 3 to 5 concise bullets, but do not force all content into bullets. Split dense content across slides. Use British English unless the user requests another convention. Keep slide prose natural and direct.
5. **Prepare the JSON spec.** Use the schema below. Put images in the same folder as the JSON or provide a valid absolute path. Add meaningful alt text for figures. Use speaker notes for presentation guidance, not as a substitute for essential slide content.
6. **Build the deck.** Run the script from the skill directory or wherever the skill is installed:
   ```bash
   python /path/to/ug-powerpoint-template/scripts/build_deck.py spec.json /path/to/output/deck.pptx
   ```
   The script locates its bundled standard template and crest assets relative to itself. It creates the output directory when needed, validates common input errors, saves the file and reopens it to check that the slide count is correct.
7. **Render and inspect.** Where LibreOffice or an equivalent renderer is available, render the deck to PDF or slide images. Inspect every slide at presentation size. Check title wrapping, text clipping, footer overlap, image/caption collisions, contrast, alignment, font consistency, legibility, source lines and empty placeholders. Automated sizing is an estimate, not a guarantee. If rendering is unavailable, state that visual validation was not completed.
8. **Fix and recheck.** Prefer shortening or splitting content, adjusting layout or resizing images before reducing font sizes. Do not shrink body text below 18 pt on standard bullet slides or 16 pt on split slide bullets. If the content cannot fit readably, use more slides or explain the limitation.
9. **Deliver transparently.** Provide the `.pptx` and, if useful, a PDF preview. Briefly mention what was condensed, any unavailable validation, and any unresolved content or citation gaps.

## Supported JSON specification

Set the optional top-level `template` field to `"standard"` (default) or `"image_rich"` to choose the bundled starting file.

```json
{
  "template": "image_rich",
  "slides": [
    {"type": "title", "title": "Presentation title", "author": "Name\\nDepartment or course"},
    {"type": "bullets", "title": "Main point", "bullets": ["Point", ["Sub point", "Sub point"], "Another point"], "notes": "Optional speaker notes"},
    {"type": "split", "title": "Evidence", "left_text": "Figure caption or short note", "image": "figure.png", "image_alt": "Description of the figure", "right_heading": "Key findings", "right_bullets": ["Finding one", "Finding two"]},
    {"type": "closing", "title": "Thank you", "author": "Questions"}
  ]
}
```

### Slide types

| Type | Use | Required content |
|---|---|---|
| `title` | Opening slide | `title`; optional `author` |
| `bullets` | Standard content slide | `title`, non-empty `bullets` array |
| `split` | Figure plus text, or two-column content | `title` and at least one of `image`, `left_text`, `right_heading`, `right_bullets` |
| `closing` | Thank-you or questions | `title`; optional `author` |

Every slide may include a `notes` string. In a bullet array, a nested array represents sub-bullets under the preceding main point. Avoid deeply nested lists; only one sub-bullet level is supported. Images on split slides are resolved relative to the JSON file when a relative path is supplied.

## Visual and content standards

- Follow the supplied template's A4 landscape dimensions, navy and gold colours, title band and University of Ghana crest lockup. The bundled standard template is the source of truth for the default build.
- The image-rich alternative is optional. Do not assume its photographs are suitable for every topic; retain, replace or remove them according to relevance and any applicable permissions.
- Keep titles concise. Long titles may need to be rewritten or split across a title and subtitle; do not assume a font-size change alone will prevent wrapping.
- Use the builder's intended font sizes as minimums, not proof that text fits. Actual wrapping varies by font and renderer.
- Keep figure captions short. Leave space for captions when sizing images; avoid placing text over figures unless the layout deliberately supports it.
- For charts and diagrams, use legible labels and a clear source note. Do not use colour alone to convey meaning. Ensure the figure remains readable when projected.
- Put brief source citations beside the claims they support. For academic presentations, create as many reference slides as needed to keep full references readable. Follow the citation style requested by the user or institution; do not assume Harvard if another style is specified.
- For tables, complex charts, equations or diagrams that do not fit the available patterns, create a clear figure as an image or propose a new layout. Do not squeeze complex content into a split slide merely to stay within the current templates.
- Do not invent statistics, findings, citations or references. Flag missing details and ask for source material when necessary.

## Restyling an existing deck

1. Read the original deck and identify its content, order, speaker notes, figures, tables, references and any required institutional wording.
2. Map each source slide to the nearest supported type. Preserve important information and verify that charts and images still convey the original meaning.
3. Summarise or reorganise only as needed for readability. Tell the user what was condensed, removed or moved. Do not claim a perfect conversion if the source contained unsupported layouts or elements.
4. Build, render, inspect every slide and fix issues before delivery when rendering tools are available.

## Files and limitations

- `assets/UG_Powerpoint_template.pptx`: bundled standard source template used by the builder.
- `assets/UG Powerpoint Template - With Images_new.pptx`: optional image-rich starting file selected with `"template": "image_rich"`.
- Template selection chooses the source presentation file, but the builder recreates slides using its supported layouts and custom placements. It may not preserve all original slide-specific photos or decorative elements. Inspect the result and use the image-rich PowerPoint manually if its original designed slides need to remain intact.
- `assets/logo_large.png` and `assets/logo_small.png`: crest assets used in title and footer lockups.
- `scripts/build_deck.py`: JSON-to-PowerPoint builder with input validation and output reopening checks.
- Only four slide patterns are implemented. Section dividers, native tables, native charts and specialised research layouts are not implemented as dedicated types.
- The script's fit calculations are heuristic. A successful build is not proof that every slide is visually correct. Always inspect rendered slides when the environment supports rendering.

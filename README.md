# University of Ghana PowerPoint Template Skill

An editable PowerPoint-generation skill for presentations using the University of Ghana house style, with a standard template and an optional image-rich alternative.

## What's included
- `ug-powerpoint-template/SKILL.md`: workflow and usage guidance.
- `ug-powerpoint-template/scripts/build_deck.py`: JSON-to-PowerPoint builder.
- `ug-powerpoint-template/tests/test_build_deck.py`: automated tests.
- `ug-powerpoint-template/CHANGELOG.md`: improvements in this release.

## Template options

- **Standard template:** `ug-powerpoint-template/assets/UG_Powerpoint_template.pptx`. This is the default used by the JSON builder.
- **Image-rich alternative:** `ug-powerpoint-template/assets/UG Powerpoint Template - With Images_new.pptx`. This is for students who prefer a template with images and want to adapt its existing visual style. Keep or replace its photos only when they suit the topic and any applicable permissions.

The builder currently uses the standard template automatically. The image-rich file is an optional manual starting point; choosing it is not yet a setting in the JSON specification.

## Required assets

The builder expects these files alongside the script:
- `ug-powerpoint-template/assets/UG_Powerpoint_template.pptx`
- `ug-powerpoint-template/assets/logo_large.png`
- `ug-powerpoint-template/assets/logo_small.png`

The optional image-rich template belongs in the same `assets/` folder with this exact filename:
- `ug-powerpoint-template/assets/UG Powerpoint Template - With Images_new.pptx`

## Quick start

Install dependencies:
```bash
python -m pip install python-pptx pillow
```

Create a JSON file such as `spec.json` following the schema in `ug-powerpoint-template/SKILL.md`, then run:
```bash
python ug-powerpoint-template/scripts/build_deck.py spec.json output/deck.pptx
```

The builder supports `title`, `bullets`, `split`, and `closing` slide types. It validates common input errors and reopens the generated deck to verify its slide count. Automated fit estimates do not replace visual inspection.

## Testing

Run:
```bash
python -m unittest discover -s ug-powerpoint-template/tests -v
```

## Limitations

This is a focused template skill, not a complete implementation of every official University of Ghana slide layout. Inspect rendered slides before presenting. Use the template and crest assets only in accordance with applicable institutional branding rules.

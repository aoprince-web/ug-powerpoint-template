# University of Ghana PowerPoint Template Skill

An editable PowerPoint-generation skill for presentations using the University of Ghana house style.

## What's included
- `ug-powerpoint-template/SKILL.md`: workflow and usage guidance.
- `ug-powerpoint-template/scripts/build_deck.py`: JSON-to-PowerPoint builder.
- `ug-powerpoint-template/tests/test_build_deck.py`: automated tests.
- `ug-powerpoint-template/CHANGELOG.md`: improvements in this release.

## Important: bundled assets
The builder expects these files alongside the script:
- `ug-powerpoint-template/assets/UG_Powerpoint_template.pptx`
- `ug-powerpoint-template/assets/logo_large.png`
- `ug-powerpoint-template/assets/logo_small.png`

The original assets are distributed in the companion source package. If they are not present in a checkout, download the source ZIP attached to the project release or add the three assets to the `assets/` folder before running the builder.

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

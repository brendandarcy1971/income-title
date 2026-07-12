# Income Title

This repository contains research, narrative planning, slide specifications, and deck-generation assets for the Income Title proposal workstream.

## Structure

- `research/`: canonical source research and background notes
- `docs/strategy/`: high-level proposal framing and executive narrative
- `docs/narrative/`: speaking materials
- `docs/slides/`: slide map, per-slide plans, and handoff documents
- `automation/`: Python scripts that generate slide decks
- `outputs/decks/`: generated PowerPoint files
- `outputs/images/`: slide imagery and exported assets
- `outputs/manifest.json`: build record for generated decks

## Repository layout

```text
income-title/
├── README.md
├── automation/
│   ├── build_slides.py
│   ├── generate_slide_04.py
│   ├── generate_slide_04_1_glossary.py
│   ├── generate_slide_08.py
│   └── generate_slides_06_07.py
├── docs/
│   ├── narrative/
│   │   └── speaking-script-45min.md
│   ├── slides/
│   │   ├── slide-map.md
│   │   ├── slide-04.plan.md
│   │   ├── slide-04.handoff.md
│   │   ├── slide-04-1-glossary.handoff.md
│   │   ├── slide-08.plan.md
│   │   ├── slide-08.handoff.md
│   │   ├── slide-10.plan.md
│   │   ├── slides-06-07.detailed.md
│   │   └── slides-06-07.handoff.md
│   └── strategy/
│       ├── executive-summary.md
│       └── proposal-brief.md
├── outputs/
│   ├── decks/
│   ├── images/
│   └── manifest.json
└── research/
    ├── income_title_outline.txt
    ├── new_home_development.txt
    ├── procurement.txt
    └── quotes.txt
```

## Typical workflow

1. Start from `docs/strategy/proposal-brief.md` and `docs/slides/slide-map.md`.
2. Update per-slide plans and handoff notes in `docs/slides/`.
3. Refresh source research only in `research/`.
4. Generate decks through `automation/build_slides.py`.
5. Review `.pptx` outputs in `outputs/decks/` and confirm the latest run in `outputs/manifest.json`.
6. Use `docs/narrative/speaking-script-45min.md` for presentation delivery.

## Build commands

From this folder:

```bash
# install slide-generation dependency
python3 -m pip install -r requirements.txt

# build all known slide decks
python3 automation/build_slides.py

# rebuild a subset
python3 automation/build_slides.py 04 08

# remove existing generated decks first, then rebuild
python3 automation/build_slides.py --clean
```

## Why this structure

- Source content and generated outputs are separated.
- Research has one canonical location.
- Slide planning files are easier to find by type.
- Deck generation now has a single entrypoint.
- Output decks and build history are grouped for review.

## Notes

- Generated binary assets remain checked into the repo under `outputs/`.
- If more slide generators are added, register them in `automation/build_slides.py`.
- Slide generation depends on `python-pptx`.

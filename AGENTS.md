# Income Title Agent Guidelines

## Purpose
This repository is for policy-story slide planning and reproducible slide generation.

## Default Workflow for Slide Tasks
1. Read the relevant plan in docs/slides.
2. Update or create the matching generator in automation.
3. Run the pipeline command with explicit slide IDs.
4. Confirm deck output in outputs/decks.
5. Confirm manifest update in outputs/manifest.json.

## Slide Generation Standards
- Keep generated slides reproducible via python scripts.
- Do not manually edit PPTX binaries.
- Preserve naming conventions: slide-XX.deck.pptx.
- Keep on-slide copy concise and speaker notes explanatory.

## Build and Validate
- Compile changed generators before running builds.
- Run: python3 automation/build_slides.py <slide-id>
- Verify produced file exists in outputs/decks.

## Planning File Standards
- Keep one plan per slide in docs/slides.
- Each plan should include objective, message, visual concept, layout, notes, and risks.
- Keep assumptions explicit and auditable.

## Policy Content Guardrails
- Frame claims as mechanisms with assumptions, not certainty.
- Distinguish illustrative scenarios from empirical estimates.
- For contentious claims, include rebuttal and risk controls.

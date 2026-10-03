---
name: slide-pipeline
description: 'Create or update policy presentation slides using the repo pipeline. Use when asked to build slides, add slide plans, generate pptx outputs, or align docs/slides with automation generators and outputs/manifest verification.'
argument-hint: 'Slide IDs and requested content changes'
user-invocable: true
---

# Slide Pipeline Skill

## When to Use
- User asks to create, revise, or generate one or more slides.
- A slide plan exists or must be created in docs/slides.
- A generator exists or must be created in automation.
- User asks why a slide is missing or not updated.

## Required Workflow
1. Read the target plan file in docs/slides.
2. Create or update the matching generator in automation.
3. Ensure automation/build_slides.py has the slide ID in SCRIPT_MAP and OUTPUT_MAP.
4. Run compile checks for changed Python files.
5. Run the build command for requested slide IDs.
6. Verify files in outputs/decks and manifest entries in outputs/manifest.json.

## Naming Convention
- Plan files: docs/slides/slide-XX.plan.md
- Generator files: automation/generate_slide_XX.py
- Deck files: outputs/decks/slide-XX.deck.pptx

## Output Contract
Always report:
- files changed
- build command run
- generated deck paths
- manifest confirmation

## Quality Checklist
- On-slide text is concise and readable at room distance.
- Notes explain assumptions and caveats.
- Data labels and totals are internally consistent.
- New IDs are buildable through automation/build_slides.py.

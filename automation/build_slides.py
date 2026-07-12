from __future__ import annotations

import argparse
import importlib.util
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
AUTOMATION_DIR = ROOT / "automation"
OUTPUTS_DIR = ROOT / "outputs"
DECKS_DIR = OUTPUTS_DIR / "decks"
MANIFEST_PATH = OUTPUTS_DIR / "manifest.json"

SCRIPT_MAP = {
    "03": "generate_slide_03_quotes.py",
    "03-1": "generate_slide_03_1_march_replay.py",
    "04": "generate_slide_04.py",
    "04-1": "generate_slide_04_1_glossary.py",
    "06-07": "generate_slides_06_07.py",
    "08": "generate_slide_08.py",
    "09": "generate_slide_09.py",
    "10": "generate_slide_10.py",
    "12": "generate_slide_12.py",
    "13": "generate_slide_13.py",
    "14": "generate_slide_14.py",
}

OUTPUT_MAP = {
    "03": "outputs/decks/slide-03-quotes.deck.pptx",
    "03-1": "outputs/decks/slide-03-1-march-replay.deck.pptx",
    "04": "outputs/decks/slide-04.deck.pptx",
    "04-1": "outputs/decks/slide-04-1-glossary.deck.pptx",
    "06-07": "outputs/decks/slides-06-07.deck.pptx",
    "08": "outputs/decks/slide-08.deck.pptx",
    "09": "outputs/decks/slide-09.deck.pptx",
    "10": "outputs/decks/slide-10.deck.pptx",
    "12": "outputs/decks/slide-12.deck.pptx",
    "13": "outputs/decks/slide-13.deck.pptx",
    "14": "outputs/decks/slide-14.deck.pptx",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build proposal slide decks.")
    parser.add_argument(
        "slides",
        nargs="*",
        help="Slide IDs to build, for example: 03 03-1 04 04-1 06-07 08 09 10 12 13 14. Defaults to all.",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Remove existing generated decks before building.",
    )
    return parser.parse_args()


def resolve_selection(selection: list[str]) -> list[str]:
    if not selection:
        return list(SCRIPT_MAP)

    unknown = [item for item in selection if item not in SCRIPT_MAP]
    if unknown:
        raise SystemExit(f"Unknown slide IDs: {', '.join(unknown)}")

    return selection


def ensure_dependencies() -> None:
    if importlib.util.find_spec("pptx") is None:
        raise SystemExit(
            "Missing dependency: python-pptx. Install it with `python3 -m pip install -r requirements.txt`."
        )


def clean_outputs() -> None:
    if not DECKS_DIR.exists():
        return

    for path in DECKS_DIR.glob("*.pptx"):
        path.unlink()


def run_script(script_name: str) -> dict[str, str]:
    script_path = AUTOMATION_DIR / script_name
    subprocess.run([sys.executable, str(script_path)], check=True, cwd=ROOT)
    return {
        "script": str(script_path.relative_to(ROOT)),
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }


def write_manifest(records: dict[str, dict[str, str]]) -> None:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "root": str(ROOT),
        "records": records,
    }
    MANIFEST_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    args = parse_args()
    selection = resolve_selection(args.slides)
    ensure_dependencies()

    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    DECKS_DIR.mkdir(parents=True, exist_ok=True)

    if args.clean:
        clean_outputs()

    manifest_records: dict[str, dict[str, str]] = {}
    for slide_id in selection:
        manifest_records[slide_id] = run_script(SCRIPT_MAP[slide_id])
        output_path = OUTPUT_MAP.get(slide_id, "(output path not registered)")
        print(f"Built slide {slide_id} -> {output_path}")

    write_manifest(manifest_records)
    print(f"Wrote manifest -> {MANIFEST_PATH.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
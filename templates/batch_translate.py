"""
Provision PPTX Translate - Batch Translation Script
=====================================================
Translate multiple PPTX files to one or more target languages.

Finds all .pptx files in an input directory and translates each
to every language specified in the config or via CLI.

Usage:
    python batch_translate.py --input-dir input/ --config config.json
    python batch_translate.py --input-dir input/ --languages Spanish Italian French
    python batch_translate.py --input-dir input/ --language Hebrew --config config.json
"""

import argparse
import json
import os
import sys
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Project directory
# ---------------------------------------------------------------------------
PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR / "templates"))

from translate_pptx import translate_pptx, load_config


# ===========================================================================
# Batch Translation
# ===========================================================================

def find_pptx_files(input_dir: str) -> list[str]:
    """Find all .pptx files in the input directory (non-recursive)."""
    pptx_files = []
    for f in sorted(os.listdir(input_dir)):
        if f.lower().endswith(".pptx") and not f.startswith("~$"):
            pptx_files.append(os.path.join(input_dir, f))
    return pptx_files


def batch_translate(
    input_dir: str,
    target_languages: list[str],
    config: dict,
) -> list[dict]:
    """
    Translate all PPTX files in input_dir to all target languages.
    Returns a list of result dicts with status info.
    """
    pptx_files = find_pptx_files(input_dir)

    if not pptx_files:
        print(f"No .pptx files found in: {input_dir}")
        return []

    print(f"\n{'=' * 60}")
    print(f"  Batch PPTX Translation")
    print(f"  Input directory: {input_dir}")
    print(f"  Files found:     {len(pptx_files)}")
    print(f"  Target languages: {', '.join(target_languages)}")
    print(f"  Total jobs:      {len(pptx_files) * len(target_languages)}")
    print(f"{'=' * 60}\n")

    results = []
    job_num = 0
    total_jobs = len(pptx_files) * len(target_languages)

    for pptx_path in pptx_files:
        filename = os.path.basename(pptx_path)

        for language in target_languages:
            job_num += 1
            print(f"\n--- Job {job_num}/{total_jobs}: {filename} -> {language} ---")

            result = {
                "file": filename,
                "language": language,
                "status": "pending",
                "output_path": None,
                "error": None,
            }

            try:
                output_path = translate_pptx(
                    input_path=pptx_path,
                    target_language=language,
                    config=config,
                )
                result["status"] = "success"
                result["output_path"] = output_path

            except Exception as e:
                result["status"] = "error"
                result["error"] = str(e)
                print(f"  ERROR: {e}")

            results.append(result)

            # Brief pause between jobs to respect API rate limits
            if job_num < total_jobs:
                time.sleep(2)

    # --- Print summary ---
    print(f"\n{'=' * 60}")
    print(f"  Batch Translation Summary")
    print(f"{'=' * 60}")

    success_count = sum(1 for r in results if r["status"] == "success")
    error_count = sum(1 for r in results if r["status"] == "error")

    for r in results:
        status_icon = "OK" if r["status"] == "success" else "FAIL"
        print(f"  [{status_icon}] {r['file']} -> {r['language']}")
        if r["output_path"]:
            print(f"        Output: {r['output_path']}")
        if r["error"]:
            print(f"        Error:  {r['error']}")

    print(f"\n  Total: {len(results)} | Success: {success_count} | Errors: {error_count}")
    print(f"{'=' * 60}\n")

    return results


# ===========================================================================
# CLI Entry Point
# ===========================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Provision PPTX Translate - Batch Translation"
    )
    parser.add_argument(
        "--input-dir", type=str, required=True,
        help="Directory containing .pptx files to translate",
    )
    parser.add_argument(
        "--config", type=str, default=None,
        help="Path to config.json file",
    )
    parser.add_argument(
        "--languages", type=str, nargs="+", default=None,
        help="Target languages (e.g., Spanish Italian French). Overrides config.",
    )
    parser.add_argument(
        "--language", type=str, default=None,
        help="Single target language (shortcut for --languages with one language)",
    )
    args = parser.parse_args()

    # Load config
    if args.config:
        config_path = args.config if os.path.isabs(args.config) else str(PROJECT_DIR / args.config)
        config = load_config(config_path)
    else:
        config = {}

    # Determine target languages
    if args.languages:
        target_languages = args.languages
    elif args.language:
        target_languages = [args.language]
    elif "target_languages" in config:
        target_languages = config["target_languages"]
    elif "target_language" in config:
        target_languages = [config["target_language"]]
    else:
        print("ERROR: No target languages specified. Use --languages, --language, or set in config.json")
        sys.exit(1)

    # Resolve input directory
    input_dir = args.input_dir if os.path.isabs(args.input_dir) else str(PROJECT_DIR / args.input_dir)

    if not os.path.isdir(input_dir):
        print(f"ERROR: Input directory not found: {input_dir}")
        sys.exit(1)

    results = batch_translate(input_dir, target_languages, config)

    # Save results summary
    output_dir = config.get("output_dir", "translated")
    results_dir = PROJECT_DIR / output_dir
    os.makedirs(results_dir, exist_ok=True)
    results_path = str(results_dir / "batch_results.json")
    with open(results_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"  Results saved to: {results_path}")


if __name__ == "__main__":
    main()

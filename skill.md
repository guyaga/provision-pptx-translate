---
name: provision-pptx-translate
description: "Translate pre-designed PowerPoint presentations to any language while preserving design, formatting, and animations. Translates slide text and speaker notes using Gemini 3.1 Pro. Keeps product names untranslated. Supports RTL languages. Use for: PPTX translation, presentation localization, multilingual slides. Triggers: provision translate pptx, translate presentation, localize pptx, provision pptx translate, translate slides"
allowed-tools: Read, Write, Edit, Bash, Glob, Grep
---

# Provision PPTX Translate

Translate pre-designed PowerPoint presentations to any language while preserving the original design, formatting, animations, and layout. This skill is for **translation ONLY** -- no video, no TTS, no dubbing.

## Quick Overview

This skill handles the complete PPTX translation workflow:

1. **Open** a .pptx file (which is a zip of XML files internally)
2. **Extract** all text from: slide text boxes, titles, subtitles, tables, charts, SmartArt, speaker notes, and group shapes
3. **Translate** all text to the target language using Gemini 3.1 Pro, preserving product names and technical terms
4. **Replace** the text in the PPTX while preserving: fonts, colors, sizes, bold/italic/underline, alignment, animations, transitions, embedded media
5. **Save** the translated PPTX as a new file with a language suffix

## Prerequisites

- **Python 3.10+**
- **python-pptx** library (for reading/writing PowerPoint files)
- **Google Gemini API key** (for translation via Gemini 3.1 Pro)
- **google-genai** Python package

Install dependencies:
```bash
pip install python-pptx google-genai
```

Set environment variable:
```bash
export GEMINI_API_KEY="your-gemini-api-key"
```

## Project Structure

```
project_dir/
  input/
    presentation.pptx       # Source PPTX file(s) to translate
  translated/
    presentation_Spanish.pptx    # Translated output
    presentation_Italian.pptx    # Another language variant
  config.json                    # Configuration file
  templates/
    translate_pptx.py            # Main translation script
    batch_translate.py           # Batch translation script
    validate_translation.py      # Translation validation helper
```

## Complete Workflow

### Step 1: Prepare the project

Create the project directory and place your PPTX file(s) in `input/`. The presentations should be fully designed -- this tool only translates text content.

### Step 2: Configure settings

Edit `config.json` to set:
- `target_language`: the language to translate to (e.g., "Spanish", "Italian", "Hebrew")
- `target_languages`: list of languages for batch translation
- `preserve_terms`: product names and technical terms that should NOT be translated (e.g., "Provision ISR", "DDA", "NVR")
- `translate_speaker_notes`: whether to also translate speaker notes (default: true)
- `output_dir`: where to save translated files

### Step 3: Run the translation

For a single file to a single language:
```bash
python templates/translate_pptx.py --input input/presentation.pptx --language Spanish
```

For a single file with config:
```bash
python templates/translate_pptx.py --input input/presentation.pptx --config config.json
```

For batch translation (multiple files and/or multiple languages):
```bash
python templates/batch_translate.py --input-dir input/ --config config.json
```

### Step 4: Validate the output

```bash
python templates/validate_translation.py --original input/presentation.pptx --translated translated/presentation_Spanish.pptx
```

### Step 5: Review output

Open the translated PPTX in PowerPoint to verify:
- All text has been translated
- Formatting is preserved (fonts, colors, sizes)
- Animations and transitions still work
- No text overflow or layout issues
- Product names remain untranslated

## Configuration

All settings are controlled via `config.json`. Key parameters:

| Parameter | Description | Default |
|-----------|-------------|---------|
| `gemini_model` | Gemini model for translation | `"gemini-3.1-pro"` |
| `target_language` | Single target language | `"Spanish"` |
| `target_languages` | List of languages for batch mode | `["Spanish", "Italian", "French"]` |
| `preserve_terms` | Terms to keep untranslated | `["Provision ISR", "DDA", "NVR", "VMS", "PTZ", "LPR"]` |
| `translate_speaker_notes` | Also translate speaker notes | `true` |
| `output_suffix` | Add language name as filename suffix | `true` |
| `output_dir` | Output directory for translated files | `"translated"` |

## How Text Extraction Works

The script extracts text from every possible location in a PPTX:

1. **Text frames**: Titles, subtitles, body text, content placeholders
2. **Tables**: Each cell is extracted and translated individually
3. **Group shapes**: Recursively descends into grouped shapes to find all text
4. **Speaker notes**: The notes_slide associated with each slide
5. **Charts and SmartArt**: Text labels and data labels where accessible via python-pptx

Text is extracted **run-by-run** (a "run" is a contiguous span of text with uniform formatting). This is critical because replacing text at the run level preserves all character-level formatting (font, size, color, bold, italic, underline).

## How Translation Works

Text segments are collected into batches and sent to Gemini 3.1 Pro with a detailed prompt:

- Translate to the specified target language
- Preserve any terms listed in `preserve_terms` exactly as-is
- Keep the same tone and formality as the original
- For technical terms, use the standard localized equivalent
- Return translations as a JSON mapping from original to translated text

The batching approach (rather than translating one text box at a time) gives Gemini full context of the presentation, resulting in more consistent translations.

## Handling RTL Languages (Arabic, Hebrew)

For RTL (right-to-left) languages like Arabic and Hebrew:

- The script detects RTL languages automatically based on the target language name
- Text alignment is flipped: left-aligned text becomes right-aligned
- Unicode RTL markers are NOT injected (PowerPoint handles bidi natively if the system has RTL support)
- Font substitution may occur if the original font does not support RTL characters -- ensure appropriate fonts are installed (e.g., Arial, David, Miriam for Hebrew; Arabic Typesetting, Simplified Arabic for Arabic)
- After translation, manually verify RTL rendering in PowerPoint

## Handling Text Overflow

When text is translated, it may become longer or shorter than the original. The script handles this by:

- **Preserving the original text box size** -- it does not resize shapes
- **Relying on PowerPoint's auto-fit** -- if the original text box has auto-shrink enabled, PowerPoint will shrink text to fit
- For critical presentations, review slides after translation and manually adjust:
  - Reduce font size if text overflows
  - Resize text boxes if needed
  - Split long text across multiple lines

A common rule of thumb: translations to German, Russian, or Portuguese are often 20-30% longer than English. Translations to Chinese, Japanese, or Korean are often shorter.

## Translating Speaker Notes

When `translate_speaker_notes` is `true` in the config:

- Speaker notes are extracted from each slide's `notes_slide`
- They are included in the translation batch alongside slide text
- Translated notes are written back to the PPTX
- Notes formatting (bold, italic) is preserved run-by-run, same as slide text

## Batch Translation

Use `batch_translate.py` to translate multiple PPTX files to one or more languages:

```bash
python templates/batch_translate.py --input-dir input/ --config config.json
```

This will:
1. Find all `.pptx` files in the input directory
2. For each file, translate to every language listed in `target_languages`
3. Save translated files to `output_dir` with language suffix
4. Print a summary of all translations performed

## Troubleshooting

### Corrupted PPTX after translation
- The script never modifies the original file -- it saves to a new file
- If the output is corrupted, try opening and re-saving the original in PowerPoint first, then re-running
- Ensure the source PPTX is a valid .pptx file (not .ppt or .ppsx)

### Missing fonts in translated file
- The script preserves the original font names -- it does NOT change fonts
- If the target language uses characters not in the original font, PowerPoint will substitute
- For best results, use fonts with broad Unicode coverage (Arial, Calibri, Noto Sans)
- For RTL languages, ensure RTL-capable fonts are installed on the system where the PPTX will be opened

### Broken formatting after translation
- This usually happens if text was replaced at the paragraph level instead of the run level
- The script replaces text run-by-run to avoid this
- If a paragraph has multiple runs and the translation merges them, the first run's formatting is used

### Empty text boxes in translated file
- Check if the Gemini API returned empty translations for some segments
- Run `validate_translation.py` to identify which text boxes are empty
- Re-run with a different batch size or check API quota

### Animations/transitions not working
- python-pptx preserves all XML that it does not explicitly modify
- Animations and transitions are stored in separate XML elements and should be untouched
- If issues persist, compare the XML of original and translated PPTX (both are zip files)

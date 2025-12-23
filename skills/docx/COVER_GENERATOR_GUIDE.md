# Cover generator quick guide

Use `generate_cover.py` (or a sibling wrapper) to update a cover template without ever modifying the source DOCX in-place. The process is automation-ready: everything runs in pure Python (no subprocesses) and can be driven by configuration values.

## Required placeholders
Keep the template pristine with these exact `<w:t>` contents:
- `No. 6461` — case number
- `APPELLANTS OPENING BRIEF` — filing name
- `Hon. Stacy Beckerman` — judge name

These values are also captured in `docx/cover_generator_config.yaml` so automation tooling can source them without touching code.

Only those scoped `<w:t>` values are replaced; other text remains untouched.

## Zero-subprocess workflow
The script works purely in Python with `zipfile`:
1. Open the `.docx` as a zip and read `word/document.xml`.
2. Replace the placeholders in the raw XML string.
3. Validate the updated XML via `defusedxml.ElementTree.fromstring`.
4. Repack with `zipfile`, preserving entry order and compression; the template file is never edited.

Because no shell commands are spawned, this flow can be embedded in automated systems (including local models) without extra sandbox permissions.

## Usage
```bash
python docx/generate_cover.py template.docx output.docx \
  --case-number "No. 1234" \
  --filing-name "APPELLANT'S BRIEF" \
  --judge "Hon. Jane Doe"
```

### Dry run
Preview exactly which placeholders would change without writing output:
```bash
python docx/generate_cover.py template.docx output.docx \
  --case-number "No. 1234" \
  --filing-name "APPELLANT'S BRIEF" \
  --judge "Hon. Jane Doe" \
  --dry-run
```

If XML validation fails after a replacement, the script reports which placeholder caused the issue.

## Automation notes
- Load `docx/cover_generator_config.yaml` to retrieve the required template markers and to confirm that subprocesses should stay disabled.
- Provide user inputs (case number, filing name, judge) via CLI arguments or your orchestration layer; the template file itself remains read-only.

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Run

```bash
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000

## Input Format

Markdown files with question blocks split by `Question No.` prefix. Each block contains:

- `Question No. <id>Q - ID : <qid>` header
- Question body text
- `Choice N` + following line (the choice text)
- `Selected Choice:` + answer(s)
- `Process Group :` / `Task/K. Area :` / `Justification:` sections
- `Domain:` / `Task:` / `Reference:` inside justification block

See [exam_input.md](exam_input.md) for canonical example. [output_format.md](output_format.md) shows input→output mapping.

## Architecture

Single-file Flask app ([app.py](app.py)). All logic lives here — no blueprints, no models layer.

**Data flow:** `.md` upload → `parse_markdown()` → list of dicts → format-specific generator → `send_file()`

**Parser** (`parse_markdown`): State machine. Splits on `(?=^Question No\.)`, then walks lines changing `state` variable (`question`, `choice`, `selected`, `process`, `task_area`, `justification`, `domain`, `task_domain`, `reference`). Emits one dict per question block.

**DOCX generator** (`generate_docx_bytes`): Landscape, 0.5" margins. Two questions per page via a 3-column outer table (left box | spacer | right box), each box containing an inner `_build_question_table`. Critical: `_d_fix_tbl_grid()` must be called on every table or Word ignores column widths. DXA constants at module top (`_D_Q_BOX`, `_D_LBL`, etc.).

**PDF generator** (`generate_pdf_bytes`): ReportLab `SimpleDocTemplate`, one question per page, single `Table` with 16 rows, row-index-based `TableStyle` background fills.

**CSV export**: Splits `Choices` (semicolon-joined string) back into `Choice 1`…`Choice 6` columns; writes UTF-8 BOM for Excel compat.

**Routes:**

- `GET /` — renders `templates/index.html`
- `POST /convert` — accepts `md_file` (multipart), `module_no` (string), `export_format` (`md`|`docx`|`pdf`|`csv`)
- `POST /shutdown` — kills process via `SIGINT`

## Key Constraints

- DOCX column widths are in DXA (1/1440 inch). Landscape content width = 14400 DXA total; each question box = 7020 DXA.
- `_d_move_into_cell()` re-parents an inner table's `_tbl` lxml element into an outer cell — must append a trailing `<w:p>` after or OOXML is invalid.
- `generate_sample.py` at root is a standalone script for generating test fixtures — not part of the app.

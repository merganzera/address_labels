# Address Label Generator

Create Avery 18660 label PDFs from an XLSX address list. The script formats
addresses into 1-4 lines, centers each line horizontally, and centers the full
address block vertically on each label.

## Setup

Use a virtual environment to keep dependencies local:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install openpyxl reportlab
```

## Usage

```bash
.venv/bin/python src/labels_from_xlsx.py \
  --input "Addresses Dec 16 2025.xlsx" \
  --output "label_outputs/labels_20251227_132227.pdf" \
  --font-name "Futura Medium" \
  --fit-mode shrink
```

### Fit modes

- `shrink` (default): shrink the entire label so the widest line fits.
- `wrap`: split overly long lines into two lines.
- `shrink-wrap`: shrink until `--max-shrink` is exceeded, then wrap instead.

Use `--max-shrink` with `shrink-wrap` to control the smallest allowed scale
(example: `0.85` means text will not shrink below 85%).

## Fonts

The script uses Jost if it is installed, otherwise it falls back to Helvetica.
To choose a font installed in macOS Font Book, pass `--font-name` (repeatable
or comma-separated). To use a specific font file, pass `--font /path/to/font.ttf`.

## Data handling

Address spreadsheets and generated PDFs are ignored by git. Keep datasets and
label outputs out of the repository.

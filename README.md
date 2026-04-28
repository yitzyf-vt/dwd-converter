# DavkaWriter (.dwd) → Word (.docx) Converter

High-fidelity converter that produces Word documents matching the visual style of DavkaWriter PDFs.

## Visual fidelity features
- Section headings (`YOU SHOULD KNOW`, `KEY WORDS`) styled as plain centered bold underlined Times Roman text matching original PDF
- Mishna headings (פרק א' משנה ב') with bold underlined David Hebrew
- Hebrew letter numbering (א, ב, ג, ד, ה...) for "YOU SHOULD KNOW" items
- Inline KEY WORDS layout: `הַסַּפָּר )א( - the barber`
- Page breaks at chapter boundaries (פרק יב, פרק יג...) and section transitions
- Paragraph alignment based on content language (RTL Hebrew, LTR English)
- Smart Davka Hebrew detection from diacritic bytes

## Format coverage
Supports all 5 DavkaWriter format variants (A through E).

## Run locally
```
pip install -r requirements.txt
python server.py
# Open http://localhost:5000
```

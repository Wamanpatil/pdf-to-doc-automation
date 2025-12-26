# PDF to MS Word Automation (Python)

This project recreates a legal PDF document (Form A – Mediation Application Form)
into an MS Word document using Python.

## Tech Stack
- Python 3
- python-docx

## Approach
The PDF was manually analyzed for structure, headings, and layout.
The document was recreated programmatically using python-docx with table-based
layouts to preserve formatting and readability.

## How to Run
```bash
pip install python-docx
python create_doc.py

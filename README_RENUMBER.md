Demo for ReferenceValidator renumbering

Overview
- This demo creates a small Word document with non-sequential citations and a numbered reference list, then runs `ReferenceValidator` to auto-renumber citations and the reference list.

Requirements
- Windows with MS Word installed (COM automation required)
- Python with `pywin32` installed

Install dependencies (PowerShell):

```powershell
python -m pip install pywin32
```

Run the demo (PowerShell):

```powershell
cd "c:\Users\muraliba\PycharmProjects\Pre Prod Hub"
python .\tests\demo_renumber.py
```

What the demo does
- Creates `tests/demo_input.docx` with three citation paragraphs styled `cite_bib` in order [3], [1], [2], and a reference list with `bib_number` entries 1..3.
 - Creates `tests/demo_input.docx` with three inline character-styled citations using `cite_bib` in order [3], [1], [2] (citations are character-style runs inside paragraphs), and a reference list composed of paragraphs using the `REF-N` paragraph style where the numeric labels are character-styled with `bib_number` (entries 1..3).
- Calls `ReferenceValidator.validate(auto_renumber=True, save_path=tests/demo_output.docx)` to renumber citations and references and save the updated file.
- Prints the validation/renumbering results and mapping.

Notes
- Always use a copy of important documents; the demo writes a new file by default (`tests/demo_output.docx`).
- If your documents use different style names or inline field citations, you may need to adapt `validator.py` to match them.

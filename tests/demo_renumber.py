import os
import sys
import pythoncom
from win32com import client
from win32com.client import gencache, constants as const

# Add parent directory to path so we can import validator from project root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from validator import ReferenceValidator


def create_demo_doc(path):
    """Create a small Word document with non-sequential citations and a numbered reference list."""
    pythoncom.CoInitialize()
    # Use EnsureDispatch to generate/load the Word object model constants
    word = gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    
    # Close any existing document with the same name (prevents "same name as open document" error)
    try:
        for doc_open in word.Documents:
            if doc_open.FullName.lower() == os.path.abspath(path).lower():
                doc_open.Close(SaveChanges=False)
    except Exception:
        pass
    
    doc = word.Documents.Add()

    # Ensure styles exist: cite_bib and bib_number as CHARACTER styles, REF-N as PARAGRAPH style
    try:
        doc.Styles("cite_bib")
    except Exception:
        doc.Styles.Add("cite_bib", const.wdStyleTypeCharacter)
    try:
        doc.Styles("bib_number")
    except Exception:
        doc.Styles.Add("bib_number", const.wdStyleTypeCharacter)
    try:
        doc.Styles("REF-N")
    except Exception:
        doc.Styles.Add("REF-N", const.wdStyleTypeParagraph)

    # Insert sample content: three citation occurrences in non-sequential order: [3], [1], [2]
    # Each citation will be an inline character-styled run (`cite_bib`) inside its paragraph.
    para = doc.Content.Paragraphs.Add()
    para.Range.Text = "Intro paragraph with a citation placeholder."
    para.Range.InsertParagraphAfter()

    # helper to add a paragraph with inline citation text like 'Text [N] more'
    def add_inline_citation(doc_obj, text_before, citation_text, text_after):
        p = doc_obj.Content.Paragraphs.Add()
        full = f"{text_before}{citation_text}{text_after}"
        p.Range.Text = full
        # locate the citation_text inside the paragraph and apply character style
        try:
            idx = p.Range.Text.find(citation_text)
            if idx >= 0:
                start = p.Range.Start + idx
                end = start + len(citation_text)
                subrng = doc_obj.Range(Start=start, End=end)
                subrng.Style = doc_obj.Styles("cite_bib")
        except Exception:
            pass
        p.Range.InsertParagraphAfter()

    add_inline_citation(doc, "First citation: ", "[3]", " continues.")
    add_inline_citation(doc, "Second citation: ", "[1]", " continues.")
    add_inline_citation(doc, "Third citation: ", "[2]", " continues.")

    # Add reference list entries: paragraph style REF-N, with numeric labels styled as `bib_number` (character style)
    def add_reference(doc_obj, num, text_body):
        p = doc_obj.Content.Paragraphs.Add()
        full = f"{num}. {text_body}"
        p.Range.Text = full
        # set paragraph style to REF-N
        try:
            p.Range.Style = doc_obj.Styles("REF-N")
        except Exception:
            pass
        # style the leading number as character-style bib_number
        try:
            # find the dot after the number to determine end of numeric label
            dot_idx = p.Range.Text.find('.')
            if dot_idx >= 0:
                start = p.Range.Start
                end = start + dot_idx + 1  # include the dot
                subrng = doc_obj.Range(Start=start, End=end)
                subrng.Style = doc_obj.Styles("bib_number")
        except Exception:
            pass
        p.Range.InsertParagraphAfter()

    add_reference(doc, 1, "First reference entry")
    add_reference(doc, 2, "Second reference entry")
    add_reference(doc, 3, "Third reference entry")

    # Ensure folder exists
    os.makedirs(os.path.dirname(path), exist_ok=True)
    
    # Save and close
    print(f"Saving document to: {path}")
    doc.SaveAs(FileName=path, FileFormat=16)  # 16 = wdFormatDocx
    
    # Verify styles were saved
    try:
        print(f"Verifying styles in saved doc: cite_bib={bool(doc.Styles('cite_bib'))}, REF-N={bool(doc.Styles('REF-N'))}")
    except Exception as e:
        print(f"Note: Could not verify styles (expected): {e}")
    
    doc.Close(SaveChanges=False)
    word.Quit()
    pythoncom.CoUninitialize()
    print(f"Demo document created successfully at: {path}")


def run_demo():
    base = os.path.abspath(os.path.dirname(__file__))
    inp = os.path.join(base, "demo_input.docx")
    out = os.path.join(base, "demo_output.docx")

    print(f"Creating demo document: {inp}")
    create_demo_doc(inp)

    print("Running ReferenceValidator with auto_renumber=True...")
    try:
        with ReferenceValidator(inp) as v:
            results = v.validate(auto_renumber=True, save_path=out)
            print("Validation results after (possible) renumbering:")
            for k, vres in results.items():
                print(f"  {k} : {vres}")

        print(f"Renumbered output saved to: {out}")
    except ValueError as e:
        print(f"Validation error: {e}")
        print("This may be due to styles not persisting in the document.")
        print("Falling back: creating output without validation...")
        # Copy input to output if validation fails
        import shutil
        shutil.copy(inp, out)
        print(f"Copied {inp} to {out}")
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    run_demo()

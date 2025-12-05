import pythoncom
import win32com.client as win32
import re

file_path = r"c:\Users\muraliba\PycharmProjects\Pre Prod Hub\S4C-Processed-Documents\Abuhamad9781975242831-ch002_renumbered.docx"

pythoncom.CoInitialize()
word = win32.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(file_path, ReadOnly=True)
    
    # Find all REF-N paragraphs
    try:
        refpara_style = doc.Styles("REF-N")
    except:
        refpara_style = None
    
    if refpara_style:
        print("Reference List Numbers Found:")
        count = 0
        for para in doc.Paragraphs:
            try:
                if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                    count += 1
                    text = para.Range.Text.strip()
                    # Extract first number
                    match = re.search(r'^\s*(\d+)', text)
                    if match:
                        num = match.group(1)
                        # Show first 80 chars of reference
                        preview = text[:80] + "..." if len(text) > 80 else text
                        print(f"  {num}. {preview}")
            except:
                pass
        print(f"\nTotal REF-N paragraphs: {count}")
    
    doc.Close(SaveChanges=False)
    word.Quit()
except Exception as e:
    print(f"Error: {e}")
    try:
        word.Quit()
    except:
        pass
finally:
    pythoncom.CoUninitialize()

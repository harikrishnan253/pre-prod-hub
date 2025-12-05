import os
from validator import ReferenceValidator

# File mentioned by user
file_path = r"c:\Users\muraliba\PycharmProjects\Pre Prod Hub\S4C-Processed-Documents\Abuhamad9781975242831-ch002.docx"
output_path = r"c:\Users\muraliba\PycharmProjects\Pre Prod Hub\S4C-Processed-Documents\Abuhamad9781975242831-ch002_renumbered.docx"

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    print(f"Processing {file_path}...")
    try:
        validator = ReferenceValidator(file_path)
        with validator:
            # Enable auto-renumbering
            results = validator.validate(auto_renumber=True, save_path=output_path)
        
        print("Validation Complete.")
        print(f"Total References: {results['total_references']}")
        print(f"Total Citations: {results['total_citations']}")
        print(f"Renumbered: {results.get('renumber_attempt', {}).get('renumbered', False)}")
        if results.get('renumber_attempt', {}).get('renumbered'):
            print(f"Renumbered file saved to: {output_path}")
            print("Renumber Map (sample):", str(results['renumber_attempt'].get('map'))[:200])
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

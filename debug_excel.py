import traceback
from doc_parser import parse_excel

def main():
    filepath = r"c:\\Users\\yash badgujar\\Downloads\\TICO\\ACC_Ph3_WCCA-furtherEdits_31032026.xlsx"
    print(f"Parsing {filepath}")
    try:
        res = parse_excel(filepath)
        if not res:
            print("Result is empty or None")
            return
        print("Success! Keys:", list(res.keys()))
        print("Sections:", len(res.get('sections', [])))
        print("Tables:", len(res.get('tables', [])))
        # Print first section heading if exists
        if res.get('sections'):
            print("First section heading:", res['sections'][0].get('heading'))
    except Exception as e:
        print("Exception during parsing:")
        traceback.print_exc()

if __name__ == "__main__":
    main()

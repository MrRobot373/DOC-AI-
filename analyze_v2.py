import pandas as pd
import json
from doc_parser import parse_document

docx_path = r"c:\Users\yash badgujar\Downloads\TICO\ACC_Ph3_HardwareDesignDocument_WithAIChk.docx"
excel_path = r"c:\Users\yash badgujar\Downloads\TICO\Review_Report_ACC_Ph3_HardwareDesignDocument_WithAIChk_Pro_20260427_0522.xlsx"

print("Parsing document...")
doc_data = parse_document(docx_path)
raw_text = doc_data["raw_text"]

print("Reading Excel report...")
df = pd.read_excel(excel_path)

findings = []
for index, row in df.iterrows():
    if pd.isna(row['No']):
        continue
    findings.append({
        'id': row['No'],
        'category': row['Category'],
        'severity': row['Severity'],
        'page': row['Page'],
        'section': row['Section'],
        'comment': str(row['Comment']),
        'fix': row['Fix']
    })

print(f"Total findings: {len(findings)}")

categories_count = {}
for f in findings:
    cat = str(f['category']).encode('ascii', 'ignore').decode('ascii')
    categories_count[cat] = categories_count.get(cat, 0) + 1

print("\nFindings by Category:")
for cat, count in sorted(categories_count.items(), key=lambda x: x[1], reverse=True):
    print(f"  {cat}: {count}")

print("\nAnalyzing specific findings for correctness...")
correct_quotes = 0
total_quotes = 0

import re
quote_pattern = re.compile(r'"([^"]*)"')

sample_analysis = []

for f in findings:
    matches = quote_pattern.findall(f['comment'])
    if matches:
        for match in matches:
            if len(match) > 5: # Skip short quotes
                total_quotes += 1
                if match in raw_text:
                    correct_quotes += 1
                else:
                    # Let's try case-insensitive or removing newlines
                    clean_match = match.replace("\n", " ").strip().lower()
                    clean_text = raw_text.replace("\n", " ").lower()
                    if clean_match in clean_text:
                        correct_quotes += 1
                    else:
                        if len(sample_analysis) < 5:
                           sample_analysis.append((f, match))
                        
if total_quotes > 0:
    print(f"\nQuote accuracy: {correct_quotes}/{total_quotes} ({correct_quotes/total_quotes*100:.1f}%)")
else:
    print("\nNo quotes found to verify.")

if sample_analysis:
    print("\nSample of unmatched quotes (hallucinations or slight mismatches):")
    for f, match in sample_analysis:
        print(f"ID {f['id']} - {f['category']}:")
        print(f"  Quote: \"{match}\"")
        print(f"  Comment: {f['comment'][:150]}...\n")

# Dump some interesting findings
print("\nTop CRITICAL/MAJOR findings:")
top_findings = [f for f in findings if f['severity'] in ['CRITICAL', 'MAJOR']]
for f in top_findings[:3]: # Let's show top 3
    status = "VERIFIED IN TEXT" if any(q in raw_text for q in quote_pattern.findall(f['comment']) if len(q)>5) else "UNVERIFIED QUOTE"
    if total_quotes > 0 and '\"' not in f['comment']:
        status = "NO QUOTE IN COMMENT"
    cat_safe = str(f['category']).encode('ascii', 'ignore').decode('ascii')
    print(f"[{f['severity']}] {cat_safe} (Section: {f['section']}, Page: {f['page']}) - {status}")
    print(f"Comment: {f['comment']}")
    print("-" * 50)

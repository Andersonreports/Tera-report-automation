import json
import re

# Load the content
with open(r'C:\Users\ander_j\.gemini\antigravity\brain\4b81bcb2-bb0e-4bd1-b382-2c99a6caef67\.system_generated\steps\684\content.md', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# The content starts after the first few header lines in the markdown
content = "".join(lines[4:])
try:
    data = json.loads(content)
except:
    # If not valid JSON, it might be a string representation of a list
    # The file shows it's likely a JSON array of arrays
    # Let's try to find the start of the JSON
    json_start = content.find('[[')
    if json_start != -1:
        data = json.loads(content[json_start:])
    else:
        print("Could not find JSON data")
        data = []

tests = set()
for row in data:
    if len(row) > 2:
        test_val = row[2].strip()
        # Filter out headers and empty values
        if test_val and test_val.upper() not in ["TEST", "S.NO", "SAMPLE_NAME"]:
             tests.add(test_val)

# Normalization logic (simple version)
norm_tests = {}
for t in tests:
    # Basic normalization to match the JS logic
    nt = t
    if re.search(r'Whole Exome', t, re.I) or t.upper() == 'WES': nt = "Whole Exome Sequencing (WES)"
    if re.search(r'Clinical Exome', t, re.I) or t.upper() == 'CES': nt = "Clinical Exome Sequencing (CES)"
    if re.search(r'Focus carrier', t, re.I) or re.search(r'Advat Focus', t, re.I): nt = "Advat Focus Carrier Screening"
    if re.search(r'Female Infertility', t, re.I): nt = "Female Infertility Panel"
    if re.search(r'Male Infertility', t, re.I): nt = "Male Infertility Panel"
    
    if nt not in norm_tests:
        norm_tests[nt] = []
    norm_tests[nt].append(t)

print("Unique Tests found in Google Sheet:")
for nt in sorted(norm_tests.keys()):
    orig = ", ".join(set(norm_tests[nt]))
    print(f"- {nt} (Original entries: {orig})")

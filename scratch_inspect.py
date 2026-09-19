import os, sys
sys.stdout.reconfigure(encoding='utf-8')

app_path = r"c:\Users\quiet\OneDrive\Desktop\ai for work\project rostering\app.py"
with open(app_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

for i in range(6745, min(6770, len(lines))):
    print(f"{i+1}: {repr(lines[i])}")

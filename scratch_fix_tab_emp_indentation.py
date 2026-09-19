import os

app_path = r"c:\Users\quiet\OneDrive\Desktop\ai for work\project rostering\app.py"
with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

idx_emp = content.find("with tab_emp:")
idx_unavail = content.find("with tab_unavail:")

before = content[:idx_emp]
tab_emp_body = content[idx_emp + len("with tab_emp:"):idx_unavail]
after = content[idx_unavail:]

lines = tab_emp_body.splitlines()
clean_lines = []
for line in lines:
    s = line.strip()
    if s == "try:" or s.startswith("except Exception") or "Error rendering Staff Members Tab" in s or s == "st.exception(e)":
        continue
    clean_lines.append(line)

indented_lines = []
for line in clean_lines:
    if not line.strip():
        indented_lines.append("")
    else:
        stripped = line.strip()
        leading_spaces = len(line) - len(line.lstrip())
        # original base inside try: was 12 spaces, but currently at 16 spaces
        rel = max(0, leading_spaces - 16)
        indented_lines.append(" " * (12 + rel) + stripped)

new_emp_block = "with tab_emp:\n        try:\n" + "\n".join(indented_lines) + "\n        except Exception as e:\n            st.error(f\"⚠️ Error rendering Staff Members Tab: {e}\")\n            st.exception(e)\n\n    # --- TAB 5: UNAVAILABILITY ---\n    "

new_content = before + new_emp_block + after[len("with tab_unavail:"): ]

with open(app_path, "w", encoding="utf-8") as f:
    f.write(new_content)

print("SUCCESS reformatting tab_emp body!")

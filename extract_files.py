import os
import pandas as pd
from pathlib import Path
from datetime import datetime

cwd = os.getcwd()
or_folder = str(Path(cwd).parents[0]) + "\\(04) NDR Call Recordings"
output_folder = str(Path(cwd).parents[0]) + "\\Today\\Recording Proof"

ALLOWED_FORMATS = {".mp3", ".png", ".jpeg", ".jpg"}

extracted = {"mp3": [], "png": [], "jpeg": [], "jpg": []}

for file_name in os.listdir(or_folder):
    ext = os.path.splitext(file_name)[1].lower()
    if ext in ALLOWED_FORMATS:
        extracted[ext.lstrip(".")].append(file_name)

rows = []
total = 0
for fmt, files in extracted.items():
    if files:
        print(f"\n[{fmt.upper()}] — {len(files)} file(s)")
        for f in files:
            print(f"  {f}")
            rows.append({"File Name": f, "Format": fmt.upper()})
        total += len(files)

print(f"\nTotal: {total} file(s) found")

# Save to Excel
if rows:
    os.makedirs(output_folder, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_folder, f"extracted_files_{timestamp}.xlsx")
    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False)
    print(f"\nSaved to: {output_path}")
else:
    print("\nNo files found — Excel not created.")

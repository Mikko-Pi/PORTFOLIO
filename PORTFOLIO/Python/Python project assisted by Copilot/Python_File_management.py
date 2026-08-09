from pathlib import Path
import shutil

#THIS IS AN EASY VIBE CODED (WITH COPILOT) FILE MANAGEMENT SCRIPT FOR DEMONSTRATION

src = Path(r"C:\Users\mikko\OneDrive\Työpöytä\Demo")
dst = src / "Python result"

rules = {
    "Calibration protocols": ["calibration protocols", "calibration information"],
    "Parameter information": ["parameter information"],
    "Special Documentation": ["special documentation"],
}

# Create destination folders
for folder in ["Calibration protocols", "Parameter information", "Special Documentation"]:
    (dst / folder).mkdir(parents=True, exist_ok=True)

# Copy files
for f in src.iterdir():
    if not f.is_file():
        continue
    name = f.name.lower()
    for folder, keys in rules.items():
        if any(k in name for k in keys):
            shutil.copy2(f, dst / folder / f.name)
            break
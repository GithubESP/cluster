import json
from pathlib import Path

# === Settings ===
# Change these filenames to match your actual file names if needed.
INPUT_FILE = "stats.ndjson"          # your original ndjson file
AREA_FILE = "terms_area.ndjson"      # lines with "fromAreaMods"
NON_AREA_FILE = "terms_non_area.ndjson"  # lines without "fromAreaMods"

def main():
    input_path = Path(INPUT_FILE)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        return

    # Open all files with UTF-8 encoding
    with input_path.open("r", encoding="utf-8") as fin, \
         open(AREA_FILE, "w", encoding="utf-8") as f_area, \
         open(NON_AREA_FILE, "w", encoding="utf-8") as f_non_area:

        for idx, line in enumerate(fin, start=1):
            raw = line.strip()
            if not raw:
                # Skip empty lines
                continue

            try:
                obj = json.loads(raw)
            except json.JSONDecodeError:
                # If a line is not valid JSON, just skip (or you can log it)
                print(f"Warning: line {idx} is not valid JSON, skipped.")
                continue

            # Check if the key "fromAreaMods" exists in this object
            if "fromAreaMods" in obj:
                f_area.write(raw + "\n")
            else:
                f_non_area.write(raw + "\n")

    print("Done.")
    print(f"Area mods  file: {AREA_FILE}")
    print(f"Non-area file: {NON_AREA_FILE}")


if __name__ == "__main__":
    main()

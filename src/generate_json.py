import pandas as pd
import json
from constants import STOCK, JSON_FILE

# File paths

# Read Excel file
df = pd.read_excel(STOCK, engine="openpyxl")  # Ensure openpyxl is installed

# Convert to list of dictionaries
data_list = df.to_dict(orient="records")

# Save as JSON file
with open(JSON_FILE, "w", encoding="utf-8") as f:
    json.dump(data_list, f, indent=4, ensure_ascii=False)

print(f"JSON file saved: {JSON_FILE}")

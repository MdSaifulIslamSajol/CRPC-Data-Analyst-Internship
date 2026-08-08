import pandas as pd

path = r"H:\stats america\FFIEC_gov\FFIEC_gov_CensusTractList2026_CRPC_Area_Only3.xlsx"

df = pd.read_excel(path)

df["Tracts in decimal"] = df["Tract"].apply(lambda v: f"{str(int(v)).zfill(6)[:-2]}.{str(int(v)).zfill(6)[-2:]}")

FFIEC_govt_Parish_ALL_Tracts_Dictionary = (
    df.groupby("County name")["FFIEC_Geography"]
    .apply(list)
    .to_dict()
)

import json

output_path = r"H:\stats america\FFIEC_gov\FFIEC_govt_Parish_ALL_Tracts_Dictionary.json"
with open(output_path, "w") as f:
    json.dump(FFIEC_govt_Parish_ALL_Tracts_Dictionary, f, indent=2)

print(f"Dictionary saved to: {output_path}")
for parish, tracts in FFIEC_govt_Parish_ALL_Tracts_Dictionary.items():
    print(f"\n{parish.strip()} ({len(tracts)} tracts):")

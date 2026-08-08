import pandas as pd

input_path  = r"H:\stats america\FFIEC_gov\FFIEC_gov_CensusTractList2026_CRPC_Area_Only2.xlsx"
output_path = r"H:\stats america\FFIEC_gov\FFIEC_gov_CensusTractList2026_CRPC_Area_Only3.xlsx"

df = pd.read_excel(input_path)

def to_census_tract(value):
    padded = str(int(value)).zfill(6)
    return f"{padded[:-2]}.{padded[-2:]}"

df.insert(
    df.columns.get_loc("Tract") + 1,
    "Tracts in decimal",
    df["Tract"].apply(to_census_tract)
)

df.to_excel(output_path, index=False)

print(f"Saved: {output_path}")
print(f"\nSample conversion:")
print(df[["Tract", "Tracts in decimal"]].head(10).to_string(index=False))

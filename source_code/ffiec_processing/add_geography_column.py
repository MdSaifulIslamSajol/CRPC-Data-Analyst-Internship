import pandas as pd

path = r"H:\stats america\FFIEC_gov\FFIEC_gov_CensusTractList2026_CRPC_Area_Only3.xlsx"

df = pd.read_excel(path)

def tract_to_decimal(value):
    padded = str(int(value)).zfill(6)
    return f"{padded[:-2]}.{padded[-2:]}"

def build_geography(row):
    parish = row["County name"].strip().replace("PARISH", "").strip().title()
    tract = tract_to_decimal(row["Tract"])
    return f"{parish} LA Tract {tract}"

df["Tracts in decimal"] = df["Tract"].apply(tract_to_decimal)

if "FFIEC_Geography" in df.columns:
    df.drop(columns=["FFIEC_Geography"], inplace=True)

df.insert(1, "FFIEC_Geography", df.apply(build_geography, axis=1))

df.to_excel(path, index=False)

print("Done. Sample output:")
print(df[["County name", "Tracts in decimal", "FFIEC_Geography"]].head(10).to_string(index=False))

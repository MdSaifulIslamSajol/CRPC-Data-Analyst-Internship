import pandas as pd

file_path = r"H:\stats america\FFIEC_gov\FFIEC_gov_CensusTractList2026_CRPC_Area_Only2.xlsx"

df = pd.read_excel(file_path)

county_names = sorted(df["County name"].dropna().unique())

print(f"County names ({len(county_names)} unique):")
for county in county_names:
    print(f"  {county}")

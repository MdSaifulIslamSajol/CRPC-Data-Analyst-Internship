import pandas as pd

src = r"H:\Business Information\County File\cbp23co\cbp23co_louisiana.csv"
out = r"H:\Business Information\County File\cbp23co\cbp23co_crpc.csv"

crpc_codes = ["005", "033", "037", "047", "063", "077", "091", "105", "117", "121", "125"]

df = pd.read_csv(src, dtype=str)

crpc = df[df["fipscty"].isin(crpc_codes)]

crpc.to_csv(out, index=False)

result_path = r"C:\Users\MSajol\AppData\Local\Temp\claude\h--stats-america-Business-data-analysis\96513bf2-524b-4d47-b640-8abe39f08050\scratchpad\filter_crpc_result.txt"
with open(result_path, "w") as f:
    f.write(f"Total Louisiana rows: {len(df)}\n")
    f.write(f"CRPC rows: {len(crpc)}\n")
    f.write(f"Unique CRPC county codes found: {sorted(crpc['fipscty'].unique())}\n")
    f.write(f"Saved to: {out}\n")

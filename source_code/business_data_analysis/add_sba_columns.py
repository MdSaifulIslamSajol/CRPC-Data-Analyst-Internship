import pandas as pd

src = r"H:\Business Information\County File\cbp23co\cbp23co_crpc.csv"
out = r"H:\Business Information\County File\cbp23co\cbp23co_crpc_sba.csv"

df = pd.read_csv(src, dtype=str)

size_cols = ["n<5", "n5_9", "n10_19", "n20_49", "n50_99",
             "n100_249", "n250_499", "n500_999", "n1000"]

# "N" (not available/not comparable) is treated as 0 for these summed counts
num = {c: pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int) for c in size_cols}

df["n_micro_enterprise"] = num["n<5"] + num["n5_9"]                                   # 0-9 employees
df["n_small_business"]   = num["n10_19"] + num["n20_49"] + num["n50_99"]              # 10-99 employees
df["n_medium_business"]  = num["n100_249"] + num["n250_499"]                          # 100-499 employees
df["n_large_employees"]  = num["n500_999"] + num["n1000"]                             # 500+ employees

df.to_csv(out, index=False)

result_path = r"C:\Users\MSajol\AppData\Local\Temp\claude\h--stats-america-Business-data-analysis\96513bf2-524b-4d47-b640-8abe39f08050\scratchpad\add_sba_result.txt"
with open(result_path, "w") as f:
    f.write(f"Rows: {len(df)}\n")
    f.write(f"Columns: {len(df.columns)}\n")
    f.write(f"Saved to: {out}\n")
    f.write("\nSample check (first 5 rows):\n")
    f.write(df[["fipscty", "naics", "est", "n_micro_enterprise", "n_small_business",
                "n_medium_business", "n_large_employees"]].head(5).to_string(index=False))

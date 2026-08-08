import pandas as pd

src = r"H:\Business Information\County File\cbp23co\cbp23co.txt"
out = r"H:\Business Information\County File\cbp23co\cbp23co_louisiana.csv"

df = pd.read_csv(src, dtype=str)

la = df[df["fipstate"] == "22"]

la.to_csv(out, index=False)

print(f"Total rows: {len(df)}")
print(f"Louisiana rows: {len(la)}")
print(f"Saved to: {out}")

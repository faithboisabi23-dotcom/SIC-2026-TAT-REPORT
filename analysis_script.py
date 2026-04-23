import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ── 1. Load ALL TOKENS ──────────────────────────────────────────────────────
try:
    all_df = pd.read_excel("TAT - ALL TOKENS.xlsx")
    print("ALL TOKENS columns:", list(all_df.columns))
    print("ALL TOKENS shape:", all_df.shape)
    print("ALL TOKENS dtypes:\n", all_df.dtypes)
    print("ALL TOKENS first 3 rows:\n", all_df.head(3).to_string())
except Exception as e:
    print(f"Error loading ALL TOKENS: {e}")

print("\n" + "="*50 + "\n")

# ── 2. Load COMPLETED TOKENS ────────────────────────────────────────────────
try:
    comp_df = pd.read_excel("TAT - ALL COMPLETED TOKENS.xlsx")
    print("COMPLETED TOKENS columns:", list(comp_df.columns))
    print("COMPLETED TOKENS shape:", comp_df.shape)
    print("COMPLETED TOKENS dtypes:\n", comp_df.dtypes)
    print("COMPLETED TOKENS first 3 rows:\n", comp_df.head(3).to_string())
except Exception as e:
    print(f"Error loading COMPLETED TOKENS: {e}")

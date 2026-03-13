import pandas as pd


def load_and_clean(filepath: str) -> pd.DataFrame:
    """
    Load a raw sales CSV and return a cleaned DataFrame.
    Handles common issues: missing values, duplicates,
    inconsistent formatting, wrong dtypes.
    """
    df = pd.read_csv(filepath,encoding="latin-1",on_bad_lines='skip')

    # ── 1. Normalise column names ──────────────────────────────────────────
    df.columns = (
        df.columns.str.strip()
                  .str.lower()
                  .str.replace(r"[\s\-/]+", "_", regex=True)
    )

    # ── 2. Drop fully empty rows ───────────────────────────────────────────
    df.dropna(how="all", inplace=True)

    # ── 3. Remove exact duplicates ─────────────────────────────────────────
    df.drop_duplicates(inplace=True)

    # ── 4. Standardise common column names (Superstore / Sample Sales) ─────
    rename_map = {
        "order_date":      "order_date",
        "orderdate":       "order_date",
        "order date":      "order_date",
        "sales":           "sales",
        "revenue":         "sales",
        "customer_name":   "customer_name",
        "customername":    "customer_name",
        "product_name":    "product_name",
        "productname":     "product_name",
        "region":          "region",
        "sub-category":    "sub_category",
        "sub_category":    "sub_category",
        "category":        "category",
        "quantity":        "quantity",
        "profit":          "profit",
        "discount":        "discount",
        "state":           "state",
        "city":            "city",
    }
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns},
              inplace=True)

    # ── 5. Parse order_date ────────────────────────────────────────────────
    if "order_date" in df.columns:
        df["order_date"] = pd.to_datetime(df["order_date"],errors="coerce")
        df["year"]  = df["order_date"].dt.year
        df["month"] = df["order_date"].dt.to_period("M").astype(str)

    # ── 6. Numeric columns – coerce and fill ──────────────────────────────
    for col in ["sales", "quantity", "profit", "discount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── 7. String columns – strip whitespace, title-case ──────────────────
    for col in ["customer_name", "product_name", "region",
                "category", "sub_category", "state", "city"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.title()

    # ── 8. Drop rows where sales is still zero/null ────────────────────────
    if "sales" in df.columns:
        df = df[df["sales"] > 0]

    df.reset_index(drop=True, inplace=True)
    return df


def cleaning_summary(raw_df: pd.DataFrame, clean_df: pd.DataFrame) -> dict:
    """Return a simple before/after summary for display."""
    return {
        "rows_before":   len(raw_df),
        "rows_after":    len(clean_df),
        "rows_removed":  len(raw_df) - len(clean_df),
        "columns_after": list(clean_df.columns),
    }
import pandas as pd


def monthly_revenue(df: pd.DataFrame) -> pd.DataFrame:
    """Revenue and profit aggregated by month."""
    if "month" not in df.columns or "sales" not in df.columns:
        return pd.DataFrame()

    cols = ["month", "sales"]
    if "profit" in df.columns:
        cols.append("profit")
    if "quantity" in df.columns:
        cols.append("quantity")

    monthly = (
        df[cols]
        .groupby("month")
        .sum()
        .reset_index()
        .sort_values("month")
    )

    monthly.rename(columns={
        "month":    "Month",
        "sales":    "Total Revenue (₹)",
        "profit":   "Total Profit (₹)",
        "quantity": "Units Sold",
    }, inplace=True)

    if "Total Profit (₹)" in monthly.columns and "Total Revenue (₹)" in monthly.columns:
        monthly["Profit Margin (%)"] = (
            monthly["Total Profit (₹)"] / monthly["Total Revenue (₹)"] * 100
        ).round(2)

    return monthly


def top_customers(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    """Top N customers by total revenue."""
    if "customer_name" not in df.columns or "sales" not in df.columns:
        return pd.DataFrame()

    top = (
        df.groupby("customer_name")["sales"]
        .sum()
        .reset_index()
        .sort_values("sales", ascending=False)
        .head(n)
        .reset_index(drop=True)
    )
    top.index += 1
    top.rename(columns={
        "customer_name": "Customer",
        "sales":         "Total Revenue (₹)",
    }, inplace=True)
    return top


def top_products(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    """Top N products by total revenue."""
    col = "product_name" if "product_name" in df.columns else (
          "sub_category"  if "sub_category"  in df.columns else None)

    if col is None or "sales" not in df.columns:
        return pd.DataFrame()

    top = (
        df.groupby(col)["sales"]
        .sum()
        .reset_index()
        .sort_values("sales", ascending=False)
        .head(n)
        .reset_index(drop=True)
    )
    top.index += 1
    top.rename(columns={col: "Product / Sub-Category", "sales": "Total Revenue (₹)"},
               inplace=True)
    return top


def region_performance(df: pd.DataFrame) -> pd.DataFrame:
    """Revenue, profit, and order count broken down by region."""
    if "region" not in df.columns or "sales" not in df.columns:
        return pd.DataFrame()

    agg: dict = {"sales": "sum"}
    if "profit" in df.columns:
        agg["profit"] = "sum"
    if "order_date" in df.columns:
        agg["order_date"] = "count"

    region = df.groupby("region").agg(agg).reset_index()
    region.rename(columns={
        "region":     "Region",
        "sales":      "Total Revenue (₹)",
        "profit":     "Total Profit (₹)",
        "order_date": "Order Count",
    }, inplace=True)
    region.sort_values("Total Revenue (₹)", ascending=False, inplace=True)
    region.reset_index(drop=True, inplace=True)
    region.index += 1
    return region


def kpi_summary(df: pd.DataFrame) -> dict:
    """High-level KPIs for the executive summary sheet."""
    kpis: dict = {}
    if "sales" in df.columns:
        kpis["Total Revenue (₹)"]    = round(df["sales"].sum(), 2)
        kpis["Average Order Value (₹)"] = round(df["sales"].mean(), 2)
    if "profit" in df.columns:
        kpis["Total Profit (₹)"]     = round(df["profit"].sum(), 2)
        kpis["Overall Profit Margin (%)"] = round(
            df["profit"].sum() / df["sales"].sum() * 100, 2
        ) if df["sales"].sum() else 0
    if "order_date" in df.columns:
        kpis["Total Orders"]         = len(df)
    if "customer_name" in df.columns:
        kpis["Unique Customers"]     = df["customer_name"].nunique()
    if "region" in df.columns:
        kpis["Regions Covered"]      = df["region"].nunique()
    return kpis
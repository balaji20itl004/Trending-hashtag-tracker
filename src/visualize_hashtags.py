#!/usr/bin/env python3
"""
Read an Excel file with columns [Hashtag, Count], create bar & pie charts,
and export cleaned CSV + PNGs for your GitHub repo / report.

Usage:
    python src/visualize_hashtags.py --input data/top_hashtags.xlsx --sheet "Top Hashtags" --outdir outputs
"""

import argparse
import os
import pandas as pd
import matplotlib.pyplot as plt

def find_columns(df):
    """Allow flexible column naming like 'hashtag' or 'Hashtag', 'count' or 'Count'."""
    cols_lower = {c.lower(): c for c in df.columns}
    hashtag_col = cols_lower.get("hashtag")
    count_col = cols_lower.get("count")
    if not hashtag_col or not count_col:
        raise ValueError("Input sheet must have columns 'Hashtag' and 'Count'. Found: "
                         f"{list(df.columns)}")
    return hashtag_col, count_col

def main(input_path, sheet_name, outdir):
    os.makedirs(outdir, exist_ok=True)

    # 1) Read Excel
    df = pd.read_excel(input_path, sheet_name=sheet_name if sheet_name else 0)

    # 2) Validate/clean
    hashtag_col, count_col = find_columns(df)
    df = df[[hashtag_col, count_col]].copy()
    df.columns = ["Hashtag", "Count"]
    df["Hashtag"] = df["Hashtag"].astype(str).str.strip()
    df["Count"] = pd.to_numeric(df["Count"], errors="coerce").fillna(0).astype(int)

    # Drop empty hashtags
    df = df[df["Hashtag"].str.len() > 0]

    # 3) Sort by count desc
    df = df.sort_values("Count", ascending=False).reset_index(drop=True)

    # 4) Save a cleaned CSV for the repo
    cleaned_csv = os.path.join(outdir, "top_hashtags_cleaned.csv")
    df.to_csv(cleaned_csv, index=False)

    # 5) BAR CHART
    plt.figure()
    plt.bar(df["Hashtag"], df["Count"])
    plt.xticks(rotation=45, ha="right")
    plt.title("Top Trending Hashtags (Bar Chart)")
    plt.tight_layout()
    bar_path = os.path.join(outdir, "top_hashtags_bar.png")
    plt.savefig(bar_path, dpi=200)
    plt.close()

    # 6) PIE CHART
    plt.figure()
    plt.pie(df["Count"], labels=df["Hashtag"], autopct="%1.1f%%")
    plt.title("Top Trending Hashtags (Pie Chart)")
    plt.tight_layout()
    pie_path = os.path.join(outdir, "top_hashtags_pie.png")
    plt.savefig(pie_path, dpi=200)
    plt.close()

    # 7) Small Markdown report
    md_path = os.path.join(outdir, "REPORT.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# Trending Hashtags – Auto Report\n\n")
        f.write("## Top 10 (preview)\n\n")
        f.write(df.head(10).to_markdown(index=False))
        f.write("\n\n## Charts\n\n")
        f.write(f"![Bar Chart](./top_hashtags_bar.png)\n\n")
        f.write(f"![Pie Chart](./top_hashtags_pie.png)\n")

    print("✅ Done:")
    print(f"- Cleaned CSV: {cleaned_csv}")
    print(f"- Bar chart:   {bar_path}")
    print(f"- Pie chart:   {pie_path}")
    print(f"- Report MD:   {md_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to Excel file")
    parser.add_argument("--sheet", default=None, help="Worksheet name (optional)")
    parser.add_argument("--outdir", default="outputs", help="Output directory")
    args = parser.parse_args()
    main(args.input, args.sheet, args.outdir)

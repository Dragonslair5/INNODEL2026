"""
Bibliometric Analysis Script
=============================
Analyzes a Scopus/WoS-style Excel export (Portuguese column headers) and produces:

1) Number of publications per year
2) Top journals publishing (by paper count)
3) Top journals getting cited (by total citations)
4) Most cited papers
5) Average citations per year (simple + normalized by years since publication)
6) Most frequent keywords (+ word cloud)
7) Keyword co-occurrence network
8) Authorship co-occurrence network
9) Lotka's Law check (author productivity distribution)
Plus: dataset summary, dual-metric journal charts, scatter plot with zoom panel,
LaTeX table/figure helpers.

Requirements:
    pip install pandas openpyxl matplotlib networkx pyvis wordcloud scipy

Usage:
    python bibliometric_analysis_full.py your_file.xlsx --sheet "Sheet1"

All outputs are saved to ./output/
"""

import argparse
import datetime
import os
from collections import Counter
from itertools import combinations

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import networkx as nx
from scipy import stats

# ----------------------------------------------------------------------
# CONFIG - adjust these if your column names differ slightly
# ----------------------------------------------------------------------
COL_TITLE = "Titulo"
COL_YEAR = "Ano"
COL_JOURNAL = "Periodico"
COL_AUTHORS = "Autores"
COL_CIT_SCOPUS = "Citacoes_Scopus"
COL_CIT_WOS = "Citacoes_WoS"
COL_KW_AUTHOR = "Palavras_chave_autor"
COL_KW_INDEX = "Palavras_chave_indexadas"

TOP_N = 15                 # how many items to show in "top N" charts/tables
KW_MIN_COOCCUR = 2         # minimum co-occurrence count to keep a keyword edge
AUTHOR_MIN_COOCCUR = 2     # minimum co-occurrence count to keep an author edge

OUTPUT_DIR = "output"


# ----------------------------------------------------------------------
# GENERIC HELPERS
# ----------------------------------------------------------------------
def split_multi(value, seps=(";", "|", ",")):
    """Split a semicolon/pipe/comma separated string into a clean list."""
    if pd.isna(value) or str(value).strip() == "":
        return []
    text = str(value)
    sep = ";"
    for s in seps:
        if s in text:
            sep = s
            break
    items = [x.strip() for x in text.split(sep)]
    return [x for x in items if x]


def get_citations(row):
    """Prefer Scopus citations, fall back to WoS if missing."""
    sc = row.get(COL_CIT_SCOPUS)
    wos = row.get(COL_CIT_WOS)
    if pd.notna(sc):
        return float(sc)
    if pd.notna(wos):
        return float(wos)
    return 0.0


def ensure_outdir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def save_bar_chart(series, title, xlabel, ylabel, filename, horizontal=False, fmt="%.0f"):
    fig, ax = plt.subplots(figsize=(10, 6))
    if horizontal:
        series = series.sort_values()
        bars = ax.barh(series.index.astype(str), series.values, color="#4C72B0")
    else:
        bars = ax.bar(series.index.astype(str), series.values, color="#4C72B0")

    ax.bar_label(bars, padding=3, fontsize=9, fmt=fmt)

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)

    if not horizontal:
        plt.xticks(rotation=45, ha="right")

    if horizontal:
        ax.set_xlim(0, series.max() * 1.15)
    else:
        ax.set_ylim(0, series.max() * 1.15)

    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(path, dpi=150)
    plt.close()
    print(f"Saved chart: {path}")


def escape_latex(text):
    """Escape LaTeX special characters in a string."""
    if pd.isna(text):
        return ""
    text = str(text)
    replacements = {
        "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
        "_": r"\_", "{": r"\{", "}": r"\}",
        "~": r"\textasciitilde{}", "^": r"\textasciicircum{}",
        "\\": r"\textbackslash{}",
    }
    for char, repl in replacements.items():
        text = text.replace(char, repl)
    return text


def save_latex_table(df, total_n, filename, caption_prefix="Most cited papers",
                      label="tab:most_cited", title_col=COL_TITLE, year_col=COL_YEAR,
                      citations_col="citations", max_title_len=80):
    """Save a LaTeX table with Title, Year, and Citations columns.
    Caption automatically includes 'out of N total' based on total_n."""
    top_n = df.shape[0]
    caption = f"{caption_prefix} (top {top_n} out of {total_n} total)"

    lines = []
    lines.append(r"\begin{table}[htbp]")
    lines.append(r"\centering")
    lines.append(r"\caption{" + escape_latex(caption) + "}")
    lines.append(r"\label{" + label + "}")
    lines.append(r"\begin{tabular}{p{9cm}cc}")
    lines.append(r"\toprule")
    lines.append(r"Title & Year & Citations \\")
    lines.append(r"\midrule")

    for i, (_, row) in enumerate(df.iterrows()):
        title = str(row[title_col])
        if max_title_len and len(title) > max_title_len:
            title = title[:max_title_len].rstrip() + "..."
        title = escape_latex(title)

        year = row[year_col]
        year = "" if pd.isna(year) else str(int(year))

        citations = row[citations_col]
        citations = "" if pd.isna(citations) else str(int(citations))

        color = "white" if i % 2 == 0 else "gray!15"
        lines.append(r"\rowcolor{" + color + "}")
        lines.append(f"{title} & {year} & {citations} \\\\")

    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\end{table}")

    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Saved LaTeX table: {path}")


def save_series_latex_table(series, filename, caption, label,
                              col1_name="Year", col2_name="Value",
                              value_fmt="{:.2f}", alternate_colors=True,
                              color1="white", color2="gray!15"):
    """Save a simple two-column LaTeX table (e.g. Year vs. a metric) from a pandas Series."""
    lines = []
    lines.append(r"\begin{table}[htbp]")
    lines.append(r"\centering")
    lines.append(r"\caption{" + escape_latex(caption) + "}")
    lines.append(r"\label{" + label + "}")
    lines.append(r"\begin{tabular}{cc}")
    lines.append(r"\toprule")
    lines.append(f"{col1_name} & {col2_name} \\\\")
    lines.append(r"\midrule")

    for i, (idx, val) in enumerate(series.items()):
        if alternate_colors:
            color = color1 if i % 2 == 0 else color2
            lines.append(r"\rowcolor{" + color + "}")
        lines.append(f"{idx} & {value_fmt.format(val)} \\\\")

    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\end{table}")

    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Saved LaTeX table: {path}")


# ----------------------------------------------------------------------
# DATA COVERAGE TABLE
# ----------------------------------------------------------------------
def is_filled(val):
    """True if a cell counts as 'present' (not NaN, not empty/whitespace string)."""
    if pd.isna(val):
        return False
    if isinstance(val, str) and val.strip() == "":
        return False
    return True


def field_coverage_pct(df, col):
    """% of rows where col is present (non-missing). If col is a boolean column
    (e.g. a derived flag like '_has_any_citations'), returns % of True values instead."""
    if col not in df.columns:
        return None
    total = df.shape[0]
    if total == 0:
        return 0.0
    if pd.api.types.is_bool_dtype(df[col]):
        return 100 * df[col].sum() / total
    filled = df[col].apply(is_filled).sum()
    return 100 * filled / total


def save_coverage_table(df, groups, filename="data_coverage.tex",
                          title=None, caption="Data coverage by field",
                          label="tab:data_coverage",
                          highlight_threshold=90.0, decimals=1):
    """Build a two-category side-by-side LaTeX coverage table:

        [optional bold title row spanning all columns]
        Category A           %      Category B           %
        ------------------------------------------------------
        Field 1               99.6   Field 1               100.0
        Field 2               14.7   Field 2               100.0
        ...

    groups: dict of {category_label: [(display_name, column_name), ...]}
            exactly two categories expected (left and right columns of the table).
    title: optional bold heading printed as a row inside the table, above the
           category headers (distinct from the LaTeX \\caption, which appears
           in the List of Tables). Omit for no title row.
    highlight_threshold: fields with coverage >= this value get \\cellcolor + \\textbf.
    """
    if len(groups) != 2:
        raise ValueError("save_coverage_table expects exactly two groups (left and right).")

    (cat1_label, cat1_fields), (cat2_label, cat2_fields) = list(groups.items())

    def compute(fields):
        rows = []
        for display_name, col in fields:
            pct = field_coverage_pct(df, col)
            rows.append((display_name, pct))
        return rows

    rows1 = compute(cat1_fields)
    rows2 = compute(cat2_fields)
    n_rows = max(len(rows1), len(rows2))

    def fmt_cell(pct):
        if pct is None:
            return "--"
        text = f"{pct:.{decimals}f}"
        if pct >= highlight_threshold:
            return r"\cellcolor{gray!25}\textbf{" + text + "}"
        return text

    lines = []
    lines.append(r"\begin{table}[htbp]")
    lines.append(r"\centering")
    lines.append(r"\caption{" + escape_latex(f"{caption} (n = {df.shape[0]} papers)") + "}")
    lines.append(r"\label{" + label + "}")
    lines.append(r"\begin{tabular}{@{}lrlr@{}}")
    lines.append(r"\toprule")
    if title:
        lines.append(r"\multicolumn{4}{c}{\textbf{" + escape_latex(title) + r"}} \\")
        lines.append(r"\midrule")
    lines.append(
        r"\textbf{\makecell[l]{" + escape_latex(cat1_label) + r"}} & \textbf{\makecell[r]{\%}} & "
        r"\textbf{\makecell[l]{" + escape_latex(cat2_label) + r"}} & \textbf{\makecell[r]{\%}} \\"
    )
    lines.append(r"\midrule")

    for i in range(n_rows):
        name1, pct1 = rows1[i] if i < len(rows1) else ("", None)
        name2, pct2 = rows2[i] if i < len(rows2) else ("", None)

        c1 = escape_latex(name1) if name1 else ""
        c2 = fmt_cell(pct1) if name1 else ""
        c3 = escape_latex(name2) if name2 else ""
        c4 = fmt_cell(pct2) if name2 else ""

        lines.append(f"{c1} & {c2} & {c3} & {c4} \\\\")

    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\end{table}")

    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Saved coverage table: {path}")

    # Also export a flat CSV version for easy reuse/inspection
    all_rows = [(cat1_label, n, p) for n, p in rows1] + [(cat2_label, n, p) for n, p in rows2]
    cov_df = pd.DataFrame(all_rows, columns=["category", "field", "pct_coverage"])
    csv_path = os.path.join(OUTPUT_DIR, filename.replace(".tex", ".csv"))
    cov_df.to_csv(csv_path, index=False)
    print(f"Saved coverage CSV: {csv_path}")

    return cov_df


# ----------------------------------------------------------------------
# DATASET SUMMARY
# ----------------------------------------------------------------------
def print_dataset_summary(df):
    """Print basic descriptive stats about the dataset before running any analysis."""
    n_papers = df.shape[0]

    n_journals = df[COL_JOURNAL].dropna().nunique() if COL_JOURNAL in df.columns else None

    n_authors = None
    if COL_AUTHORS in df.columns:
        all_authors = set()
        for val in df[COL_AUTHORS]:
            all_authors.update(a.strip() for a in split_multi(val) if a.strip())
        n_authors = len(all_authors)

    year_min, year_max = None, None
    if COL_YEAR in df.columns:
        years = df[COL_YEAR].dropna()
        if not years.empty:
            year_min, year_max = int(years.min()), int(years.max())

    print("=" * 50)
    print("DATASET SUMMARY")
    print("=" * 50)
    print(f"Total papers:            {n_papers}")
    print(f"Distinct journals:       {n_journals if n_journals is not None else 'N/A'}")
    print(f"Distinct authors:        {n_authors if n_authors is not None else 'N/A'}")
    if year_min is not None:
        print(f"Year range:              {year_min}\u2013{year_max} ({year_max - year_min + 1} years)")
    else:
        print("Year range:              N/A")
    print("=" * 50)
    print()


# ----------------------------------------------------------------------
# 1) PUBLICATIONS PER YEAR
# ----------------------------------------------------------------------
def publications_per_year(df):
    counts = df[COL_YEAR].dropna().astype(int).value_counts().sort_index()
    counts.to_csv(os.path.join(OUTPUT_DIR, "pubs_per_year.csv"), header=["n_publications"])
    save_bar_chart(counts, "Publications per Year", "Year", "Number of Publications",
                    "pubs_per_year.svg")
    return counts


# ----------------------------------------------------------------------
# 2) TOP JOURNALS PUBLISHING
# ----------------------------------------------------------------------
def top_journals_publishing(df, top_n=TOP_N):
    counts_all = df[COL_JOURNAL].dropna().value_counts()
    total_journals = counts_all.shape[0]
    counts = counts_all.head(top_n)

    counts.to_csv(os.path.join(OUTPUT_DIR, "top_journals_publishing.csv"), header=["n_publications"])

    title = f"Top {top_n} Journals by Publication Count (out of {total_journals} journals)"
    save_bar_chart(counts, title, "Number of Publications", "Journal",
                    "top_journals_publishing.svg", horizontal=True)
    return counts


# ----------------------------------------------------------------------
# 3) TOP JOURNALS CITED
# ----------------------------------------------------------------------
def top_journals_cited(df, top_n=TOP_N):
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)
    grouped_all = tmp.groupby(COL_JOURNAL)["citations"].sum().sort_values(ascending=False)

    total_journals = grouped_all.shape[0]
    grouped = grouped_all.head(top_n)

    grouped.to_csv(os.path.join(OUTPUT_DIR, "top_journals_cited.csv"), header=["total_citations"])

    title = f"Top {top_n} Journals by Total Citations (out of {total_journals} journals)"
    save_bar_chart(grouped, title, "Total Citations", "Journal", "top_journals_cited.svg",
                    horizontal=True)
    return grouped


# ----------------------------------------------------------------------
# 2b/3b) DUAL-METRIC JOURNAL CHARTS (twin x-axis)
# ----------------------------------------------------------------------
def top_journals_dual_chart(df, top_n=30, filename="top_journals_dual.svg"):
    """Select the top N journals by publication count, order by publications,
    and plot both metrics with a twin x-axis (separate scales)."""
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)

    pub_counts_all = tmp[COL_JOURNAL].dropna().value_counts()
    total_journals = pub_counts_all.shape[0]
    top_journals = pub_counts_all.head(top_n).index

    subset = tmp[tmp[COL_JOURNAL].isin(top_journals)]
    citation_totals = subset.groupby(COL_JOURNAL)["citations"].sum()
    pub_totals = subset.groupby(COL_JOURNAL)[COL_JOURNAL].count()

    combined = pd.DataFrame({
        "publications": pub_totals,
        "citations": citation_totals,
    }).sort_values("publications", ascending=True)

    fig, ax_pub = plt.subplots(figsize=(12, max(8, top_n * 0.35)))
    ax_cit = ax_pub.twiny()

    y = range(len(combined))
    bar_height = 0.38

    bars_pub = ax_pub.barh([p + bar_height / 2 for p in y], combined["publications"],
                            height=bar_height, label="Publications", color="#4C72B0")
    bars_cit = ax_cit.barh([p - bar_height / 2 for p in y], combined["citations"],
                            height=bar_height, label="Citations", color="#C44E52")

    ax_pub.bar_label(bars_pub, padding=3, fontsize=7, fmt="%.0f", color="#4C72B0")
    ax_cit.bar_label(bars_cit, padding=3, fontsize=7, fmt="%.0f", color="#C44E52")

    ax_pub.set_yticks(list(y))
    ax_pub.set_yticklabels(combined.index)

    ax_pub.set_xlabel("Number of Publications", color="#4C72B0")
    ax_cit.set_xlabel("Number of Citations", color="#C44E52")
    ax_pub.tick_params(axis="x", labelcolor="#4C72B0")
    ax_cit.tick_params(axis="x", labelcolor="#C44E52")

    ax_pub.set_xlim(0, combined["publications"].max() * 1.2)
    ax_cit.set_xlim(0, combined["citations"].max() * 1.2)

    ax_pub.set_title(f"Top {top_n} Journals by Publication Count "
                      f"(out of {total_journals} journals)")

    lines_labels = [ax_pub.get_legend_handles_labels(), ax_cit.get_legend_handles_labels()]
    handles = [h for hl in lines_labels for h in hl[0]]
    labels = [l for hl in lines_labels for l in hl[1]]
    ax_pub.legend(handles, labels, loc="lower right")

    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(path, dpi=150)
    plt.close()
    print(f"Saved dual-axis chart: {path}")

    combined.sort_values("publications", ascending=False).to_csv(
        os.path.join(OUTPUT_DIR, "top_journals_dual.csv"))

    return combined


def top_cited_journals_dual_chart(df, top_n=30, filename="top_cited_journals_dual.svg"):
    """Select the top N journals by total citations, order by citations,
    show publication counts alongside on a twin x-axis, and include the
    citations-per-paper ratio in the journal labels."""
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)

    citation_totals_all = tmp.groupby(COL_JOURNAL)["citations"].sum().sort_values(ascending=False)
    total_journals = citation_totals_all.shape[0]
    top_journals = citation_totals_all.head(top_n).index

    subset = tmp[tmp[COL_JOURNAL].isin(top_journals)]
    citation_totals = subset.groupby(COL_JOURNAL)["citations"].sum()
    pub_totals = subset.groupby(COL_JOURNAL)[COL_JOURNAL].count()

    combined = pd.DataFrame({
        "publications": pub_totals,
        "citations": citation_totals,
    }).sort_values("citations", ascending=True)
    combined["citations_per_paper"] = combined["citations"] / combined["publications"]

    y_labels = [f"{journal}  ({row['citations_per_paper']:.1f} cit/paper)"
                for journal, row in combined.iterrows()]

    fig, ax_cit = plt.subplots(figsize=(13, max(8, top_n * 0.35)))
    ax_pub = ax_cit.twiny()

    y = range(len(combined))
    bar_height = 0.38

    bars_cit = ax_cit.barh([p + bar_height / 2 for p in y], combined["citations"],
                            height=bar_height, label="Citations", color="#C44E52")
    bars_pub = ax_pub.barh([p - bar_height / 2 for p in y], combined["publications"],
                            height=bar_height, label="Publications", color="#4C72B0")

    ax_cit.bar_label(bars_cit, padding=3, fontsize=7, fmt="%.0f", color="#C44E52")
    ax_pub.bar_label(bars_pub, padding=3, fontsize=7, fmt="%.0f", color="#4C72B0")

    ax_cit.set_yticks(list(y))
    ax_cit.set_yticklabels(y_labels)

    ax_cit.set_xlabel("Number of Citations", color="#C44E52")
    ax_pub.set_xlabel("Number of Publications", color="#4C72B0")
    ax_cit.tick_params(axis="x", labelcolor="#C44E52")
    ax_pub.tick_params(axis="x", labelcolor="#4C72B0")

    ax_cit.set_xlim(0, combined["citations"].max() * 1.2)
    ax_pub.set_xlim(0, combined["publications"].max() * 1.2)

    ax_cit.set_title(f"Top {top_n} Journals by Citations "
                      f"(out of {total_journals} journals), with Publication Counts "
                      f"and Citations/Paper Ratio")

    lines_labels = [ax_cit.get_legend_handles_labels(), ax_pub.get_legend_handles_labels()]
    handles = [h for hl in lines_labels for h in hl[0]]
    labels = [l for hl in lines_labels for l in hl[1]]
    ax_cit.legend(handles, labels, loc="lower right")

    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(path, dpi=150)
    plt.close()
    print(f"Saved dual-axis chart: {path}")

    combined.sort_values("citations", ascending=False).to_csv(
        os.path.join(OUTPUT_DIR, "top_cited_journals_dual.csv"))

    return combined


def top_journals_by_citation_ratio(df, top_n=15, min_pubs=3, filename="top_journals_ratio.svg"):
    """Rank journals by average citations per paper, filtering out journals with
    very few papers to avoid small-sample noise."""
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)

    pub_totals = tmp[COL_JOURNAL].dropna().value_counts()
    citation_totals = tmp.groupby(COL_JOURNAL)["citations"].sum()

    combined = pd.DataFrame({
        "publications": pub_totals,
        "citations": citation_totals,
    }).dropna()
    combined = combined[combined["publications"] >= min_pubs]
    combined["citations_per_paper"] = combined["citations"] / combined["publications"]

    top = combined.sort_values("citations_per_paper", ascending=False).head(top_n)

    save_bar_chart(top["citations_per_paper"],
                    f"Top {top_n} Journals by Citations per Paper (min. {min_pubs} papers)",
                    "Avg. Citations per Paper", "Journal", filename,
                    horizontal=True, fmt="%.1f")

    top.to_csv(os.path.join(OUTPUT_DIR, "top_journals_ratio.csv"))
    return top


def _add_labels(ax, labeled_df, x_col, y_col):
    """Add journal-name labels to points, auto-positioned to avoid overlap and
    stay within the axes bounds (via adjustText if installed; falls back to a
    simple clipped offset otherwise)."""
    try:
        from adjustText import adjust_text
        texts = [ax.text(row[x_col], row[y_col], journal, fontsize=8)
                 for journal, row in labeled_df.iterrows()]
        adjust_text(texts, ax=ax,
                    expand_points=(1.3, 1.3),
                    arrowprops=dict(arrowstyle="-", color="gray", lw=0.5))
    except ImportError:
        print("[i] Install 'adjustText' (pip install adjustText) for better label "
              "placement. Falling back to simple clipped offsets.")
        for journal, row in labeled_df.iterrows():
            ax.annotate(journal, (row[x_col], row[y_col]),
                        fontsize=8, xytext=(5, 5), textcoords="offset points",
                        clip_on=True)


def journals_scatter_pub_vs_citations(df, filename="journals_scatter.svg",
                                        label_top_n=10, min_pubs=1,
                                        zoom_max_pubs=None, zoom_max_citations=None,
                                        zoom_label_top_n=10, stacked_layout=False,
                                        axis_pad_pct=0.03):
    """Scatter plot: one dot per journal, x=publications, y=citations.
    Labels the top N journals by citations. If zoom_max_pubs/zoom_max_citations
    are given, adds a second panel zoomed into that region, and draws a rectangle
    on the main plot marking the zoomed region.

    stacked_layout=True arranges panels vertically (tall figure) - use this if the
    final figure will be placed with \\sidewaysfigure in LaTeX (rotation makes a
    tall image render wide on the page). stacked_layout=False (default) places
    panels side by side - use this for a normal upright figure placement.

    axis_pad_pct: fraction of the axis max used as padding below zero, so points
    sitting exactly at 0 citations/publications aren't visually clipped by the
    axis line/spine.
    """
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)

    pub_totals = tmp[COL_JOURNAL].dropna().value_counts()
    citation_totals = tmp.groupby(COL_JOURNAL)["citations"].sum()

    combined = pd.DataFrame({
        "publications": pub_totals,
        "citations": citation_totals,
    }).dropna()
    combined = combined[combined["publications"] >= min_pubs]

    use_zoom = zoom_max_pubs is not None and zoom_max_citations is not None

    if use_zoom:
        if stacked_layout:
            fig, (ax_main, ax_zoom) = plt.subplots(2, 1, figsize=(8, 12))
        else:
            fig, (ax_main, ax_zoom) = plt.subplots(1, 2, figsize=(18, 8))
    else:
        fig, ax_main = plt.subplots(figsize=(10, 8))

    ax_main.scatter(combined["publications"], combined["citations"],
                     alpha=0.6, color="#4C72B0", edgecolors="black", linewidths=0.5)

    top_labeled = combined.sort_values("citations", ascending=False).head(label_top_n)

    x_pad = combined["publications"].max() * axis_pad_pct
    y_pad = combined["citations"].max() * axis_pad_pct
    ax_main.set_xlim(-x_pad, combined["publications"].max() * 1.15)
    ax_main.set_ylim(-y_pad, combined["citations"].max() * 1.15)

    _add_labels(ax_main, top_labeled, "publications", "citations")

    med_pub = combined["publications"].median()
    med_cit = combined["citations"].median()
    ax_main.axvline(med_pub, color="gray", linestyle="--", linewidth=0.8, alpha=0.6)
    ax_main.axhline(med_cit, color="gray", linestyle="--", linewidth=0.8, alpha=0.6)

    ax_main.set_xlabel("Number of Publications")
    ax_main.set_ylabel("Number of Citations")
    ax_main.set_title(f"All Journals ({combined.shape[0]} total)\ndashed lines = medians")

    if use_zoom:
        from matplotlib.patches import Rectangle
        rect = Rectangle((0, 0), zoom_max_pubs, zoom_max_citations,
                          linewidth=1.5, edgecolor="#C44E52", facecolor="none",
                          linestyle="-", zorder=5)
        ax_main.add_patch(rect)

        zoom_subset = combined[(combined["publications"] <= zoom_max_pubs) &
                                (combined["citations"] <= zoom_max_citations)]
        n_in_zoom = zoom_subset.shape[0]
        n_excluded = combined.shape[0] - n_in_zoom

        ax_zoom.scatter(zoom_subset["publications"], zoom_subset["citations"],
                         alpha=0.6, color="#4C72B0", edgecolors="black", linewidths=0.5)

        zoom_x_pad = zoom_max_pubs * axis_pad_pct
        zoom_y_pad = zoom_max_citations * axis_pad_pct
        ax_zoom.set_xlim(-zoom_x_pad, zoom_max_pubs)
        ax_zoom.set_ylim(-zoom_y_pad, zoom_max_citations)

        zoom_labeled = zoom_subset.sort_values("citations", ascending=False).head(zoom_label_top_n)
        _add_labels(ax_zoom, zoom_labeled, "publications", "citations")

        ax_zoom.axvline(med_pub, color="gray", linestyle="--", linewidth=0.8, alpha=0.6)
        ax_zoom.axhline(med_cit, color="gray", linestyle="--", linewidth=0.8, alpha=0.6)
        ax_zoom.set_xlabel("Number of Publications")
        ax_zoom.set_ylabel("Number of Citations")
        ax_zoom.set_title(f"Zoomed: publications \u2264 {zoom_max_pubs}, "
                           f"citations \u2264 {zoom_max_citations}\n"
                           f"({n_in_zoom} journals shown, {n_excluded} outliers excluded)")

        for spine in ax_zoom.spines.values():
            spine.set_edgecolor("#C44E52")
            spine.set_linewidth(1.5)

    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(path, dpi=150)
    plt.close()
    print(f"Saved scatter chart: {path}")

    combined.to_csv(os.path.join(OUTPUT_DIR, "journals_scatter_data.csv"))
    return combined


# ----------------------------------------------------------------------
# 4) MOST CITED PAPERS
# ----------------------------------------------------------------------
def most_cited_papers(df, top_n=TOP_N):
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)

    total_papers = tmp.shape[0]

    cols = [COL_TITLE, COL_YEAR, COL_JOURNAL, COL_AUTHORS, "citations"]
    cols = [c for c in cols if c in tmp.columns]
    top = tmp.sort_values("citations", ascending=False).head(top_n)[cols]

    csv_path = os.path.join(OUTPUT_DIR, "most_cited_papers.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write(f"# Top {top_n} most cited papers (out of {total_papers} papers total)\n")
        top.to_csv(f, index=False)
    print(f"Saved: {csv_path}  (top {top_n} of {total_papers} papers)")

    save_latex_table(
        top,
        total_n=total_papers,
        filename="most_cited_papers.tex",
        caption_prefix="Most cited papers",
        label="tab:most_cited",
    )

    return top


def top_cited_papers_with_journal_context(df, top_n=15, filename_stem="most_cited_papers_journal_context",
                                            max_title_len=None, max_journal_len=None):
    """Top N cited papers, each annotated with:
    - the journal's rank among all journals by total citations
    - the journal's overall citations-per-paper ratio
    Produces CSV + a LaTeX table (wrapped in \\resizebox to fit the page width).
    Title/Journal columns use p{} widths, which wrap long text onto multiple lines
    automatically - so max_title_len/max_journal_len default to None (no truncation).
    Set them to an integer only if you want hard truncation instead of wrapping."""
    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)
    total_papers = tmp.shape[0]

    # Journal-level stats: total citations, total publications, rank by citations, ratio
    journal_citations = tmp.groupby(COL_JOURNAL)["citations"].sum()
    journal_pubs = tmp.groupby(COL_JOURNAL)[COL_JOURNAL].count()
    journal_rank = journal_citations.rank(method="min", ascending=False).astype(int)
    journal_ratio = journal_citations / journal_pubs

    # Top N papers by citations
    cols = [COL_TITLE, COL_JOURNAL, "citations"]
    cols = [c for c in cols if c in tmp.columns]
    top = tmp.sort_values("citations", ascending=False).head(top_n)[cols].copy()

    top["journal_rank_by_citations"] = top[COL_JOURNAL].map(journal_rank)
    top["journal_citations_per_paper"] = top[COL_JOURNAL].map(journal_ratio)

    csv_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.csv")
    top.to_csv(csv_path, index=False)
    print(f"Saved: {csv_path}")

    # --- LaTeX table ---
    lines = []
    lines.append(r"\begin{table}[htbp]")
    lines.append(r"\centering")
    caption = (f"Top {top_n} most cited papers (out of {total_papers} total), with their "
               f"journal's rank by total citations and citations-per-paper ratio")
    lines.append(r"\caption{" + escape_latex(caption) + "}")
    lines.append(r"\label{tab:" + filename_stem + "}")
    lines.append(r"\resizebox{\textwidth}{!}{%")
    lines.append(r"\begin{tabular}{@{}p{7cm}p{4cm}rrr@{}}")
    lines.append(r"\toprule")
    lines.append(r"Title & Journal & Citations & Journal Rank & Journal Cit./Paper \\")
    lines.append(r"\midrule")

    for i, (_, row) in enumerate(top.iterrows()):
        title = str(row[COL_TITLE])
        if max_title_len and len(title) > max_title_len:
            title = title[:max_title_len].rstrip() + "..."
        title = escape_latex(title)

        journal = str(row[COL_JOURNAL])
        if max_journal_len and len(journal) > max_journal_len:
            journal = journal[:max_journal_len].rstrip() + "..."
        journal = escape_latex(journal)

        citations = int(row["citations"]) if pd.notna(row["citations"]) else 0
        rank = int(row["journal_rank_by_citations"]) if pd.notna(row["journal_rank_by_citations"]) else "--"
        ratio = f"{row['journal_citations_per_paper']:.1f}" if pd.notna(row["journal_citations_per_paper"]) else "--"

        color = "white" if i % 2 == 0 else "gray!15"
        lines.append(r"\rowcolor{" + color + "}")
        lines.append(f"{title} & {journal} & {citations} & {rank} & {ratio} \\\\")

    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"}")  # close resizebox
    lines.append(r"\end{table}")

    tex_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Saved LaTeX table: {tex_path}")

    return top


# ----------------------------------------------------------------------
# 5) AVERAGE CITATIONS PER YEAR
# ----------------------------------------------------------------------
def avg_citations_per_year(df, current_year=None):
    """Average citations per paper, grouped by publication year.
    Also computes citations/(years since publication) as a normalized metric."""
    if current_year is None:
        current_year = datetime.datetime.now().year

    tmp = df.copy()
    tmp["citations"] = tmp.apply(get_citations, axis=1)
    tmp = tmp.dropna(subset=[COL_YEAR])
    tmp[COL_YEAR] = tmp[COL_YEAR].astype(int)

    total_papers = tmp.shape[0]

    avg_by_year = tmp.groupby(COL_YEAR)["citations"].mean().sort_index()
    avg_by_year.to_csv(os.path.join(OUTPUT_DIR, "avg_citations_per_year.csv"),
                        header=["avg_citations"])
    save_bar_chart(avg_by_year, "Average Citations per Paper, by Publication Year",
                    "Year", "Average Citations", "avg_citations_per_year.svg",
                    fmt="%.2f")
    save_series_latex_table(
        avg_by_year,
        filename="avg_citations_per_year.tex",
        caption=f"Average citations per paper by publication year (based on {total_papers} papers)",
        label="tab:avg_citations_per_year",
        col1_name="Year", col2_name="Avg. Citations",
        value_fmt="{:.2f}",
    )

    tmp["years_since_pub"] = (current_year - tmp[COL_YEAR]).clip(lower=1)
    tmp["citations_per_year"] = tmp["citations"] / tmp["years_since_pub"]
    norm_by_year = tmp.groupby(COL_YEAR)["citations_per_year"].mean().sort_index()
    norm_by_year.to_csv(os.path.join(OUTPUT_DIR, "avg_citations_per_year_normalized.csv"),
                         header=["avg_citations_per_year_since_pub"])
    save_bar_chart(norm_by_year, "Average Citations per Year Since Publication, by Year",
                    "Year", "Avg. Citations / Year Since Publication",
                    "avg_citations_per_year_normalized.svg",
                    fmt="%.2f")
    save_series_latex_table(
        norm_by_year,
        filename="avg_citations_per_year_normalized.tex",
        caption=(f"Average citations per year since publication, by publication year "
                 f"(normalized to {current_year}, based on {total_papers} papers)"),
        label="tab:avg_citations_per_year_normalized",
        col1_name="Year", col2_name="Avg. Citations/Year",
        value_fmt="{:.2f}",
    )

    return avg_by_year, norm_by_year


# ----------------------------------------------------------------------
# 6) TOP / MOST FREQUENT KEYWORDS + WORD CLOUD
# ----------------------------------------------------------------------
def top_keywords(df, top_n=TOP_N, use_author_kw=True, use_index_kw=True):
    counter = Counter()
    cols = []
    if use_author_kw and COL_KW_AUTHOR in df.columns:
        cols.append(COL_KW_AUTHOR)
    if use_index_kw and COL_KW_INDEX in df.columns:
        cols.append(COL_KW_INDEX)

    for col in cols:
        for val in df[col]:
            kws = split_multi(val)
            counter.update([k.lower().strip() for k in kws])

    total_keywords = len(counter)

    top = pd.Series(dict(counter.most_common(top_n)))
    top.to_csv(os.path.join(OUTPUT_DIR, "top_keywords.csv"), header=["count"])

    title = f"Top {top_n} Keywords (out of {total_keywords} unique keywords)"
    save_bar_chart(top, title, "Count", "Keyword", "top_keywords.svg", horizontal=True)

    return top, counter


def keyword_wordcloud(df, filename="keyword_wordcloud.svg", use_author_kw=True,
                        use_index_kw=True, max_words=150, width=1600, height=900,
                        background_color="white", colormap="viridis"):
    """Generate a tag cloud (word cloud) image from keyword frequencies."""
    from wordcloud import WordCloud

    counter = Counter()
    cols = []
    if use_author_kw and COL_KW_AUTHOR in df.columns:
        cols.append(COL_KW_AUTHOR)
    if use_index_kw and COL_KW_INDEX in df.columns:
        cols.append(COL_KW_INDEX)

    for col in cols:
        for val in df[col]:
            kws = split_multi(val)
            counter.update([k.lower().strip() for k in kws])

    if not counter:
        print("[!] No keywords found - check use_author_kw/use_index_kw and column names.")
        return None

    total_keywords = len(counter)

    wc = WordCloud(
        width=width, height=height, background_color=background_color,
        colormap=colormap, max_words=max_words, collocations=False,
        prefer_horizontal=0.9,
    ).generate_from_frequencies(counter)

    plt.figure(figsize=(width / 100, height / 100))
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")
    plt.title(f"Keyword Cloud ({total_keywords} unique keywords)", fontsize=14)
    plt.tight_layout()

    path = os.path.join(OUTPUT_DIR, filename)
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Saved word cloud: {path}")

    return wc


# ----------------------------------------------------------------------
# 7-8) CO-OCCURRENCE NETWORKS (keywords & authors)
# ----------------------------------------------------------------------
def _draw_network(G, title, filename_stem, layout="kamada_kawai", min_font=7, max_font=14,
                    figsize=(12, 10), for_pdf=True):
    """Draw a network and save as a publication-ready PDF (vector format),
    plus a PNG preview and an interactive HTML via pyvis."""
    if G.number_of_nodes() == 0:
        print(f"[!] No nodes to draw for {title} (try lowering the min_cooccur threshold).")
        return

    if layout == "kamada_kawai":
        try:
            pos = nx.kamada_kawai_layout(G, weight="weight")
        except ImportError:
            print("[i] scipy not installed - falling back to spring_layout. "
                  "Run 'pip install scipy' for better layouts.")
            pos = nx.spring_layout(G, k=0.6, seed=42, weight="weight", iterations=100)
    else:
        pos = nx.spring_layout(G, k=0.6, seed=42, weight="weight", iterations=100)

    weights = [G.nodes[n].get("weight", 1) for n in G.nodes]
    w_min, w_max = min(weights), max(weights)

    def scale_size(w):
        if w_max == w_min:
            return 300
        return 150 + (w - w_min) / (w_max - w_min) * 1200

    def scale_font(w):
        if w_max == w_min:
            return (min_font + max_font) / 2
        return min_font + (w - w_min) / (w_max - w_min) * (max_font - min_font)

    node_sizes = [scale_size(w) for w in weights]
    edge_weights = [G[u][v].get("weight", 1) for u, v in G.edges]
    e_min, e_max = (min(edge_weights), max(edge_weights)) if edge_weights else (1, 1)
    edge_widths = [0.4 + (w - e_min) / (e_max - e_min + 1e-9) * 2.5 for w in edge_weights]

    try:
        communities = nx.community.greedy_modularity_communities(G, weight="weight")
        community_map = {}
        for i, comm in enumerate(communities):
            for node in comm:
                community_map[node] = i
        node_colors = [community_map.get(n, 0) for n in G.nodes]
        cmap = plt.cm.tab20
    except Exception:
        node_colors = "#4C72B0"
        cmap = None

    fig, ax = plt.subplots(figsize=figsize)

    nx.draw_networkx_edges(G, pos, ax=ax, width=edge_widths, alpha=0.35, edge_color="gray")
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                            node_color=node_colors, cmap=cmap, alpha=0.9,
                            linewidths=0.5, edgecolors="black")

    for node, (x, y) in pos.items():
        ax.text(x, y, node, fontsize=scale_font(G.nodes[node].get("weight", 1)),
                ha="center", va="center", zorder=5,
                bbox=dict(facecolor="white", edgecolor="none", alpha=0.6, pad=0.5))

    ax.set_title(title, fontsize=16)
    ax.axis("off")
    plt.tight_layout()

    if for_pdf:
        pdf_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.pdf")
        plt.savefig(pdf_path, format="pdf", bbox_inches="tight")
        print(f"Saved print-ready PDF: {pdf_path}")

    png_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.svg")
    plt.savefig(png_path, dpi=200, bbox_inches="tight")
    plt.close()
    print(f"Saved preview PNG: {png_path}")

    try:
        from pyvis.network import Network
        net = Network(height="800px", width="100%", notebook=False, cdn_resources="in_line")
        for n, data in G.nodes(data=True):
            net.add_node(n, label=n, value=data.get("weight", 1))
        for u, v, data in G.edges(data=True):
            net.add_edge(u, v, value=data.get("weight", 1))
        html_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.html")
        #net.write_html(html_path)
        html = net.generate_html()
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Saved interactive network: {html_path}")
    except ImportError:
        print("[i] Install 'pyvis' for an interactive HTML version (pip install pyvis).")


def save_node_stats_table(G, filename_stem, min_cooccur, name_col="Name", top_n=None,
                            sort_by="weight", caption_prefix="Network node statistics",
                            label=None, max_name_len=50):
    """Export a table of node name, size (frequency), degree (n. connections), and
    weighted degree (sum of edge weights) from a networkx graph. Saves CSV + LaTeX.

    sort_by: 'weight' (node frequency), 'degree' (n. connections), or
             'weighted_degree' (total co-occurrence volume)
    top_n: if set, only the top N nodes by sort_by are included; None = all nodes
    min_cooccur: the co-occurrence threshold used to build G, shown in the caption
    """
    rows = []
    for node in G.nodes:
        rows.append({
            name_col: node,
            "size": G.nodes[node].get("weight", 0),
            "degree": G.degree(node),
            "weighted_degree": G.degree(node, weight="weight"),
        })

    stats = pd.DataFrame(rows)

    sort_col_map = {"weight": "size", "size": "size", "degree": "degree",
                     "weighted_degree": "weighted_degree"}
    sort_col = sort_col_map.get(sort_by, "size")

    stats = stats.sort_values(sort_col, ascending=False).reset_index(drop=True)
    total_nodes = stats.shape[0]

    stats_out = stats.head(top_n) if top_n is not None else stats

    csv_path = os.path.join(OUTPUT_DIR, f"{filename_stem}_node_stats.csv")
    stats_out.to_csv(csv_path, index=False)
    print(f"Saved node stats CSV: {csv_path}")

    shown_n = stats_out.shape[0]
    label = label or f"tab:{filename_stem}_node_stats"
    caption = (f"{caption_prefix} (showing {shown_n} of {total_nodes} nodes, "
               f"sorted by {sort_by}, min. co-occurrence = {min_cooccur})")

    lines = []
    lines.append(r"\begin{table}[htbp]")
    lines.append(r"\centering")
    lines.append(r"\caption{" + escape_latex(caption) + "}")
    lines.append(r"\label{" + label + "}")
    lines.append(r"\begin{tabular}{p{6cm}ccc}")
    lines.append(r"\toprule")
    lines.append(f"{name_col} & Size (freq.) & Degree & Weighted Degree \\\\")
    lines.append(r"\midrule")

    for i, row in stats_out.iterrows():
        name = str(row[name_col])
        if max_name_len and len(name) > max_name_len:
            name = name[:max_name_len].rstrip() + "..."
        name = escape_latex(name)

        color = "white" if i % 2 == 0 else "gray!15"
        lines.append(r"\rowcolor{" + color + "}")
        lines.append(f"{name} & {int(row['size'])} & {int(row['degree'])} & "
                      f"{int(row['weighted_degree'])} \\\\")

    lines.append(r"\bottomrule")
    lines.append(r"\end{tabular}")
    lines.append(r"\end{table}")

    tex_path = os.path.join(OUTPUT_DIR, f"{filename_stem}_node_stats.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Saved node stats LaTeX table: {tex_path}")

    return stats


def keyword_cooccurrence_network(df, min_cooccur=KW_MIN_COOCCUR, use_author_kw=True,
                                   use_index_kw=True, top_n_nodes=50):
    cols = []
    if use_author_kw and COL_KW_AUTHOR in df.columns:
        cols.append(COL_KW_AUTHOR)
    if use_index_kw and COL_KW_INDEX in df.columns:
        cols.append(COL_KW_INDEX)

    edge_counter = Counter()
    node_counter = Counter()

    for _, row in df.iterrows():
        kws = set()
        for col in cols:
            kws.update(k.lower().strip() for k in split_multi(row.get(col)))
        kws = sorted(kws)
        node_counter.update(kws)
        for a, b in combinations(kws, 2):
            edge_counter[(a, b)] += 1

    total_keywords = len(node_counter)

    top_nodes = {k for k, _ in node_counter.most_common(top_n_nodes)}

    G = nx.Graph()
    for kw, freq in node_counter.items():
        if kw in top_nodes:
            G.add_node(kw, weight=freq)
    for (a, b), w in edge_counter.items():
        if a in top_nodes and b in top_nodes and w >= min_cooccur:
            G.add_edge(a, b, weight=w)

    nodes_with_edges = {n for n in G.nodes if G.degree(n) > 0}
    shown_n = len(nodes_with_edges)

    title = (f"Keyword Co-occurrence Network\n"
             f"({shown_n} of {total_keywords} keywords shown, "
             f"min. co-occurrence = {min_cooccur})")
    _draw_network(G, title, "keyword_cooccurrence_network")

    save_node_stats_table(G, "keyword_cooccurrence_network", min_cooccur=min_cooccur,
                            name_col="Keyword",
                            caption_prefix="Keyword co-occurrence network statistics",
                            top_n=30)
    return G


def author_cooccurrence_network(df, min_cooccur=AUTHOR_MIN_COOCCUR, top_n_nodes=50):
    edge_counter = Counter()
    node_counter = Counter()

    for val in df[COL_AUTHORS]:
        authors = sorted(set(a.strip() for a in split_multi(val) if a.strip()))
        node_counter.update(authors)
        for a, b in combinations(authors, 2):
            edge_counter[(a, b)] += 1

    total_authors = len(node_counter)

    top_nodes = {k for k, _ in node_counter.most_common(top_n_nodes)}

    G = nx.Graph()
    for author, freq in node_counter.items():
        if author in top_nodes:
            G.add_node(author, weight=freq)
    for (a, b), w in edge_counter.items():
        if a in top_nodes and b in top_nodes and w >= min_cooccur:
            G.add_edge(a, b, weight=w)

    nodes_with_edges = {n for n in G.nodes if G.degree(n) > 0}
    shown_n = len(nodes_with_edges)

    title = (f"Author Co-authorship Network\n"
             f"({shown_n} of {total_authors} authors shown, "
             f"min. co-authorship = {min_cooccur})")
    _draw_network(G, title, "author_cooccurrence_network")

    save_node_stats_table(G, "author_cooccurrence_network", min_cooccur=min_cooccur,
                            name_col="Author",
                            caption_prefix="Author co-authorship network statistics",
                            top_n=30)
    return G


# ----------------------------------------------------------------------
# 9) LOTKA'S LAW CHECK (classic c=2 vs. empirically fitted c)
# ----------------------------------------------------------------------
def lotkas_law_check(df, filename_stem="lotkas_law", max_n_display=10):
    """Check author productivity distribution against Lotka's Law, comparing:
    - Observed distribution (your actual data)
    - Classic theoretical Lotka curve (c = 2, Lotka 1926)
    - Empirically fitted curve (c estimated from this dataset via log-log
      linear regression), more appropriate for an emergent/narrow field where
      the classic c=2 (derived from mature 1920s chemistry/physics literature)
      may not hold.
    Also runs a Kolmogorov-Smirnov goodness-of-fit test (per Pao, 1985) for the
    fitted curve against the observed data.
    """
    if COL_AUTHORS not in df.columns:
        print("[!] Authors column not found - skipping Lotka's Law check.")
        return None

    author_paper_count = Counter()
    for val in df[COL_AUTHORS]:
        authors = set(a.strip() for a in split_multi(val) if a.strip())
        for author in authors:
            author_paper_count[author] += 1

    total_authors = len(author_paper_count)
    if total_authors == 0:
        print("[!] No authors found - skipping Lotka's Law check.")
        return None

    productivity = Counter(author_paper_count.values())
    max_n = max(productivity.keys())

    ns = np.array(sorted(productivity.keys()))
    a_n = np.array([productivity[n] for n in ns])

    # --- Empirically fit c via log-log linear regression: log(a_n) = log(a_1) - c*log(n) ---
    log_n = np.log(ns)
    log_a = np.log(a_n)
    slope, intercept, r_value, p_value, std_err = stats.linregress(log_n, log_a)
    c_fitted = -slope
    r_squared = r_value ** 2

    print(f"Empirically fitted Lotka exponent: c = {c_fitted:.3f} "
          f"(vs. classic c = 2.0); R-squared of log-log fit = {r_squared:.3f}")

    # --- Build distribution table: observed, classic (c=2), and fitted ---
    rows = []
    for n in range(1, max_n + 1):
        n_authors_observed = productivity.get(n, 0)
        pct_observed = 100 * n_authors_observed / total_authors
        rows.append({"n_papers": n, "n_authors_observed": n_authors_observed,
                      "pct_observed": pct_observed})

    dist = pd.DataFrame(rows)

    c_classic_normalizer = sum(1 / (n ** 2) for n in dist["n_papers"])
    dist["pct_theoretical_classic"] = dist["n_papers"].apply(
        lambda n: 100 * (1 / (n ** 2)) / c_classic_normalizer
    )

    c_fitted_normalizer = sum(1 / (n ** c_fitted) for n in dist["n_papers"])
    dist["pct_theoretical_fitted"] = dist["n_papers"].apply(
        lambda n: 100 * (1 / (n ** c_fitted)) / c_fitted_normalizer
    )

    # --- Goodness-of-fit: D statistic computed directly on the grouped cumulative
    # frequency table (Pao, 1985), NOT via scipy's ks_1samp. ks_1samp assumes a
    # continuous distribution with no ties; applying it to this data (where
    # thousands of authors tie at n=1) produces a meaningless, inflated statistic
    # (its internal tie-handling compares the hypothesized CDF at n=1 against an
    # empirical CDF of 0 "just before" the data starts, rather than the actual
    # gap between the two cumulative curves). The correct procedure for grouped
    # discrete bibliometric data is to compare cumulative percentages directly at
    # each n, then check D against the standard K-S critical value 1.36/sqrt(N)
    # (alpha=0.05, two-sided) rather than a derived p-value.
    cum_obs = np.cumsum(dist["pct_observed"].values) / 100
    cum_fit = np.cumsum(dist["pct_theoretical_fitted"].values) / 100
    D = np.max(np.abs(cum_obs - cum_fit))
    D_critical_05 = 1.36 / np.sqrt(total_authors)
    D_critical_01 = 1.63 / np.sqrt(total_authors)
    fit_ok = D <= D_critical_05

    print(f"Goodness-of-fit (Pao 1985 method): D = {D:.4f}, "
          f"critical value (alpha=0.05) = {D_critical_05:.4f} "
          f"({'fit acceptable' if fit_ok else 'fit rejected'} at alpha=0.05)")

    csv_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.csv")
    dist.to_csv(csv_path, index=False)
    print(f"Saved Lotka's Law distribution CSV: {csv_path}")

    # --- Chart: observed bars + classic curve + fitted curve ---
    plot_df = dist[dist["n_papers"] <= max_n_display].set_index("n_papers")

    fig, ax = plt.subplots(figsize=(10, 6))
    x = plot_df.index.astype(str)
    positions = range(len(x))

    bars_obs = ax.bar(positions, plot_df["pct_observed"],
                       width=0.6, label="Observed", color="#4C72B0")
    ax.plot(positions, plot_df["pct_theoretical_classic"], color="#C44E52",
            marker="o", linewidth=2, linestyle="--",
            label="Theoretical, classic (c = 2.0, Lotka 1926)")
    ax.plot(positions, plot_df["pct_theoretical_fitted"], color="#55A868",
            marker="s", linewidth=2,
            label=f"Theoretical, fitted (c = {c_fitted:.2f}, this dataset)")

    ax.bar_label(bars_obs, padding=3, fontsize=8, fmt="%.1f")
    ax.set_xticks(list(positions))
    ax.set_xticklabels(x)
    ax.set_xlabel("Number of Papers Published (n)")
    ax.set_ylabel("% of Authors")
    ax.set_title(f"Lotka's Law Check: Author Productivity Distribution\n"
                 f"({total_authors} distinct authors, showing n=1-{max_n_display}, "
                 f"fitted R\u00b2={r_squared:.3f})")
    ax.legend()
    plt.tight_layout()

    png_path = os.path.join(OUTPUT_DIR, f"{filename_stem}.svg")
    plt.savefig(png_path, dpi=150)
    plt.close()
    print(f"Saved Lotka's Law chart: {png_path}")

    n_one_paper = productivity.get(1, 0)
    pct_one_paper = 100 * n_one_paper / total_authors
    print(f"\nLotka's Law summary:")
    print(f"  Total distinct authors: {total_authors}")
    print(f"  Authors with exactly 1 paper: {n_one_paper} ({pct_one_paper:.1f}%)")
    print(f"  (Classic Lotka's Law, c=2, predicts ~60% of authors publish only 1 paper)")
    print(f"  Fitted exponent for this dataset: c = {c_fitted:.3f} (R\u00b2 = {r_squared:.3f})")
    print(f"  Goodness-of-fit: D = {D:.4f} vs. critical {D_critical_05:.4f} (alpha=0.05) "
          f"-> {'acceptable' if fit_ok else 'rejected'}")

    return dist


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Bibliometric analysis of a Scopus/WoS export.")
    parser.add_argument("excel_path", help="Path to the .xlsx file")
    parser.add_argument("--sheet", default=0, help="Sheet name or index (default: first sheet)")
    parser.add_argument("--top_n", type=int, default=TOP_N, help="Top N items for rankings")
    args = parser.parse_args()

    ensure_outdir()

    print(f"Reading {args.excel_path} ...")
    df = pd.read_excel(args.excel_path, sheet_name=args.sheet)
    print(f"Loaded {len(df)} rows.")

    print_dataset_summary(df)

    print("\nData coverage by field")
    df["_has_any_citations"] = df.apply(
        lambda r: is_filled(r.get(COL_CIT_SCOPUS)) or is_filled(r.get(COL_CIT_WOS)), axis=1
    )
    coverage_groups = {
        "Authorship and keywords": [
            ("Authors", COL_AUTHORS),
            ("Affiliations", "Afiliacoes"),
            ("Authors with affiliation", "Autores_com_afiliacao"),
            ("Author keywords", COL_KW_AUTHOR),
            ("Indexed keywords", COL_KW_INDEX),
        ],
        "Publication metadata": [
            ("Title", COL_TITLE),
            ("Journal", COL_JOURNAL),
            ("Year", COL_YEAR),
            ("Language", "Idioma"),
            ("Publisher", "Editora"),
            ("Issue", "Fasciculo"),
            ("Funding", "Financiamento"),
            ("Citations (Scopus)", COL_CIT_SCOPUS),
            ("Citations (WoS)", COL_CIT_WOS),
            ("Citations (Scopus or WoS)", "_has_any_citations"),
        ],
    }
    save_coverage_table(df, coverage_groups,
                         title="Data Coverage",
                         caption="Data coverage by field",
                         label="tab:data_coverage")
    df = df.drop(columns=["_has_any_citations"])

    print("\n1) Publications per year")
    publications_per_year(df)

    print("\n2) Top journals publishing")
    top_journals_publishing(df, top_n=args.top_n)
    top_journals_dual_chart(df, top_n=30)

    print("\n3) Top journals getting cited")
    top_journals_cited(df, top_n=args.top_n)
    top_cited_journals_dual_chart(df, top_n=30)
    top_journals_by_citation_ratio(df, top_n=15, min_pubs=3)
    journals_scatter_pub_vs_citations(df)

    print("\n4) Most cited papers")
    most_cited_papers(df, top_n=args.top_n)
    top_cited_papers_with_journal_context(df, top_n=args.top_n)

    print("\n5) Average citations per year")
    avg_citations_per_year(df)

    print("\n6) Most frequent keywords")
    top_keywords(df, top_n=args.top_n)
    keyword_wordcloud(df)

    print("\n7) Keyword co-occurrence network")
    keyword_cooccurrence_network(df)

    print("\n8) Authorship co-occurrence network")
    author_cooccurrence_network(df)

    print("\n9) Lotka's Law check (author productivity distribution)")
    lotkas_law_check(df)

    print(f"\nDone. All results saved in ./{OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
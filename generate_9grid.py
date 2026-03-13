from __future__ import annotations

from pathlib import Path
import math

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Rectangle


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "9Grid exercice.xlsx"
OUTPUT_DIR = SCRIPT_DIR / "output"
SHEET_NAME = "Data"

FULL_REQUIRED_COLUMNS = [
    "Team Member",
    "Action Bucket",
    "Owner",
    "Churn Risk",
    "Main Strength",
    "Main Concern",
    "Rationale",
    "Performance Score",
    "Trajectory Score",
]

COMPACT_REQUIRED_COLUMNS = [
    "Lead Name",
    "Team Member",
    "9Grid number",
]

X_LABELS = {
    1: "Below bar",
    2: "At bar",
    3: "Above bar",
}

Y_LABELS = {
    1: "Stable",
    2: "Growable",
    3: "High trajectory",
}

OWNER_COLOR_MAP = {
    "brecht": "#1F77B4",
    "gary": "#2CA02C",
    "stephanie": "#7A52A1",
    "bart": "#008B8B",
}

FALLBACK_COLORS = [
    "#4E79A7",
    "#59A14F",
    "#F28E2B",
    "#E15759",
    "#76B7B2",
    "#EDC948",
    "#B07AA1",
    "#9C755F",
]

PREFERRED_OWNER_ORDER = ["Brecht", "Gary", "Stephanie", "Bart"]

CHURN_RISK_STYLES = {
    "low": {"edgecolor": "#C7CDD6", "linewidth": 0.9, "size": 360, "halo_size": 0, "halo_width": 0},
    "medium": {"edgecolor": "#FF9F1C", "linewidth": 3.0, "size": 420, "halo_size": 500, "halo_width": 1.8},
    "high": {"edgecolor": "#D92D20", "linewidth": 4.2, "size": 480, "halo_size": 620, "halo_width": 2.8},
    "unknown": {"edgecolor": "#C7CDD6", "linewidth": 0.9, "size": 360, "halo_size": 0, "halo_width": 0},
}

GRID_NUMBER_MAP = {
    1: (1, 1),
    2: (2, 1),
    3: (3, 1),
    4: (1, 2),
    5: (2, 2),
    6: (3, 2),
    7: (1, 3),
    8: (2, 3),
    9: (3, 3),
}


def load_data(excel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    df = normalize_columns(df)

    if all(col in df.columns for col in FULL_REQUIRED_COLUMNS):
        return load_full_format(df)

    if all(col in df.columns for col in COMPACT_REQUIRED_COLUMNS):
        return load_compact_format(df)

    missing_full = [col for col in FULL_REQUIRED_COLUMNS if col not in df.columns]
    missing_compact = [col for col in COMPACT_REQUIRED_COLUMNS if col not in df.columns]
    raise ValueError(
        "Unsupported sheet structure.\n"
        f"Missing columns for full format: {', '.join(missing_full) or 'none'}\n"
        f"Missing columns for compact format: {', '.join(missing_compact) or 'none'}"
    )


def load_full_format(df: pd.DataFrame) -> pd.DataFrame:
    df = df[FULL_REQUIRED_COLUMNS].copy()
    df["Team Member"] = df["Team Member"].astype("string").str.strip()
    df = df[df["Team Member"].notna() & (df["Team Member"] != "")]
    df = df[df["Performance Score"].notna() & df["Trajectory Score"].notna()]

    df["Performance Score"] = pd.to_numeric(df["Performance Score"], errors="coerce")
    df["Trajectory Score"] = pd.to_numeric(df["Trajectory Score"], errors="coerce")
    df = df[df["Performance Score"].between(1, 3) & df["Trajectory Score"].between(1, 3)]

    for col in ["Action Bucket", "Owner", "Churn Risk", "Main Strength", "Main Concern", "Rationale"]:
        df[col] = df[col].fillna("").astype("string").str.strip()

    df["Assigned Score"] = (
        df["Performance Score"].astype(int) + (df["Trajectory Score"].astype(int) - 1) * 3
    )
    df = df.reset_index(drop=True)
    return assign_plot_numbers(df)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed_columns = {}
    for col in df.columns:
        normalized = " ".join(str(col).replace("\n", " ").split()).strip()
        lowered = normalized.casefold()

        if lowered == "lead name":
            renamed_columns[col] = "Lead Name"
        elif lowered == "team member":
            renamed_columns[col] = "Team Member"
        elif lowered in {"9grid number", "9 grid number", "9-grid number", "9grid", "9 grid"}:
            renamed_columns[col] = "9Grid number"
        elif lowered == "action bucket":
            renamed_columns[col] = "Action Bucket"
        elif lowered == "owner":
            renamed_columns[col] = "Owner"
        elif lowered in {
            "churn risk",
            "churnrisk",
            "risk of churn (1->3 - low medium high)",
            "risk of churn",
        }:
            renamed_columns[col] = "Churn Risk"
        elif lowered == "main strength":
            renamed_columns[col] = "Main Strength"
        elif lowered == "main concern":
            renamed_columns[col] = "Main Concern"
        elif lowered == "rationale":
            renamed_columns[col] = "Rationale"
        elif lowered == "performance score":
            renamed_columns[col] = "Performance Score"
        elif lowered == "trajectory score":
            renamed_columns[col] = "Trajectory Score"
        else:
            renamed_columns[col] = normalized

    return df.rename(columns=renamed_columns)


def load_compact_format(df: pd.DataFrame) -> pd.DataFrame:
    selected_columns = COMPACT_REQUIRED_COLUMNS.copy()
    if "Action Bucket" in df.columns:
        selected_columns.append("Action Bucket")

    compact = df[selected_columns].copy()
    compact["Team Member"] = compact["Team Member"].astype("string").str.strip()
    compact = compact[compact["Team Member"].notna() & (compact["Team Member"] != "")]
    compact["9Grid number"] = pd.to_numeric(compact["9Grid number"], errors="coerce")
    compact = compact[compact["9Grid number"].isin(GRID_NUMBER_MAP)]

    compact["Lead Name"] = compact["Lead Name"].fillna("").astype("string").str.strip()
    if "Action Bucket" not in compact.columns:
        compact["Action Bucket"] = ""
    compact["Action Bucket"] = compact["Action Bucket"].fillna("").astype("string").str.strip()
    compact["Owner"] = compact["Lead Name"]
    compact["Churn Risk"] = df["Churn Risk"].fillna("").astype("string").str.strip() if "Churn Risk" in df.columns else ""
    compact["Main Strength"] = ""
    compact["Main Concern"] = ""
    compact["Rationale"] = ""
    compact["Performance Score"] = compact["9Grid number"].map(lambda value: GRID_NUMBER_MAP[int(value)][0])
    compact["Trajectory Score"] = compact["9Grid number"].map(lambda value: GRID_NUMBER_MAP[int(value)][1])
    compact["Assigned Score"] = compact["9Grid number"].astype(int)

    compact = compact.reset_index(drop=True)
    compact = compact[
        [
            "Team Member",
            "Action Bucket",
            "Owner",
            "Churn Risk",
            "Main Strength",
            "Main Concern",
            "Rationale",
            "Performance Score",
            "Trajectory Score",
            "Assigned Score",
        ]
    ]
    return assign_plot_numbers(compact)


def assign_plot_numbers(df: pd.DataFrame) -> pd.DataFrame:
    numbered = df.copy()
    numbered["Team Member Sort"] = numbered["Team Member"].astype(str).str.casefold()
    unique_members = (
        numbered[["Team Member", "Team Member Sort"]]
        .drop_duplicates()
        .sort_values(["Team Member Sort", "Team Member"])
        .reset_index(drop=True)
    )
    unique_members["Plot Number"] = range(1, len(unique_members) + 1)
    numbered = numbered.merge(unique_members[["Team Member", "Plot Number"]], on="Team Member", how="left")
    return numbered.drop(columns=["Team Member Sort"])


def get_cluster_centers(base_x: float, base_y: float, cluster_count: int) -> list[tuple[float, float]]:
    if cluster_count <= 1:
        return [(base_x, base_y)]
    if cluster_count == 2:
        return [(base_x - 0.22, base_y), (base_x + 0.22, base_y)]
    if cluster_count == 3:
        return [
            (base_x, base_y + 0.19),
            (base_x - 0.22, base_y - 0.16),
            (base_x + 0.22, base_y - 0.16),
        ]
    return [
        (base_x - 0.22, base_y + 0.18),
        (base_x + 0.22, base_y + 0.18),
        (base_x - 0.22, base_y - 0.18),
        (base_x + 0.22, base_y - 0.18),
    ]


def spread_cluster(
    plotted: pd.DataFrame,
    indices: list[int],
    center_x: float,
    center_y: float,
    jitter_radius: float,
) -> None:
    count = len(indices)
    if count == 1:
        plotted.at[indices[0], "x"] = center_x
        plotted.at[indices[0], "y"] = center_y
        return

    angles = [2 * math.pi * i / count for i in range(count)]
    for idx, angle in zip(indices, angles):
        plotted.at[idx, "x"] = center_x + jitter_radius * math.cos(angle)
        plotted.at[idx, "y"] = center_y + jitter_radius * math.sin(angle)


def compute_positions(
    df: pd.DataFrame,
    jitter_radius: float = 0.15,
    cluster_threshold: int = 12,
) -> pd.DataFrame:
    plotted = df.copy()
    plotted["x"] = plotted["Performance Score"].astype(float)
    plotted["y"] = plotted["Trajectory Score"].astype(float)

    for _, cell_idx in plotted.groupby(["Performance Score", "Trajectory Score"]).groups.items():
        indices = list(cell_idx)
        count = len(indices)
        if count <= 1:
            continue

        base_x = float(plotted.at[indices[0], "Performance Score"])
        base_y = float(plotted.at[indices[0], "Trajectory Score"])
        cluster_count = max(1, math.ceil(count / cluster_threshold))
        cluster_count = min(cluster_count, 4)
        centers = get_cluster_centers(base_x, base_y, cluster_count)

        for cluster_index, center in enumerate(centers):
            start = cluster_index * cluster_threshold
            end = start + cluster_threshold
            cluster_indices = indices[start:end]
            if not cluster_indices:
                continue
            spread_cluster(plotted, cluster_indices, center[0], center[1], jitter_radius)

    return plotted


def sanitize_filename(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in value.strip())
    cleaned = "_".join(cleaned.split())
    return cleaned or "unknown"


def build_legend_table(df: pd.DataFrame) -> pd.DataFrame:
    legend = df[["Plot Number", "Team Member", "Action Bucket", "Owner", "Churn Risk"]].copy()
    legend = legend.rename(columns={"Plot Number": "Number"})
    return legend


def get_owner_display_order(df: pd.DataFrame) -> list[str]:
    owners = []
    for owner in df["Owner"].fillna("").astype(str):
        clean_owner = owner.strip() or "Unassigned"
        if clean_owner not in owners:
            owners.append(clean_owner)

    preferred = [owner for owner in PREFERRED_OWNER_ORDER if owner in owners]
    remaining = sorted([owner for owner in owners if owner not in preferred], key=str.casefold)
    return preferred + remaining


def build_overview_summary_table(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df[["Plot Number", "Team Member", "Owner", "Assigned Score"]]
        .copy()
        .assign(Owner=lambda data: data["Owner"].replace("", "Unassigned"))
    )
    score_table = (
        summary.pivot_table(
            index=["Plot Number", "Team Member"],
            columns="Owner",
            values="Assigned Score",
            aggfunc=lambda values: "/".join(str(int(v)) for v in sorted(set(values))),
            fill_value="",
        )
        .reset_index()
    )

    owner_columns = get_owner_display_order(df)
    for owner in owner_columns:
        if owner not in score_table.columns:
            score_table[owner] = ""

    score_table = score_table[["Plot Number", "Team Member", *owner_columns]]
    return score_table.rename(columns={"Plot Number": "Number"})


def get_owner_colors(df: pd.DataFrame) -> dict[str, str]:
    owners = get_owner_display_order(df)
    colors: dict[str, str] = {}
    fallback_index = 0
    for owner in owners:
        mapped = OWNER_COLOR_MAP.get(owner.casefold())
        if mapped:
            colors[owner] = mapped
            continue
        colors[owner] = FALLBACK_COLORS[fallback_index % len(FALLBACK_COLORS)]
        fallback_index += 1

    return colors


def normalize_churn_risk(value: object) -> str:
    text = str(value).strip().casefold()
    if text in {"3", "high", "h", "red"}:
        return "high"
    if text in {"2", "medium", "med", "m", "moderate", "amber", "orange"}:
        return "medium"
    if text in {"1", "low", "l", "green"}:
        return "low"
    return "unknown"


def draw_grid(ax: plt.Axes) -> None:
    ax.set_xlim(0.5, 3.5)
    ax.set_ylim(0.5, 3.5)
    ax.set_xticks([1, 2, 3], [X_LABELS[1], X_LABELS[2], X_LABELS[3]])
    ax.set_yticks([1, 2, 3], [Y_LABELS[1], Y_LABELS[2], Y_LABELS[3]])
    ax.set_xlabel("Current performance", fontsize=11, labelpad=10)
    ax.set_ylabel("Growth trajectory / ownership", fontsize=11, labelpad=10)

    for x in [0.5, 1.5, 2.5]:
        for y in [0.5, 1.5, 2.5]:
            ax.add_patch(
                Rectangle(
                    (x, y),
                    1.0,
                    1.0,
                    facecolor="white",
                    edgecolor="#B8C0CC",
                    linewidth=1.0,
                    zorder=0,
                )
            )

    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(axis="both", length=0, labelsize=10)


def draw_points(ax: plt.Axes, plotted_df: pd.DataFrame, owner_colors: dict[str, str]) -> None:
    for _, row in plotted_df.iterrows():
        owner = str(row["Owner"]).strip() or "Unassigned"
        fill_color = owner_colors.get(owner, "#1F4E79")
        churn_style = CHURN_RISK_STYLES[normalize_churn_risk(row.get("Churn Risk", ""))]
        if churn_style["halo_size"] > 0:
            ax.scatter(
                [row["x"]],
                [row["y"]],
                s=churn_style["halo_size"],
                facecolors="none",
                edgecolors=[churn_style["edgecolor"]],
                linewidths=churn_style["halo_width"],
                zorder=2.8,
            )
        ax.scatter(
            [row["x"]],
            [row["y"]],
            s=churn_style["size"],
            c=[fill_color],
            edgecolors=[churn_style["edgecolor"]],
            linewidths=churn_style["linewidth"],
            zorder=3,
        )
        ax.text(
            row["x"],
            row["y"],
            str(int(row["Plot Number"])),
            ha="center",
            va="center",
            color="white",
            fontsize=10,
            fontweight="bold",
            zorder=4,
        )


def draw_legend_panel(
    ax: plt.Axes,
    legend_df: pd.DataFrame,
    owner_colors: dict[str, str],
    is_overview: bool,
) -> None:
    ax.axis("off")
    ax.set_title("Legend / Discussion", loc="left", fontsize=12, fontweight="bold", pad=10)

    y = 0.98
    ax.text(0.0, y, "Manager colors", transform=ax.transAxes, fontsize=10, fontweight="bold", va="top")
    y -= 0.04
    for owner, color in owner_colors.items():
        ax.text(0.0, y, "\u25CF", color=color, transform=ax.transAxes, fontsize=12, va="top")
        ax.text(0.04, y, owner, transform=ax.transAxes, fontsize=9, va="top", color="#203040")
        y -= 0.035

    y -= 0.01
    ax.text(0.0, y, "Churn risk outline", transform=ax.transAxes, fontsize=10, fontweight="bold", va="top")
    y -= 0.04
    for label, key in [("Low", "low"), ("Medium", "medium"), ("High", "high")]:
        style = CHURN_RISK_STYLES[key]
        ax.scatter(
            [0.015],
            [y - 0.01],
            s=style["size"] * 0.28,
            facecolors="white",
            edgecolors=style["edgecolor"],
            linewidths=style["linewidth"],
            transform=ax.transAxes,
            clip_on=False,
            zorder=3,
        )
        if style["halo_size"] > 0:
            ax.scatter(
                [0.015],
                [y - 0.01],
                s=style["halo_size"] * 0.28,
                facecolors="none",
                edgecolors=style["edgecolor"],
                linewidths=style["halo_width"],
                transform=ax.transAxes,
                clip_on=False,
                zorder=2,
            )
        ax.text(0.04, y, label, transform=ax.transAxes, fontsize=9, va="top", color="#203040")
        y -= 0.035

    headers = list(legend_df.columns)
    table_rows = [headers]
    for _, row in legend_df.iterrows():
        table_rows.append([str(row[col]) for col in headers])

    if is_overview:
        col_widths = [0.09, 0.33]
        owner_count = max(1, len(headers) - 2)
        score_width = (1.0 - sum(col_widths)) / owner_count
        col_widths.extend([score_width] * owner_count)
    else:
        col_widths = [0.07, 0.28, 0.22, 0.16, 0.27]

    table = ax.table(
        cellText=table_rows,
        cellLoc="left",
        colLabels=None,
        colWidths=col_widths,
        bbox=[0.0, 0.0, 1.0, max(0.65, y - 0.02)],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#D9DDE3")
        cell.set_linewidth(0.75)
        if row == 0:
            cell.set_facecolor("#EAF0F6")
            cell.set_text_props(weight="bold", color="#203040")
        else:
            cell.set_facecolor("white")


def create_chart(plotted_df: pd.DataFrame, title: str, png_path: Path, csv_path: Path) -> None:
    is_overview = "overview" in png_path.stem.casefold()
    legend_df = build_overview_summary_table(plotted_df) if is_overview else build_legend_table(plotted_df)
    owner_colors = get_owner_colors(plotted_df)

    fig = plt.figure(figsize=(16, 9), constrained_layout=True)
    gs = fig.add_gridspec(1, 2, width_ratios=[3.0, 2.2])
    ax_chart = fig.add_subplot(gs[0, 0])
    ax_legend = fig.add_subplot(gs[0, 1])

    draw_grid(ax_chart)
    draw_points(ax_chart, plotted_df, owner_colors)
    draw_legend_panel(ax_legend, legend_df, owner_colors, is_overview)

    fig.suptitle(title, fontsize=16, fontweight="bold", x=0.34)
    fig.savefig(png_path, dpi=300, facecolor="white", bbox_inches="tight")
    plt.close(fig)

    legend_df.to_csv(csv_path, index=False, encoding="utf-8-sig")


def export_owner_views(df: pd.DataFrame) -> int:
    owner_count = 0
    for owner, owner_df in df.groupby("Owner", dropna=False):
        owner_name = str(owner).strip() or "Unassigned"
        owner_df = owner_df.reset_index(drop=True)

        # "Enough data" is treated as at least 2 plotted employees for that owner.
        if len(owner_df) < 2:
            continue

        owner_df["Plot Number"] = range(1, len(owner_df) + 1)
        plotted_owner_df = compute_positions(owner_df)
        safe_owner = sanitize_filename(owner_name)
        create_chart(
            plotted_owner_df,
            f"9-Grid Talent Calibration - {owner_name}",
            OUTPUT_DIR / f"9grid_owner_{safe_owner}.png",
            OUTPUT_DIR / f"9grid_owner_{safe_owner}_legend.csv",
        )
        owner_count += 1

    return owner_count


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not DEFAULT_INPUT.exists():
        raise FileNotFoundError(
            f'Input file not found: "{DEFAULT_INPUT}". Place "9Grid exercice.xlsx" next to this script.'
        )

    df = load_data(DEFAULT_INPUT)
    if df.empty:
        raise ValueError("No valid rows found after applying the required filters.")

    plotted_df = compute_positions(df)
    create_chart(
        plotted_df,
        "9-Grid Talent Calibration",
        OUTPUT_DIR / "9grid_overview.png",
        OUTPUT_DIR / "9grid_overview_legend.csv",
    )

    owner_views = export_owner_views(df)
    print(f"Created overview chart for {len(df)} employees.")
    print(f"Created {owner_views} owner-specific chart(s).")
    print(f"Files saved to: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()

from __future__ import annotations

from datetime import date
from pathlib import Path
import math
import re
import shutil

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Rectangle
import win32com.client as win32


SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "output"
SHEET_NAME = "Data"
POWERBI_TEMPLATE_CANDIDATES = [
    SCRIPT_DIR / "9GRID.xlsb",
    Path.home() / "Downloads" / "9GRID.xlsb",
]
MANAGER_FILE_PATTERN = re.compile(r"^9grid_(?P<manager>.+)\.xlsx$", re.IGNORECASE)
DEFAULT_INPUT_CANDIDATES = [
    SCRIPT_DIR / "9Grid exercice.xlsx",
    SCRIPT_DIR / "9Grid_exercice.xlsx",
]
POWERBI_EXPORT_COLUMNS = [
    "Department",
    "Employee ID",
    "Name",
    "9Grid_Date",
    "Flight Risk",
    "Performance",
    "Potential",
    "Grid Box",
    "Feedback",
]
OPTIONAL_SOURCE_COLUMNS = [
    "Department",
    "Employee ID",
    "9Grid_Date",
    "Flight Risk",
    "Feedback",
]

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

HYBRID_REQUIRED_COLUMNS = [
    "Lead Name",
    "Team Member",
    "Action Bucket",
    "Owner",
    "Main Strength",
    "Main Concern",
    "Rationale",
    "Churn Risk",
    "Performance Score",
    "Trajectory Score",
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
EXCLUDED_TEAM_MEMBERS = {"john doe"}
SCORE_LABELS = {
    1: "Low",
    2: "Moderate",
    3: "High",
}


def extract_manager_name(excel_path: Path) -> str | None:
    match = MANAGER_FILE_PATTERN.match(excel_path.name)
    if not match:
        return None

    manager = match.group("manager").replace("_", " ").strip()
    return manager or None


def resolve_powerbi_template_path() -> Path:
    for candidate in POWERBI_TEMPLATE_CANDIDATES:
        if candidate.exists():
            return candidate

    expected_names = ", ".join(f'"{path}"' for path in POWERBI_TEMPLATE_CANDIDATES)
    raise FileNotFoundError(f"Power BI template not found. Expected one of: {expected_names}")


def resolve_input_paths() -> list[Path]:
    manager_files = sorted(
        [path for path in SCRIPT_DIR.glob("*.xlsx") if extract_manager_name(path)],
        key=lambda path: path.name.casefold(),
    )
    if manager_files:
        return manager_files

    for candidate in DEFAULT_INPUT_CANDIDATES:
        if candidate.exists():
            return [candidate]

    expected_names = ", ".join(f'"{path.name}"' for path in DEFAULT_INPUT_CANDIDATES)
    raise FileNotFoundError(
        "Input file not found. Place one or more files matching "
        '"9grid_<manager>.xlsx" next to this script, or use one of '
        f"{expected_names}."
    )


def load_data(excel_path: Path, manager_name: str | None = None) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    df = normalize_columns(df)

    if all(col in df.columns for col in HYBRID_REQUIRED_COLUMNS):
        loaded = load_hybrid_format(df)
    elif all(col in df.columns for col in FULL_REQUIRED_COLUMNS):
        loaded = load_full_format(df)
    elif all(col in df.columns for col in COMPACT_REQUIRED_COLUMNS):
        loaded = load_compact_format(df)
    else:
        missing_full = [col for col in FULL_REQUIRED_COLUMNS if col not in df.columns]
        missing_hybrid = [col for col in HYBRID_REQUIRED_COLUMNS if col not in df.columns]
        missing_compact = [col for col in COMPACT_REQUIRED_COLUMNS if col not in df.columns]
        raise ValueError(
            f'Unsupported sheet structure in "{excel_path.name}".\n'
            f"Missing columns for full format: {', '.join(missing_full) or 'none'}\n"
            f"Missing columns for hybrid format: {', '.join(missing_hybrid) or 'none'}\n"
            f"Missing columns for compact format: {', '.join(missing_compact) or 'none'}"
        )

    if manager_name:
        loaded = loaded.copy()
        loaded["Owner"] = manager_name

    loaded["Source File"] = excel_path.name
    return loaded


def load_full_format(df: pd.DataFrame) -> pd.DataFrame:
    selected_columns = FULL_REQUIRED_COLUMNS + [col for col in OPTIONAL_SOURCE_COLUMNS if col in df.columns]
    return prepare_plotting_frame(df[selected_columns].copy())


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed_columns = {}
    for col in df.columns:
        normalized = " ".join(str(col).replace("\n", " ").split()).strip()
        lowered = re.sub(r"\.\d+$", "", normalized).strip().casefold()

        if lowered == "lead name":
            renamed_columns[col] = "Lead Name"
        elif lowered == "name":
            renamed_columns[col] = "Team Member"
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
        elif lowered == "performance":
            renamed_columns[col] = "Performance Score"
        elif lowered == "performance score":
            renamed_columns[col] = "Performance Score"
        elif lowered == "potential":
            renamed_columns[col] = "Trajectory Score"
        elif lowered == "trajectory score":
            renamed_columns[col] = "Trajectory Score"
        else:
            renamed_columns[col] = normalized

    renamed_df = df.rename(columns=renamed_columns)
    return collapse_duplicate_columns(renamed_df)


def collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    collapsed = pd.DataFrame(index=df.index)
    for column_name in dict.fromkeys(df.columns):
        column_data = df.loc[:, df.columns == column_name]
        if isinstance(column_data, pd.Series):
            collapsed[column_name] = column_data
            continue

        merged = column_data.iloc[:, 0]
        for index in range(1, column_data.shape[1]):
            candidate = column_data.iloc[:, index]
            merged = merged.where(
                merged.notna() & (merged.astype("string").str.strip() != ""),
                candidate,
            )
        collapsed[column_name] = merged

    return collapsed


def prepare_plotting_frame(df: pd.DataFrame) -> pd.DataFrame:
    prepared = df.copy()
    prepared["Team Member"] = prepared["Team Member"].astype("string").str.strip()
    prepared = prepared[prepared["Team Member"].notna() & (prepared["Team Member"] != "")]
    prepared = prepared[~prepared["Team Member"].astype(str).str.casefold().isin(EXCLUDED_TEAM_MEMBERS)]

    prepared["Performance Score"] = pd.to_numeric(prepared["Performance Score"], errors="coerce")
    prepared["Trajectory Score"] = pd.to_numeric(prepared["Trajectory Score"], errors="coerce")
    prepared = prepared[prepared["Performance Score"].between(1, 3) & prepared["Trajectory Score"].between(1, 3)]

    for col in ["Action Bucket", "Owner", "Churn Risk", "Main Strength", "Main Concern", "Rationale"]:
        prepared[col] = prepared[col].fillna("").astype("string").str.strip()

    prepared["Assigned Score"] = (
        prepared["Performance Score"].astype(int) + (prepared["Trajectory Score"].astype(int) - 1) * 3
    )
    prepared = prepared.reset_index(drop=True)
    return assign_plot_numbers(prepared)


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
    compact["Department"] = ""
    compact["Employee ID"] = ""
    compact["9Grid_Date"] = ""
    compact["Flight Risk"] = ""
    compact["Feedback"] = ""
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
            "Department",
            "Employee ID",
            "9Grid_Date",
            "Flight Risk",
            "Feedback",
            "Performance Score",
            "Trajectory Score",
            "Assigned Score",
        ]
    ]
    return assign_plot_numbers(compact)


def load_hybrid_format(df: pd.DataFrame) -> pd.DataFrame:
    selected_columns = HYBRID_REQUIRED_COLUMNS + [col for col in OPTIONAL_SOURCE_COLUMNS if col in df.columns]
    hybrid = df[selected_columns].copy()
    lead_names = hybrid["Lead Name"].fillna("").astype("string").str.strip()
    owners = hybrid["Owner"].fillna("").astype("string").str.strip()
    hybrid["Owner"] = owners.where(owners != "", lead_names)
    return prepare_plotting_frame(
        hybrid[
            [
                "Team Member",
                "Action Bucket",
                "Owner",
                "Churn Risk",
                "Main Strength",
                "Main Concern",
                "Rationale",
                *[col for col in OPTIONAL_SOURCE_COLUMNS if col in hybrid.columns],
                "Performance Score",
                "Trajectory Score",
            ]
        ]
    )


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
    legend = df[["Plot Number", "Team Member", "Action Bucket", "Churn Risk"]].copy()
    legend = legend.rename(columns={"Plot Number": "Nr"})
    return legend


def get_owner_display_order(df: pd.DataFrame) -> list[str]:
    owners = []
    for owner in df["Owner"].fillna("").astype(str):
        clean_owner = owner.strip() or "Unassigned"
        if clean_owner not in owners:
            owners.append(clean_owner)

    return sorted(owners, key=str.casefold)


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
    return score_table.rename(columns={"Plot Number": "Nr"})


def get_owner_colors(df: pd.DataFrame) -> dict[str, str]:
    owners = get_owner_display_order(df)
    colors: dict[str, str] = {}
    for index, owner in enumerate(owners):
        colors[owner] = FALLBACK_COLORS[index % len(FALLBACK_COLORS)]

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


def format_score_label(value: object) -> str:
    numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(numeric):
        return SCORE_LABELS.get(int(numeric), "")

    text = str(value).strip().casefold()
    if text in {"1", "low", "l"}:
        return "Low"
    if text in {"2", "moderate", "medium", "med", "m"}:
        return "Moderate"
    if text in {"3", "high", "h"}:
        return "High"
    return ""


def format_flight_risk(value: object) -> str:
    normalized = normalize_churn_risk(value)
    if normalized == "low":
        return "Low"
    if normalized == "medium":
        return "Moderate"
    if normalized == "high":
        return "High"
    return ""


def build_grid_box_label(performance_label: str, potential_label: str) -> str:
    if not performance_label or not potential_label:
        return ""
    return f"{performance_label}-{potential_label}"


def clean_export_value(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def build_feedback_series(df: pd.DataFrame) -> pd.Series:
    feedback = pd.Series([""] * len(df), index=df.index, dtype="string")
    if "Feedback" in df.columns:
        feedback = df["Feedback"].map(clean_export_value).astype("string")

    rationale = df["Rationale"].map(clean_export_value).astype("string")
    return feedback.where(feedback.fillna("") != "", rationale).fillna("")


def build_powerbi_export_table(df: pd.DataFrame, export_date: str) -> pd.DataFrame:
    export = pd.DataFrame()
    export["Department"] = (
        df["Department"].map(clean_export_value) if "Department" in df.columns else ""
    )
    export["Employee ID"] = (
        df["Employee ID"].map(clean_export_value) if "Employee ID" in df.columns else ""
    )
    export["Name"] = df["Team Member"].map(clean_export_value)

    if "9Grid_Date" in df.columns:
        export["9Grid_Date"] = df["9Grid_Date"].map(clean_export_value).replace("", export_date)
    else:
        export["9Grid_Date"] = export_date

    if "Flight Risk" in df.columns:
        export["Flight Risk"] = df["Flight Risk"].map(format_flight_risk)
    else:
        export["Flight Risk"] = df["Churn Risk"].map(format_flight_risk)

    export["Performance"] = df["Performance Score"].map(format_score_label)
    export["Potential"] = df["Trajectory Score"].map(format_score_label)
    export["Grid Box"] = ""
    export["Feedback"] = build_feedback_series(df)

    return export[POWERBI_EXPORT_COLUMNS].reset_index(drop=True)


def write_powerbi_workbook(template_path: Path, output_path: Path, export_df: pd.DataFrame) -> None:
    shutil.copy2(template_path, output_path)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    worksheet = None

    try:
        workbook = excel.Workbooks.Open(str(output_path))
        worksheet = workbook.Worksheets("9grid")
        last_row = max(2, worksheet.UsedRange.Rows.Count)
        worksheet.Range(f"A2:I{last_row}").ClearContents()

        if not export_df.empty:
            rows = [list(row) for row in export_df.itertuples(index=False, name=None)]
            end_row = len(rows) + 1
            worksheet.Range(f"A2:I{end_row}").Value = rows
            worksheet.Range(f"H2:H{end_row}").FormulaR1C1 = (
                '=IFERROR(INDEX(HELP!R3C2:R11C2,MATCH(RC[-2]&"-"&RC[-1],HELP!R3C5:R11C5,0)),"")'
            )

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        if worksheet is not None:
            del worksheet
        if workbook is not None:
            del workbook
        del excel


def export_powerbi_workbooks(datasets: list[tuple[Path, str | None, pd.DataFrame]]) -> int:
    template_path = resolve_powerbi_template_path()
    export_date = date.today().isoformat()
    export_count = 0

    for input_path, manager_name, df in datasets:
        safe_manager = sanitize_filename(manager_name or input_path.stem)
        output_path = OUTPUT_DIR / f"9GRID_{safe_manager}.xlsb"
        export_df = build_powerbi_export_table(df, export_date)
        write_powerbi_workbook(template_path, output_path, export_df)
        export_count += 1

    return export_count


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

    def draw_manager_colors(start_x: float, start_y: float) -> float:
        y_pos = start_y
        ax.text(start_x, y_pos, "Manager colors", transform=ax.transAxes, fontsize=10, fontweight="bold", va="top")
        y_pos -= 0.04
        for owner, color in owner_colors.items():
            ax.text(start_x, y_pos, "\u25CF", color=color, transform=ax.transAxes, fontsize=12, va="top")
            ax.text(start_x + 0.04, y_pos, owner, transform=ax.transAxes, fontsize=9, va="top", color="#203040")
            y_pos -= 0.035
        return y_pos

    def draw_churn_risk(start_x: float, start_y: float) -> float:
        y_pos = start_y
        ax.text(
            start_x,
            y_pos,
            "Churn risk outline",
            transform=ax.transAxes,
            fontsize=10,
            fontweight="bold",
            va="top",
        )
        y_pos -= 0.04
        for label, key in [("Low", "low"), ("Medium", "medium"), ("High", "high")]:
            style = CHURN_RISK_STYLES[key]
            ax.scatter(
                [start_x + 0.015],
                [y_pos - 0.01],
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
                    [start_x + 0.015],
                    [y_pos - 0.01],
                    s=style["halo_size"] * 0.28,
                    facecolors="none",
                    edgecolors=style["edgecolor"],
                    linewidths=style["halo_width"],
                    transform=ax.transAxes,
                    clip_on=False,
                    zorder=2,
                )
            ax.text(start_x + 0.04, y_pos, label, transform=ax.transAxes, fontsize=9, va="top", color="#203040")
            y_pos -= 0.035
        return y_pos

    if is_overview:
        bottom_manager_y = draw_manager_colors(0.0, 0.98)
        bottom_churn_y = draw_churn_risk(0.5, 0.98)
        y = min(bottom_manager_y, bottom_churn_y) - 0.02
    else:
        y = draw_manager_colors(0.0, 0.98) - 0.01
        y = draw_churn_risk(0.0, y) - 0.02

    headers = list(legend_df.columns)
    table_rows = [headers]
    for _, row in legend_df.iterrows():
        table_rows.append([str(row[col]) for col in headers])

    if is_overview:
        col_widths = [0.07, 0.31]
        owner_count = max(1, len(headers) - 2)
        score_width = (1.0 - sum(col_widths)) / owner_count
        col_widths.extend([score_width] * owner_count)
    else:
        col_widths = [0.08, 0.38, 0.24, 0.30]

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


def create_chart(plotted_df: pd.DataFrame, title: str, png_path: Path) -> None:
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
        )
        owner_count += 1

    return owner_count


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    input_paths = resolve_input_paths()
    datasets: list[tuple[Path, str | None, pd.DataFrame]] = []
    for input_path in input_paths:
        manager_name = extract_manager_name(input_path)
        datasets.append((input_path, manager_name, load_data(input_path, manager_name=manager_name)))

    loaded_frames = [df for _, _, df in datasets]
    df = pd.concat(loaded_frames, ignore_index=True)
    if df.empty:
        raise ValueError("No valid rows found after applying the required filters.")

    plotted_df = compute_positions(df)
    create_chart(
        plotted_df,
        "9-Grid Talent Calibration",
        OUTPUT_DIR / "9grid_overview.png",
    )

    owner_views = export_owner_views(df)
    powerbi_exports = export_powerbi_workbooks(datasets)
    print(f"Loaded {len(input_paths)} workbook(s).")
    print(f"Created overview chart for {len(df)} employees.")
    print(f"Created {owner_views} owner-specific chart(s).")
    print(f"Created {powerbi_exports} Power BI workbook(s).")
    print(f"Files saved to: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()

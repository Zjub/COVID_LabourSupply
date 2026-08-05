from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(r"C:\Users\OEM\Documents\GitHub\COVID_LabourSupply\SAM")
OUT = ROOT / "presentations" / "assets"
MODEL = Path(r"C:\Users\OEM\Documents\GitHub\TVHENZ\e61 Projects\COVID Aust NZ\SAM")
EMP = Path(r"C:\Users\OEM\Documents\GitHub\TVHENZ\e61 Projects\COVID Aust NZ\initial_RDD")

TEAL = "#004f54"
AQUA = "#3aa7ac"
LIGHT_AQUA = "#99e8e8"
ORANGE = "#e86f3d"
GREY = "#6f7779"
LIGHT_GREY = "#e9eeee"
BLACK = "#111111"


def finish(fig: plt.Figure, name: str) -> None:
    fig.savefig(OUT / name, dpi=220, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def style_axes(ax: plt.Axes) -> None:
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#777777")
    ax.grid(axis="y", color="#dddddd", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=9, colors="#333333")


def empirical_series() -> None:
    jfr = pd.read_csv(EMP / "JFR_by_group 1 .csv")
    sep = pd.read_csv(EMP / "SR_by_group 1 .csv")
    for df in (jfr, sep):
        df["date"] = pd.to_datetime(df["date"])
        df["group"] = np.where(df["nz"].eq(1), "New Zealand citizen", "Australian")

    start, end = pd.Timestamp("2020-01-16"), pd.Timestamp("2020-07-16")
    fig, axes = plt.subplots(1, 2, figsize=(12.2, 4.4), sharex=True)
    for ax, df, title in zip(
        axes,
        (jfr, sep),
        ("Job-finding rate", "Exit from recorded payroll employment"),
    ):
        df = df[df["date"].between(start, end)].copy()
        for group, color in (("Australian", TEAL), ("New Zealand citizen", AQUA)):
            sub = df[df["group"].eq(group)].sort_values("date")
            ax.plot(sub["date"], 100 * sub["prop"], marker="o", markersize=3.2,
                    linewidth=2.0, color=color, label=group)
        ax.axvline(pd.Timestamp("2020-03-22"), color=ORANGE, linestyle="--", linewidth=1.6)
        ax.set_title(title, loc="left", fontsize=13, fontweight="bold", color=BLACK)
        ax.set_ylabel("Per cent per week", fontsize=10)
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
        style_axes(ax)
    axes[0].legend(frameon=False, fontsize=9, loc="upper right")
    fig.suptitle("Current cleared weekly series: informative patterns, selected risk sets",
                 x=0.055, ha="left", fontsize=16, fontweight="bold", color=TEAL)
    fig.text(0.055, -0.01,
             "Dashed line: 22 March announcement. These series inherit the draft's realised-receipt and future non-employment selection rules.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.04, 0.05, 1, 0.90))
    finish(fig, "empirical_weekly_series.png")


def empirical_cells() -> None:
    labels = ["NZ pre", "Australia pre", "NZ post", "Australia post"]
    jfr = [10.02, 8.68, 8.23, 5.18]
    sep = [4.80, 4.84, 5.50, 9.26]
    colors = [AQUA, TEAL, AQUA, TEAL]
    fig, axes = plt.subplots(1, 2, figsize=(11.8, 4.3))
    for ax, values, title, did in zip(
        axes,
        (jfr, sep),
        ("Job finding", "Payroll-employment exits"),
        ("DiD = −1.71 pp", "DiD = +3.72 pp"),
    ):
        bars = ax.bar(np.arange(4), values, color=colors, width=0.68)
        ax.set_xticks(np.arange(4), labels, rotation=18, ha="right")
        ax.set_ylabel("Per cent per week")
        ax.set_title(title, loc="left", fontsize=13, fontweight="bold")
        ax.text(0.98, 0.94, did, transform=ax.transAxes, ha="right", va="top",
                fontsize=12, fontweight="bold", color=ORANGE)
        for b, value in zip(bars, values):
            ax.text(b.get_x() + b.get_width()/2, value + 0.18, f"{value:.2f}",
                    ha="center", fontsize=9)
        style_axes(ax)
        ax.set_ylim(0, max(values) * 1.32)
    fig.suptitle("The draft's four empirical cells discipline model 15",
                 x=0.055, ha="left", fontsize=16, fontweight="bold", color=TEAL)
    fig.text(0.055, -0.015,
             "Source: manuscript Tables 3–4 / model-15 target vector. Percentage-point effects use Australian minus New Zealand changes.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.04, 0.05, 1, 0.90))
    finish(fig, "empirical_did_cells.png")


def model_fit() -> None:
    s = pd.read_csv(MODEL / "15_structural_separation_summary.csv")
    mapping = {
        "f_N_pre": "NZ pre", "f_R_pre": "AU pre",
        "f_N_post": "NZ post", "f_R_post": "AU post",
        "sep_N_pre": "NZ pre", "sep_R_pre": "AU pre",
        "sep_N_post": "NZ post", "sep_R_post": "AU post",
    }
    fig, axes = plt.subplots(1, 2, figsize=(12.0, 4.4))
    for ax, moments, title in (
        (axes[0], ["f_N_pre", "f_R_pre", "f_N_post", "f_R_post"], "Job-finding moments"),
        (axes[1], ["sep_N_pre", "sep_R_pre", "sep_N_post", "sep_R_post"], "Employment-exit moments"),
    ):
        sub = s.set_index("moment").loc[moments]
        x = np.arange(4)
        ax.bar(x - 0.18, 100 * sub["data"], 0.36, label="Data", color=GREY)
        ax.bar(x + 0.18, 100 * sub["model"], 0.36, label="Model 15", color=TEAL)
        ax.set_xticks(x, [mapping[m] for m in moments])
        ax.set_ylabel("Per cent per week")
        ax.set_title(title, loc="left", fontsize=13, fontweight="bold")
        style_axes(ax)
    axes[0].legend(frameon=False, fontsize=9, loc="upper right")
    fig.suptitle("Fit is close by construction—except the held-out Australian post JFR",
                 x=0.055, ha="left", fontsize=16, fontweight="bold", color=TEAL)
    fig.text(0.055, -0.01,
             "Benchmark predicts Australian post JFR 5.58% versus 5.18% in the draft (+0.40 pp). All four exit cells are calibration targets.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.04, 0.05, 1, 0.90))
    finish(fig, "model_fit.png")


def model_scenarios() -> None:
    d = pd.read_csv(MODEL / "15_structural_separation_decomposition.csv")
    order = ["benefit_only", "common_covid_only", "differential_covid_only", "full"]
    labels = ["Benefit\nonly", "Common COVID\nonly", "Differential wedge\nonly", "Full\nenvironment"]
    fig, axes = plt.subplots(1, 2, figsize=(12.2, 4.4))
    for ax, outcome, title in (
        (axes[0], "job_finding", "Job-finding DiD"),
        (axes[1], "separation", "Employment-exit DiD"),
    ):
        sub = d[d["outcome"].eq(outcome)].set_index("scenario").loc[order]
        values = sub["model_did_pp"].to_numpy()
        bars = ax.bar(np.arange(4), values, color=[AQUA, LIGHT_AQUA, ORANGE, TEAL], width=0.68)
        ax.axhline(sub["data_did_pp"].iloc[0], color=BLACK, linestyle="--", linewidth=1.5,
                   label="Draft data DiD")
        ax.axhline(0, color="#777777", linewidth=0.8)
        ax.set_xticks(np.arange(4), labels)
        ax.set_ylabel("Percentage points")
        ax.set_title(title, loc="left", fontsize=13, fontweight="bold")
        for b, value in zip(bars, values):
            va = "bottom" if value >= 0 else "top"
            ax.text(b.get_x()+b.get_width()/2, value + (0.08 if value >= 0 else -0.08),
                    f"{value:+.2f}", ha="center", va=va, fontsize=9)
        style_axes(ax)
    axes[0].legend(frameon=False, fontsize=9, loc="lower right")
    fig.suptitle("These are counterfactual environments—not additive shares",
                 x=0.055, ha="left", fontsize=16, fontweight="bold", color=TEAL)
    fig.text(0.055, -0.01,
             "Nonlinear interactions matter. For exits, full DiD 3.72 pp exceeds the sum of standalone effects by about 1.97 pp.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.04, 0.05, 1, 0.90))
    finish(fig, "model_scenarios.png")


def model_margins() -> None:
    d = pd.read_csv(MODEL / "15_structural_separation_scenarios.csv")
    d = d.set_index("scenario").loc[["benefit_only", "common_covid_only", "differential_covid_only", "full"]]
    labels = ["Benefit", "Common", "Differential", "Full"]
    x = np.arange(4)
    fig, axes = plt.subplots(1, 2, figsize=(12.0, 4.2))
    for ax, cols, title, ylim in (
        (axes[0], ("search_R", "search_N"), "Search effort (normalised)", (0.96, 1.005)),
        (axes[1], ("accept_R", "accept_N"), "Acceptance probability", (0.44, 0.78)),
    ):
        ax.plot(x, d[cols[0]], marker="o", linewidth=2.2, color=TEAL, label="Australian / R")
        ax.plot(x, d[cols[1]], marker="o", linewidth=2.2, color=AQUA, label="NZ / N")
        ax.set_xticks(x, labels)
        ax.set_ylim(*ylim)
        ax.set_title(title, loc="left", fontsize=13, fontweight="bold")
        style_axes(ax)
    axes[0].legend(frameon=False, fontsize=9, loc="lower left")
    fig.suptitle("In the benchmark, acceptance—not search—drives most JFR movement",
                 x=0.055, ha="left", fontsize=16, fontweight="bold", color=TEAL)
    fig.text(0.055, -0.01,
             "Selected search curvature η = 100 pins search near one; fixed acceptance thresholds then carry the behavioural response.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.04, 0.05, 1, 0.90))
    finish(fig, "model_margins.png")


def model_identification() -> None:
    d = pd.read_csv(MODEL / "15_structural_separation_restricted_grid.csv")
    near = d[d["fitted_loss"] < 0.01].copy()
    fig, ax = plt.subplots(figsize=(7.8, 4.8))
    sc = ax.scatter(near["heldout_AUS_post_jfr_gap_pp"], near["differential_health_pct_wage"],
                    c=near["search_curvature"], cmap="viridis", s=70,
                    edgecolor="white", linewidth=0.6)
    ax.axvline(0, color=ORANGE, linestyle="--", linewidth=1.4)
    ax.set_xlabel("Held-out Australian post-JFR prediction error (pp)")
    ax.set_ylabel("Differential work wedge (% of wage)")
    ax.set_title("Near-identical fit, very different mechanisms", loc="left",
                 fontsize=15, fontweight="bold", color=TEAL)
    style_axes(ax)
    cbar = fig.colorbar(sc, ax=ax, pad=0.02)
    cbar.set_label("Search curvature η", fontsize=9)
    fig.text(0.10, -0.01,
             "Near-fitting set: restricted-grid loss < 0.01. This is functional-form/set-identification uncertainty, not ordinary sampling error.",
             fontsize=8.5, color=GREY)
    fig.tight_layout(rect=(0.03, 0.05, 1, 1))
    finish(fig, "model_identification.png")


def replacement_validation() -> None:
    d = pd.read_csv(MODEL / "15_replacement_rate_gradient_validation.csv")
    x = np.arange(len(d))
    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    ax.plot(x, d["data_did_pp"], marker="o", linewidth=2.2, color=BLACK, label="Empirical income-bin DiD")
    ax.plot(x, d["model_full_did_pp"], marker="o", linewidth=2.2, color=TEAL, label="Model full")
    ax.plot(x, d["model_benefit_only_did_pp"], marker="o", linewidth=2.2, color=AQUA, label="Model benefit only")
    ax.set_xticks(x, d["group"])
    ax.set_ylabel("Job-finding DiD (pp)")
    ax.set_title("Do not use as validation yet", loc="left", fontsize=15, fontweight="bold", color=ORANGE)
    ax.text(0.02, 0.06,
            "Model rates (14/28%, 28/56%, 45/90%) are placeholders;\nempirical points are income bins, not entitlement-based replacement-rate bins.",
            transform=ax.transAxes, fontsize=10, color=ORANGE,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#fff2ec", edgecolor=ORANGE))
    ax.legend(frameon=False, fontsize=9, loc="upper right")
    style_axes(ax)
    fig.tight_layout()
    finish(fig, "replacement_rate_placeholder.png")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    empirical_series()
    empirical_cells()
    model_fit()
    model_scenarios()
    model_margins()
    model_identification()
    replacement_validation()
    print(f"Wrote deck figures to {OUT}")


if __name__ == "__main__":
    main()

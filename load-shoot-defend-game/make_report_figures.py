from __future__ import annotations

import json
import re
import subprocess
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parent
FONT_PATH = Path("/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf")
FONT = (
    FontProperties(fname=str(FONT_PATH))
    if FONT_PATH.exists()
    else FontProperties(family="DejaVu Serif")
)
SOURCE = ROOT / "explore_game.cpp"
BINARY = ROOT / "explore_game"


def build_explorer() -> None:
    """Compile the floating-point explorer when needed."""
    if BINARY.exists() and BINARY.stat().st_mtime >= SOURCE.stat().st_mtime:
        return
    subprocess.run(
        ["g++", "-std=c++17", "-O2", "-DNDEBUG", str(SOURCE), "-o", str(BINARY)],
        cwd=ROOT,
        text=True,
        check=True,
    )


def run_explorer(nmax: int = 80) -> subprocess.CompletedProcess[str]:
    build_explorer()
    return subprocess.run(
        [str(BINARY), str(nmax)],
        cwd=ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )


BOUNDS_RE = re.compile(
    r"^bounds\s+(?P<n>\d+)\s+(?P<lo52>[-+0-9.eE]+)\s+"
    r"(?P<hi52>[-+0-9.eE]+)\s+(?P<gap52>[-+0-9.eE]+)\s+"
    r"(?P<lo21>[-+0-9.eE]+)\s+(?P<hi21>[-+0-9.eE]+)\s+"
    r"(?P<gap21>[-+0-9.eE]+)$"
)


def parse_bounds(proc: subprocess.CompletedProcess[str]) -> list[dict[str, float]]:
    rows: list[dict[str, float]] = []
    for line in proc.stderr.splitlines():
        match = BOUNDS_RE.match(line.strip())
        if not match:
            continue
        rows.append(
            {
                "n": int(match.group("n")),
                "lo52": float(match.group("lo52")),
                "hi52": float(match.group("hi52")),
                "gap52": float(match.group("gap52")),
                "lo21": float(match.group("lo21")),
                "hi21": float(match.group("hi21")),
                "gap21": float(match.group("gap21")),
            }
        )
    if not rows:
        raise RuntimeError("explore_game did not produce bound data")
    return rows


SEQ_RE = re.compile(
    r"^seq\s+(?P<n>\d+)\s+(?P<a>[-+0-9.eE]+)\s+"
    r"(?P<b>[-+0-9.eE]+)\s+(?P<c>[-+0-9.eE]+)$"
)


def parse_zero_values(proc: subprocess.CompletedProcess[str]) -> list[dict[str, float]]:
    rows: list[dict[str, float]] = []
    for line in proc.stderr.splitlines():
        match = SEQ_RE.match(line.strip())
        if not match:
            continue
        n = int(match.group("n"))
        a = float(match.group("a"))
        b = float(match.group("b"))
        if n >= 3 and a > 0 and b > 0:
            den = 1.0 + a + b
            rows.append(
                {
                    "continuation_horizon": n,
                    "horizon_at_33": n + 1,
                    "A": a,
                    "B": b,
                    "L": b / den,
                    "S": a / den,
                    "D": 1.0 / den,
                }
            )
    if not rows:
        raise RuntimeError("explore_game did not produce sequence data")
    return rows


def style_axes(ax) -> None:
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(FONT)
        label.set_fontsize(9)
    ax.grid(True, which="major", linewidth=0.5, alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def make_gap_chart(rows: list[dict[str, float]]) -> None:
    ns = [row["n"] for row in rows]
    gap52 = [row["gap52"] for row in rows]
    gap21 = [row["gap21"] for row in rows]
    fig, ax = plt.subplots(figsize=(7.0, 4.1), constrained_layout=True)
    ax.semilogy(ns, gap52, marker="o", markersize=3.2, linewidth=1.6, label="G_N(5,2)")
    ax.semilogy(ns, gap21, marker="s", markersize=3.0, linewidth=1.6, label="G_N(2,1)")
    ax.set_xlabel("Cutoff horizon N", fontproperties=FONT, fontsize=10.5)
    ax.set_ylabel("Upper-lower gap (log scale)", fontproperties=FONT, fontsize=10.5)
    ax.set_title("Finite-horizon gaps at two key states", fontproperties=FONT, fontsize=13)
    style_axes(ax)
    ax.legend(prop=FONT, frameon=False)
    ax.text(
        0.01,
        0.015,
        "Double-precision exploration; the proof does not rely on this plot.",
        transform=ax.transAxes,
        fontproperties=FONT,
        fontsize=8.5,
        color="#555555",
    )
    fig.savefig(ROOT / "figure_gap.png", dpi=300, facecolor="white")
    plt.close(fig)


def make_strategy_chart(rows: list[dict[str, float]]) -> None:
    ns = [row["horizon_at_33"] for row in rows]
    fig, ax = plt.subplots(figsize=(7.0, 4.1), constrained_layout=True)
    for key, marker, label in (("L", "o", "Load L"), ("S", "s", "Shoot S"), ("D", "^", "Defend D")):
        ax.plot(ns, [row[key] for row in rows], marker=marker, markersize=3.0, linewidth=1.5, label=label)
    ax.set_xlabel("Remaining horizon at state (3,3)", fontproperties=FONT, fontsize=10.5)
    ax.set_ylabel("Finite-horizon equilibrium probability", fontproperties=FONT, fontsize=10.5)
    ax.set_ylim(0.20, 0.48)
    ax.set_title("Numerical trend of the three second-round actions", fontproperties=FONT, fontsize=13)
    style_axes(ax)
    ax.legend(prop=FONT, frameon=False, ncol=3, loc="upper right")
    ax.text(
        0.01,
        0.015,
        "Floating-point zero-cutoff values inserted into the exact formula.",
        transform=ax.transAxes,
        fontproperties=FONT,
        fontsize=8.5,
        color="#555555",
    )
    fig.savefig(ROOT / "figure_strategy.png", dpi=300, facecolor="white")
    plt.close(fig)


def main() -> None:
    proc = run_explorer(80)
    bounds = parse_bounds(proc)
    strategy = parse_zero_values(proc)
    make_gap_chart(bounds)
    make_strategy_chart(strategy)
    payload = {"bounds": bounds, "strategy": strategy}
    (ROOT / "report_data.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"wrote {len(bounds)} bound rows and {len(strategy)} strategy rows")


if __name__ == "__main__":
    main()

"""
Sample: UV-vis absorbance data visualization using origin_pro_support

Demonstrates:
  - Loading a CSV into a pandas DataFrame
  - Importing data into an Origin workbook
  - Creating a multi-line graph with custom colors, line styles, and widths
  - Creating a line+symbol graph for a single wavelength vs. time
  - Setting axis labels, tick spacings, legends, page size, and layer geometry

Run from the repository root:
    python sample/sample_UV-vis.py
"""
import sys
import os
import traceback

import pandas as pd

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, REPO_ROOT)

import origin_pro_support as ops
from origin_pro_support.layer.enums import (
    XYPlotType,
    AxisType,
    GroupMode,
    OriginColorIndex,
    LineStyle,
    LegendAnchor,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SAMPLE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH   = os.path.join(SAMPLE_DIR, "sample_data", "sample_UV-vis.csv")
PROJ_PATH  = os.path.join(SAMPLE_DIR, "sample_UV-vis.opju")

# Page size: 8 inch x 6 inch
# set_width / set_height accept values in inches.
PAGE_WIDTH  = 8.0   # inches
PAGE_HEIGHT = 6.0   # inches

# Layer geometry (% of page)
LAYER_LEFT   = 15
LAYER_TOP    =  5
LAYER_WIDTH  = 80
LAYER_HEIGHT = 80


def main() -> int:
    if os.path.exists(PROJ_PATH):
        os.remove(PROJ_PATH)
        print(f"Removed existing project: {PROJ_PATH}")

    print("=== Starting Origin ===")
    origin = ops.OriginInstance(PROJ_PATH)
    try:
        _build(origin)
        print("\n=== Done ===")
        return 0
    except Exception as exc:
        print(f"\n=== FAILED: {exc} ===")
        traceback.print_exc()
        return 1
    finally:
        origin.save()
        origin.close()
        print("Origin closed.")


def _build(origin: ops.OriginInstance) -> None:
    # -----------------------------------------------------------------------
    # Load CSV
    # -----------------------------------------------------------------------
    df = pd.read_csv(CSV_PATH)
    print(f"Loaded CSV: {CSV_PATH}  shape={df.shape}")

    # Column names: "wavelength (nm)", "0", "50", "100", ...
    wavelength_col = df.columns[0]   # "wavelength (nm)"

    # -----------------------------------------------------------------------
    # Graph 1: UV-vis spectra (wavelength vs. absorbance for 0, 50, 100 min)
    # -----------------------------------------------------------------------
    _make_graph1(origin, df, wavelength_col)

    # -----------------------------------------------------------------------
    # Graph 2: Absorbance at 600 nm vs. time
    # -----------------------------------------------------------------------
    _make_graph2(origin, df, wavelength_col)


# ---------------------------------------------------------------------------
# Graph 1 helpers
# ---------------------------------------------------------------------------

def _make_graph1(
    origin: ops.OriginInstance,
    df: pd.DataFrame,
    wavelength_col: str,
) -> None:
    """Create the UV-vis spectra graph (wavelength on X, three time points)."""

    # --- Workbook / worksheet -----------------------------------------------
    wbook = origin.new_workbook("UVvis_Spectra")
    ws    = wbook.get_layer(0)

    # Load only the columns we need: wavelength, 0, 50, 100
    sub = df[[wavelength_col, "0", "50", "100"]].copy()
    ws.add_column_from_data(sub)

    # Column indices (0-based).
    # Worksheet.__init__ pre-creates 2 empty columns (index 0, 1).
    # add_column_from_data appends starting at the current column count,
    # so the DataFrame's 4 columns land at indices 2, 3, 4, 5.
    col_wavelength = 2
    col_0   = 3
    col_50  = 4
    col_100 = 5

    # --- Graph page ---------------------------------------------------------
    gpage = origin.new_graph("UVvis_Spectra_Graph", XYPlotType.LINE)
    gpage.set_width(PAGE_WIDTH)
    gpage.set_height(PAGE_HEIGHT)

    layer = gpage.get_layer(0)
    
    # --- Add plots with colors set at creation time -------------------------
    # Passing color= at creation avoids the need to ungroup plots afterwards.
    layer.add_xy_plot(ws, col_wavelength, col_0,   XYPlotType.LINE, color=OriginColorIndex.BLACK)
    layer.add_xy_plot(ws, col_wavelength, col_50,  XYPlotType.LINE, color=OriginColorIndex.RED)
    layer.add_xy_plot(ws, col_wavelength, col_100, XYPlotType.LINE, color=OriginColorIndex.BLUE)

    # Ungroup before setting per-plot styles; Origin auto-groups when multiple
    # plots are added, which causes set commands to be overridden by the group.
    layer.group_plots(GroupMode.NONE)
    plots = layer.data_plots
    plot_0, plot_50, plot_100 = plots[0], plots[1], plots[2]

    # --- Line styles --------------------------------------------------------
    plot_0.line_style   = LineStyle.DOT    # dotted
    plot_50.line_style  = LineStyle.SOLID  # solid
    plot_100.line_style = LineStyle.SOLID  # solid

    # --- Line widths (2 pt each) --------------------------------------------
    plot_0.line_width   = 2.0
    plot_50.line_width  = 2.0
    plot_100.line_width = 2.0

    # --- Axis labels --------------------------------------------------------
    ax_x = layer.get_axis(AxisType.X)
    ax_y = layer.get_axis(AxisType.Y)
    ax_x.label_text = "Wavelength [nm]"
    ax_y.label_text = "Absorbance [-]"

    # --- Set layer geometry -------------------------------------------------
    layer.set_scale(LAYER_LEFT, LAYER_TOP, LAYER_WIDTH, LAYER_HEIGHT)

    # --- Axis tick spacing --------------------------------------------------
    ax_x.set_major_tick_spacing(200)
    ax_x.set_minor_tick_count(1)

    # --- Legend: custom labels, upper-left position -------------------------
    legend = layer.get_legend()
    legend.reconstruct()
    # Override text: one entry per line.
    # Use Python \n (converted to CRLF internally) — NOT raw-string \n.
    legend.text = r"\l(1) 0 min" + "\n" + r"\l(2) 50 min" + "\n" + r"\l(3) 100 min"

    # Place legend so its top-left corner coincides with the layer top-left (0%, 100%).
    legend.set_position_pct(0, 100, anchor=LegendAnchor.TOP_LEFT)

    print("Graph 1 created: UVvis_Spectra_Graph")


# ---------------------------------------------------------------------------
# Graph 2 helpers
# ---------------------------------------------------------------------------

def _make_graph2(
    origin: ops.OriginInstance,
    df: pd.DataFrame,
    wavelength_col: str,
) -> None:
    """Create the absorbance-at-600-nm vs. time graph."""

    # --- Prepare data -------------------------------------------------------
    # Transpose: index becomes time points (column headers of original df)
    # Filter to wavelength == 600
    row_600 = df[df[wavelength_col] == 600]
    if row_600.empty:
        raise ValueError("No row found with wavelength == 600 in the CSV.")

    # Drop the wavelength column, transpose so time is the index
    # Original time columns: "0", "50", "100", "150", ...
    time_values_str = [c for c in df.columns if c != wavelength_col]
    time_values = [float(t) for t in time_values_str]
    abs_values  = [float(row_600[t].iloc[0]) for t in time_values_str]

    # --- Workbook / worksheet -----------------------------------------------
    wbook = origin.new_workbook("UVvis_600nm")
    ws    = wbook.get_layer(0)

    ws.add_column_from_data(time_values,  lname="Time")
    ws.add_column_from_data(abs_values,   lname="Absorbance")

    # Same offset: Worksheet.__init__ pre-creates 2 empty columns.
    # First add_column_from_data call → index 2, second → index 3.
    col_time = 2
    col_abs  = 3

    # --- Graph page ---------------------------------------------------------
    gpage = origin.new_graph("UVvis_600nm_Graph", XYPlotType.LINE_SYMBOL)
    gpage.set_width(PAGE_WIDTH)
    gpage.set_height(PAGE_HEIGHT)

    layer = gpage.get_layer(0)
    
    # --- Add plot with color set at creation time --------------------------
    layer.add_xy_plot(ws, col_time, col_abs, XYPlotType.LINE_SYMBOL,
                      color=(187, 51, 187))

    # Ungroup before setting per-plot styles.
    layer.group_plots(GroupMode.NONE)
    plot = layer.data_plots[0]

    # --- Symbol size (12 pt) and line width (2 pt) --------------------------
    plot.symbol_size = 12.0
    plot.line_width  = 2.0

    # --- Set layer geometry -------------------------------------------------
    layer.set_scale(LAYER_LEFT, LAYER_TOP, LAYER_WIDTH, LAYER_HEIGHT)

    # --- Axis labels --------------------------------------------------------
    ax_x = layer.get_axis(AxisType.X)
    ax_y = layer.get_axis(AxisType.Y)
    ax_x.label_text = "Time [min]"
    ax_y.label_text = "Absorbance [-]"

    # --- Axis tick spacing --------------------------------------------------
    ax_x.set_major_tick_spacing(100)
    ax_x.set_minor_tick_count(1)

    print("Graph 2 created: UVvis_600nm_Graph")


if __name__ == "__main__":
    sys.exit(main())

import os
import tempfile
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib import patches
import itertools

def draw_page_frame_and_title_block(fig, inner_left, inner_bottom, inner_w, inner_h,
                                    rapport_nr, figur_nr, tegn, kontr, godkj, dato,
                                    logo_path):
    # Page border
    fig.patches.append(patches.Rectangle(
        (inner_left, inner_bottom), inner_w, inner_h,
        transform=fig.transFigure, fill=False, linewidth=2.0, edgecolor="black"
    ))

    # Title block (same geometry you approved)
    tb_width  = inner_w * 0.30
    tb_height = inner_h * (0.26 - 0.04)
    tb_left   = inner_left + inner_w - tb_width
    tb_bottom = inner_bottom

    fig.patches.append(patches.Rectangle(
        (tb_left, tb_bottom), tb_width, tb_height,
        transform=fig.transFigure, fill=False, linewidth=1.5, edgecolor="black"
    ))

    # Rows grid (≈45%) + logo (≈55%)
    logo_h_frac = 0.55
    grid_height = tb_height * (1 - logo_h_frac)
    row_h = grid_height / 3.0
    r1_top = tb_bottom + tb_height
    r2_top = r1_top - row_h
    r3_top = r2_top - row_h
    r4_top = r3_top - row_h  # top of logo area

    for y in [r2_top, r3_top, r4_top]:
        fig.lines.append(plt.Line2D([tb_left, tb_left + tb_width], [y, y],
                                    transform=fig.transFigure, linewidth=1.0, color="black"))

    v_r1 = tb_left + tb_width * (2/3)
    fig.lines.append(plt.Line2D([v_r1, v_r1], [r2_top, r1_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))
    v2_1 = tb_left + tb_width/3
    v2_2 = tb_left + 2*tb_width/3
    fig.lines.append(plt.Line2D([v2_1, v2_1], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))
    fig.lines.append(plt.Line2D([v2_2, v2_2], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    def ftxt(x, y, s, ha='left', va='center', size=8, weight=None):
        fig.text(x, y, s, ha=ha, va=va, fontsize=size, fontweight=weight)

    pad = 0.008
    # Row 1
    ftxt(tb_left + pad, r1_top - pad, f"Rapport Nr.: {rapport_nr}", va='top', size=9, weight='bold')
    ftxt(v_r1   + pad, r1_top - pad, f"Figur Nr.: {figur_nr}",      va='top', size=9, weight='bold')
    # Row 2
    ftxt(tb_left + pad, r2_top - pad, f"Tegn: {tegn}",  va='top', size=9)
    ftxt(v2_1   + pad, r2_top - pad, f"Kontr: {kontr}", va='top', size=9)
    ftxt(v2_2   + pad, r2_top - pad, f"Godkj: {godkj}", va='top', size=9)
    # Row 3
    ftxt(tb_left + pad, r3_top - pad, f"Dato: {dato}",   va='top', size=9)

    # Logo (bottom) — bottom-right, no stretching
    area_left   = tb_left + pad
    area_right  = tb_left + tb_width - pad
    area_bottom = tb_bottom + pad
    area_top    = r4_top - pad
    area_w = area_right - area_left
    area_h = area_top - area_bottom
    if logo_path and os.path.exists(logo_path):
        try:
            img = mpimg.imread(logo_path)
            h, w = img.shape[0], img.shape[1]
            aspect = w / h
            width_eff  = min(area_w, area_h * aspect)
            height_eff = width_eff / aspect
            logo_left   = area_left + (area_w - width_eff)
            logo_bottom = area_bottom
            logo_ax = fig.add_axes([logo_left, logo_bottom, width_eff, height_eff])
            logo_ax.imshow(img); logo_ax.axis('off')
        except Exception:
            fig.text(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', fontsize=14, fontweight='bold')
    else:
        fig.text(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', fontsize=14, fontweight='bold')

    return tb_left, tb_bottom, tb_width, tb_height

def add_box_spines(ax):
    for s in ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

def export_curfc_pdf(
    data_series,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    depth_ylim=(0, 35),
    margin_cm=1.0
    ):
   
    # Defaults
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C3")

    # Figure: A4 landscape (inches)
    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    # 1 cm margin -> figure fractions
    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    tb_left, tb_bottom, tb_width, tb_height = draw_page_frame_and_title_block(
        fig, inner_left, inner_bottom, inner_w, inner_h,
        rapport_nr, figur_nr, tegn, kontr, godkj, dato, logo_path
    )
   
    # Bottom fixed 0.3 cm above title block top
    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h
    # Lower the top so titles/ticks sit inside page border
    charts_top = inner_top - 0.07
    charts_height = max(0.05, charts_top - charts_bottom)

    # Axes: LEFT depth, RIGHT elevation. Leave a bit of inner margins so labels/ticks stay inside border.
    left_ax = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    # X-axis formatting (as in your good version)
    def setup_xaxis(ax):
        ax.set_xscale('log')
        ax.set_xticks([0.1,0.2,0.5,1,2,5,10,20,30])
        ax.set_xticklabels(['0.1','0.2','0.5','1','2','5','10','20','30'])
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='both', linewidth=0.5, alpha=0.4)

    # Colors/markers cycle for series
    colors = ['b','g','r','c','m','y','k']
    markers = ['o','x','s','^']
    cmb = itertools.cycle([(c, m) for c in colors for m in markers])
    style_map = {name: next(cmb) for name, *_ in data_series}

    # LEFT: Depth vs strength
    for workbook_name, x_vals, y_vals, elev_vals, terrain_level in data_series:
        c, m = style_map[workbook_name]
        left_ax.scatter(x_vals, y_vals, label=f"{workbook_name}, {terrain_level:.1f} m",
                        color=c, marker=m, s=25)
    left_ax.set_xlabel("Omrørt skjærstyrke (kPa)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim)
    left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax)
    for s in left_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # RIGHT: Elevation vs strength
    for workbook_name, x_vals, y_vals, elev_vals, terrain_level in data_series:
        c, m = style_map[workbook_name]
        right_ax.scatter(x_vals, elev_vals, color=c, marker=m, s=25)
    right_ax.set_xlabel("Omrørt skjærstyrke (kPa)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax)
    for s in right_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")
        
        # --- Vertikal referanselinje (x = 1.27) ---
    left_ax.axvline(1.27, color='k', linestyle='--', linewidth=1.0, zorder=3)

    # ===== Legend: 4 columns, vertically centered to title block =====
    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)

    handles, labels = left_ax.get_legend_handles_labels()
    if handles:
        fig.legend(
            handles, labels,
            loc='upper left',
            bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
            bbox_transform=fig.transFigure,
            ncol=4, frameon=True, fontsize=8,
            columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6
        )

    # Save
    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png:
        fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

def export_density_pdf(
    series_density,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    xlim=(15, 25),
    depth_ylim=(0, 35),
    margin_cm=1.0
):
    """
    A4 landscape, same layout as previous figures:
      - 1 cm page border
      - Left: depth vs density (y left, inverted), boxed spines
      - Right: elevation vs density (y right), boxed spines
      - Bottom fixed 0.3 cm above title block; top lowered so labels inside border
      - Legend: boreholes (dedup), 4 columns, vertically centered to title block
      - Title block layout:
          Row 1: 2/3 Rapport Nr., 1/3 Figur Nr.
          Row 2: 3 equal (Tegn, Kontr, Godkj)
          Row 3: full width (Dato)
          Bottom: Logo (taller), bottom-right, no stretching
    """
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C5")
    
    # Figure: A4 landscape (inches)
    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    # 1 cm margin -> figure fractions
    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    # Page border
    fig.patches.append(patches.Rectangle(
        (inner_left, inner_bottom), inner_w, inner_h,
        transform=fig.transFigure, fill=False, linewidth=2.0, edgecolor="black"
    ))

    # ---- TITLE BLOCK ----
    tb_width  = inner_w * 0.30
    tb_height = inner_h * (0.26 - 0.04)
    tb_left   = inner_left + inner_w - tb_width
    tb_bottom = inner_bottom  # overlap page border

    # Frame
    fig.patches.append(patches.Rectangle(
        (tb_left, tb_bottom), tb_width, tb_height,
        transform=fig.transFigure, fill=False, linewidth=1.5, edgecolor="black"
    ))

    # Rows: compact grid (≈45%) + logo (≈55%)
    logo_h_frac = 0.55
    grid_height = tb_height * (1 - logo_h_frac)
    row_h = grid_height / 3.0
    r1_top = tb_bottom + tb_height
    r2_top = r1_top - row_h
    r3_top = r2_top - row_h
    r4_top = r3_top - row_h  # top of logo area

    # Horizontal lines
    for y in [r2_top, r3_top, r4_top]:
        fig.lines.append(plt.Line2D([tb_left, tb_left + tb_width], [y, y],
                                    transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 1: 2/3 + 1/3
    v_r1 = tb_left + tb_width * (2/3)
    fig.lines.append(plt.Line2D([v_r1, v_r1], [r2_top, r1_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 2: 3 equal
    v2_1 = tb_left + tb_width/3
    v2_2 = tb_left + 2*tb_width/3
    fig.lines.append(plt.Line2D([v2_1, v2_1], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))
    fig.lines.append(plt.Line2D([v2_2, v2_2], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Text helper
    def ftxt(x, y, s, ha='left', va='center', size=8, weight=None):
        fig.text(x, y, s, ha=ha, va=va, fontsize=size, fontweight=weight)

    pad = 0.008
    ftxt(tb_left + pad, r1_top - pad, f"Rapport Nr.: {rapport_nr}", va='top', size=9, weight='bold')
    ftxt(v_r1   + pad, r1_top - pad, f"Figur Nr.: {figur_nr}",    va='top', size=9, weight='bold')
    ftxt(tb_left + pad, r2_top - pad, f"Tegn: {tegn}",            va='top', size=9)
    ftxt(v2_1   + pad, r2_top - pad, f"Kontr: {kontr}",           va='top', size=9)
    ftxt(v2_2   + pad, r2_top - pad, f"Godkj: {godkj}",           va='top', size=9)
    ftxt(tb_left + pad, r3_top - pad, f"Dato: {dato}",            va='top', size=9)

    # Logo (bottom) — bottom-right, no stretching
    area_left   = tb_left + pad
    area_right  = tb_left + tb_width - pad
    area_bottom = tb_bottom + pad
    area_top    = r4_top - pad
    area_w = area_right - area_left
    area_h = area_top - area_bottom
    if logo_path and os.path.exists(logo_path):
        try:
            img = mpimg.imread(logo_path)
            h, w = img.shape[0], img.shape[1]
            aspect = w / h
            width_eff  = min(area_w, area_h * aspect)
            height_eff = width_eff / aspect
            logo_left   = area_left + (area_w - width_eff)
            logo_bottom = area_bottom
            logo_ax = fig.add_axes([logo_left, logo_bottom, width_eff, height_eff])
            logo_ax.imshow(img); logo_ax.axis('off')
        except Exception:
            ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')
    else:
        ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')

    # ===== Charts =====
    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h  # fixed
    charts_top = inner_top - 0.07                               # keep titles/ticks inside border
    charts_height = max(0.05, charts_top - charts_bottom)

    # Axes (leave inner margins so labels/ticks stay inside border)
    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    # X formatting (linear, 15–25, major ticks = 1)
    def setup_xaxis(ax):
        ax.set_xlim(*xlim)
        ax.xaxis.set_major_locator(MultipleLocator(1))
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='major', linewidth=0.5, alpha=0.4)

    # Colour+marker cycling (like your original script)
    colors = ['b','g','r','c','m','y','k']
    markers = ['o','x','s','^']
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # LEFT: Depth vs Tyngdetetthet
    for bh, xs, ds, elevs, Z in series_density:
        c,m = next(color_marker_combos)
        left_ax.scatter(xs, ds, color=c, marker=m, s=25, label=f"{bh}, {Z:.1f} m")
    left_ax.set_xlabel("Tyngdetetthet (kN/m$^3$)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim)
    left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax)
    for s in left_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # Reset cycle for right plot (so colours/markers match)
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # RIGHT: Elevation vs Tyngdetetthet
    for bh, xs, ds, elevs, Z in series_density:
        c,m = next(color_marker_combos)
        right_ax.scatter(xs, elevs, color=c, marker=m, s=25)
    right_ax.set_xlabel("Tyngdetetthet (kN/m$^3$)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax)
    for s in right_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # ===== Legend (no title) =====
    handles, labels = left_ax.get_legend_handles_labels()
    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)

    if handles:
        fig.legend(
            handles, labels,
            loc='upper left',
            bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
            bbox_transform=fig.transFigure,
            ncol=4, frameon=True, fontsize=8,
            columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6
        )

    # Save
    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png:
        fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

def export_cu_enaks_konus_pdf(
    series_konus,
    series_enaks,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    depth_ylim=(0, 35),
    margin_cm=1.0
):
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C4")

    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    tb_left, tb_bottom, tb_width, tb_height = draw_page_frame_and_title_block(
        fig, inner_left, inner_bottom, inner_w, inner_h,
        rapport_nr, figur_nr, tegn, kontr, godkj, dato, logo_path
    )

    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h
    charts_top = inner_top - 0.07
    charts_height = max(0.05, charts_top - charts_bottom)

    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    def setup_xaxis(ax):
        if USE_LOG_X:
            ax.set_xscale('log')
            ax.set_xticks([0.1,0.2,0.5,1,2,5,10,20,30])
            ax.set_xticklabels(['0.1','0.2','0.5','1','2','5','10','20','30'])
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='both' if USE_LOG_X else 'major', linewidth=0.5, alpha=0.4)

    for bh, xs, ds, elevs, Z in series_konus:
        left_ax.scatter(xs, ds, color=bh_color[bh], marker='^', s=25, label=f"{bh}, {Z:.1f} m")
    for bh, xs, ds, elevs, Z in series_enaks:
        left_ax.scatter(xs, ds, color=bh_color[bh], marker='o', s=25)
    left_ax.set_xlabel("Direkte skjærstyrke (kPa)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim); left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax); add_box_spines(left_ax)

    for bh, xs, ds, elevs, Z in series_konus:
        right_ax.scatter(xs, elevs, color=bh_color[bh], marker='^', s=25)
    for bh, xs, ds, elevs, Z in series_enaks:
        right_ax.scatter(xs, elevs, color=bh_color[bh], marker='o', s=25)
    right_ax.set_xlabel("Direkte skjærstyrke (kPa)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax); add_box_spines(right_ax)

    # Legend (no title), 4 cols
    bh_to_Z = {}
    for bh, *_ , Z in series_konus: bh_to_Z[bh] = Z
    for bh, *_ , Z in series_enaks: bh_to_Z[bh] = Z
    handles, labels = [], []
    for bh in sorted(bh_to_Z.keys()):
        handles.append(plt.Line2D([], [], linestyle='', marker='s', markersize=8, color=bh_color[bh]))
        labels.append(f"{bh}, {bh_to_Z[bh]:.1f} m")

    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)
    if handles:
        fig.legend(handles, labels, loc='upper left',
                   bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
                   bbox_transform=fig.transFigure, ncol=4, frameon=True, fontsize=8,
                   columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6)

    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png: fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

def export_wc_pdf(
    series_wc,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    xlim=(0, 100),
    depth_ylim=(0, 35),
    margin_cm=1.0
):
    """
    A4 landscape:
      - 1 cm page border
      - Left: depth vs water content (y left, inverted), boxed spines
      - Right: elevation vs water content (y right), boxed spines
      - Bottom fixed 0.3 cm above title block; top lowered so labels inside border
      - Legend: boreholes (no title), 4 columns, vertically centered to title block
      - Title block layout:
          Row 1: 2/3 Rapport Nr., 1/3 Figur Nr.
          Row 2: 3 equal (Tegn, Kontr, Godkj)
          Row 3: full width (Dato)
          Bottom: Logo (taller), bottom-right, no stretching
      - Colour+marker cycling (matches your original script)
    """
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C7")

    # Figure: A4 landscape (inches)
    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    # 1 cm margin -> figure fractions
    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    # Page border
    fig.patches.append(patches.Rectangle(
        (inner_left, inner_bottom), inner_w, inner_h,
        transform=fig.transFigure, fill=False, linewidth=2.0, edgecolor="black"
    ))

    # ---- TITLE BLOCK ----
    tb_width  = inner_w * 0.30
    tb_height = inner_h * (0.26 - 0.04)
    tb_left   = inner_left + inner_w - tb_width
    tb_bottom = inner_bottom  # overlap page border

    # Frame
    fig.patches.append(patches.Rectangle(
        (tb_left, tb_bottom), tb_width, tb_height,
        transform=fig.transFigure, fill=False, linewidth=1.5, edgecolor="black"
    ))

    # Rows: compact grid (≈45%) + logo (≈55%)
    logo_h_frac = 0.55
    grid_height = tb_height * (1 - logo_h_frac)
    row_h = grid_height / 3.0
    r1_top = tb_bottom + tb_height
    r2_top = r1_top - row_h
    r3_top = r2_top - row_h
    r4_top = r3_top - row_h  # top of logo area

    # Horizontal lines
    for y in [r2_top, r3_top, r4_top]:
        fig.lines.append(plt.Line2D([tb_left, tb_left + tb_width], [y, y],
                                    transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 1: 2/3 + 1/3
    v_r1 = tb_left + tb_width * (2/3)
    fig.lines.append(plt.Line2D([v_r1, v_r1], [r2_top, r1_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 2: 3 equal
    v2_1 = tb_left + tb_width/3
    v2_2 = tb_left + 2*tb_width/3
    fig.lines.append(plt.Line2D([v2_1, v2_1], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))
    fig.lines.append(plt.Line2D([v2_2, v2_2], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Text helper
    def ftxt(x, y, s, ha='left', va='center', size=8, weight=None):
        fig.text(x, y, s, ha=ha, va=va, fontsize=size, fontweight=weight)

    pad = 0.008
    ftxt(tb_left + pad, r1_top - pad, f"Rapport Nr.: {rapport_nr}", va='top', size=9, weight='bold')
    ftxt(v_r1   + pad, r1_top - pad, f"Figur Nr.: {figur_nr}",    va='top', size=9, weight='bold')
    ftxt(tb_left + pad, r2_top - pad, f"Tegn: {tegn}",            va='top', size=9)
    ftxt(v2_1   + pad, r2_top - pad, f"Kontr: {kontr}",           va='top', size=9)
    ftxt(v2_2   + pad, r2_top - pad, f"Godkj: {godkj}",           va='top', size=9)
    ftxt(tb_left + pad, r3_top - pad, f"Dato: {dato}",            va='top', size=9)

    # Logo (bottom) — bottom-right, no stretching
    area_left   = tb_left + pad
    area_right  = tb_left + tb_width - pad
    area_bottom = tb_bottom + pad
    area_top    = r4_top - pad
    area_w = area_right - area_left
    area_h = area_top - area_bottom
    if logo_path and os.path.exists(logo_path):
        try:
            img = mpimg.imread(logo_path)
            h, w = img.shape[0], img.shape[1]
            aspect = w / h
            width_eff  = min(area_w, area_h * aspect)
            height_eff = width_eff / aspect
            logo_left   = area_left + (area_w - width_eff)
            logo_bottom = area_bottom
            logo_ax = fig.add_axes([logo_left, logo_bottom, width_eff, height_eff])
            logo_ax.imshow(img); logo_ax.axis('off')
        except Exception:
            ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')
    else:
        ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')

    # ===== Charts =====
    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h  # fixed
    charts_top = inner_top - 0.07                               # keep titles/ticks inside border
    charts_height = max(0.05, charts_top - charts_bottom)

    # Axes (leave inner margins so labels/ticks stay inside border)
    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    # X formatting (linear, 0–100, major ticks = 10)
    def setup_xaxis(ax):
        ax.set_xlim(*xlim)
        ax.xaxis.set_major_locator(MultipleLocator(10))
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='major', linewidth=0.5, alpha=0.4)

    # Colour+marker cycling (like your script)
    colors  = ['b','g','r','c','m','y','k']
    markers = ['o','x','s','^']
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # LEFT: Depth vs Vanninnhold
    for bh, xs, ds, elevs, Z in series_wc:
        c, m = next(color_marker_combos)
        left_ax.scatter(xs, ds, color=c, marker=m, s=25, label=f"{bh}, {Z:.1f} m")
    left_ax.set_xlabel("Vanninnhold (%)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim)
    left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax)
    for s in left_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # Reset cycle so right plot matches colours/markers
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # RIGHT: Elevation vs Vanninnhold
    for bh, xs, ds, elevs, Z in series_wc:
        c, m = next(color_marker_combos)
        right_ax.scatter(xs, elevs, color=c, marker=m, s=25)
    right_ax.set_xlabel("Vanninnhold (%)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax)
    for s in right_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # ===== Legend (no title) =====
    handles, labels = left_ax.get_legend_handles_labels()
    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)

    if handles:
        fig.legend(
            handles, labels,
            loc='upper left',
            bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
            bbox_transform=fig.transFigure,
            ncol=4, frameon=True, fontsize=8,
            columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6
        )

    # Save
    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png:
        fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

def export_ip_pdf(
    series_ip,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    xlim=(0, 25),
    depth_ylim=(0, 35),
    margin_cm=1.0
):
    """
    A4 landscape:
      - 1 cm page border
      - Left: depth vs Ip (y left, inverted), boxed spines
      - Right: elevation vs Ip (y right), boxed spines
      - Bottom fixed 0.3 cm above title block; top lowered so labels inside border
      - Legend: boreholes (dedup), 4 columns, vertically centered to title block
      - Title block:
          Row 1: 2/3 Rapport Nr., 1/3 Figur Nr.
          Row 2: 3 equal (Tegn, Kontr, Godkj)
          Row 3: full width (Dato)
          Bottom: Logo (taller), bottom-right, no stretching
      - Colour+marker cycling like your original script
    """
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C6")

    # Figure: A4 landscape (inches)
    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    # 1 cm margin -> figure fractions
    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    # Page border
    fig.patches.append(patches.Rectangle(
        (inner_left, inner_bottom), inner_w, inner_h,
        transform=fig.transFigure, fill=False, linewidth=2.0, edgecolor="black"
    ))

    # ---- TITLE BLOCK ----
    tb_width  = inner_w * 0.30
    tb_height = inner_h * (0.26 - 0.04)
    tb_left   = inner_left + inner_w - tb_width
    tb_bottom = inner_bottom  # overlap page border

    # Frame
    fig.patches.append(patches.Rectangle(
        (tb_left, tb_bottom), tb_width, tb_height,
        transform=fig.transFigure, fill=False, linewidth=1.5, edgecolor="black"
    ))

    # Rows: compact grid (≈45%) + logo (≈55%)
    logo_h_frac = 0.55
    grid_height = tb_height * (1 - logo_h_frac)
    row_h = grid_height / 3.0
    r1_top = tb_bottom + tb_height
    r2_top = r1_top - row_h
    r3_top = r2_top - row_h
    r4_top = r3_top - row_h  # top of logo area

    # Horizontal lines
    for y in [r2_top, r3_top, r4_top]:
        fig.lines.append(plt.Line2D([tb_left, tb_left + tb_width], [y, y],
                                    transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 1: 2/3 + 1/3
    v_r1 = tb_left + tb_width * (2/3)
    fig.lines.append(plt.Line2D([v_r1, v_r1], [r2_top, r1_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Row 2: 3 equal
    v2_1 = tb_left + tb_width/3
    v2_2 = tb_left + 2*tb_width/3
    fig.lines.append(plt.Line2D([v2_1, v2_1], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))
    fig.lines.append(plt.Line2D([v2_2, v2_2], [r3_top, r2_top],
                                transform=fig.transFigure, linewidth=1.0, color="black"))

    # Text helper
    def ftxt(x, y, s, ha='left', va='center', size=8, weight=None):
        fig.text(x, y, s, ha=ha, va=va, fontsize=size, fontweight=weight)

    pad = 0.008
    ftxt(tb_left + pad, r1_top - pad, f"Rapport Nr.: {rapport_nr}", va='top', size=9, weight='bold')
    ftxt(v_r1   + pad, r1_top - pad, f"Figur Nr.: {figur_nr}",    va='top', size=9, weight='bold')
    ftxt(tb_left + pad, r2_top - pad, f"Tegn: {tegn}",            va='top', size=9)
    ftxt(v2_1   + pad, r2_top - pad, f"Kontr: {kontr}",           va='top', size=9)
    ftxt(v2_2   + pad, r2_top - pad, f"Godkj: {godkj}",           va='top', size=9)
    ftxt(tb_left + pad, r3_top - pad, f"Dato: {dato}",            va='top', size=9)

    # Logo (bottom) — bottom-right, no stretching
    area_left   = tb_left + pad
    area_right  = tb_left + tb_width - pad
    area_bottom = tb_bottom + pad
    area_top    = r4_top - pad
    area_w = area_right - area_left
    area_h = area_top - area_bottom
    if logo_path and os.path.exists(logo_path):
        try:
            img = mpimg.imread(logo_path)
            h, w = img.shape[0], img.shape[1]
            aspect = w / h
            width_eff  = min(area_w, area_h * aspect)
            height_eff = width_eff / aspect
            logo_left   = area_left + (area_w - width_eff)
            logo_bottom = area_bottom
            logo_ax = fig.add_axes([logo_left, logo_bottom, width_eff, height_eff])
            logo_ax.imshow(img); logo_ax.axis('off')
        except Exception:
            ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')
    else:
        ftxt(area_right, area_bottom, "GEOVITA", ha='right', va='bottom', size=14, weight='bold')

    # ===== Charts =====
    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h  # fixed
    charts_top = inner_top - 0.07                               # keep titles/ticks inside border
    charts_height = max(0.05, charts_top - charts_bottom)

    # Axes (leave inner margins so labels/ticks stay inside border)
    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    def setup_xaxis(ax):
        ax.set_xlim(*xlim)
        ax.xaxis.set_major_locator(MultipleLocator(5))
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='major', linewidth=0.5, alpha=0.4)

    # Colour+marker cycling (match your original script)
    colors  = ['b','g','r','c','m','y','k']
    markers = ['o','x','s','^']
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # LEFT: Depth vs Ip
    for bh, xs, ds, elevs, Z in series_ip:
        c, m = next(color_marker_combos)
        left_ax.scatter(xs, ds, color=c, marker=m, s=25, label=f"{bh}, {Z:.1f} m")
    left_ax.set_xlabel("Ip (%)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim)
    left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax)
    for s in left_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # Reset cycle so right plot matches colours/markers
    color_marker_combos = itertools.cycle([(c,m) for c in colors for m in markers])

    # RIGHT: Elevation vs Ip
    for bh, xs, ds, elevs, Z in series_ip:
        c, m = next(color_marker_combos)
        right_ax.scatter(xs, elevs, color=c, marker=m, s=25)
    right_ax.set_xlabel("Ip (%)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax)
    for s in right_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # ===== Legend (no title) =====
    handles, labels = left_ax.get_legend_handles_labels()
    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)

    if handles:
        fig.legend(
            handles, labels,
            loc='upper left',
            bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
            bbox_transform=fig.transFigure,
            ncol=4, frameon=True, fontsize=8,
            columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6
        )

    # Save
    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png:
        fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

    
    def export_sensitivity_pdf(
    series_sens,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    depth_ylim=(0, 35),
    margin_cm=1.0,
    x_label="Sensitivitet (S = cu/cur)"
):
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C2")

    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    tb_left, tb_bottom, tb_width, tb_height = draw_page_frame_and_title_block(
        fig, inner_left, inner_bottom, inner_w, inner_h,
        rapport_nr, figur_nr, tegn, kontr, godkj, dato, logo_path
    )

    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h
    charts_top = inner_top - 0.07
    charts_height = max(0.05, charts_top - charts_bottom)

    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    def setup_xaxis(ax):
        ticks = [1, 2, 5, 10, 20, 50, 100, 200, 500]
        ax.set_xscale('log')
        ax.set_xlim(ticks[0], ticks[-1])
        ax.set_xticks(ticks)
        ax.set_xticklabels([str(t) for t in ticks])
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='both', linewidth=0.5, alpha=0.4)

    for bh, sens, dep, elev, Z in series_sens:
        left_ax.scatter(sens, dep, color=bh_color.get(bh, 'k'), marker='D', s=25, label=f"{bh}, {Z:.1f} m")
    left_ax.set_xlabel(x_label)
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim); left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax); add_box_spines(left_ax)

    for bh, sens, dep, elev, Z in series_sens:
        right_ax.scatter(sens, elev, color=bh_color.get(bh, 'k'), marker='D', s=25)
    right_ax.set_xlabel(x_label)
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax); add_box_spines(right_ax)

    # Legend (no title), 4 cols
    handles, labels = [], []
    seen = set()
    for bh, *_ , Z in series_sens:
        lab = f"{bh}, {Z:.1f} m"
        if lab in seen: continue
        seen.add(lab)
        handles.append(plt.Line2D([], [], linestyle='', marker='s', markersize=8, color=bh_color.get(bh, 'k')))
        labels.append(lab)

    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)
    if handles:
        fig.legend(handles, labels, loc='upper left',
                   bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
                   bbox_transform=fig.transFigure, ncol=4, frameon=True, fontsize=8,
                   columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6)

    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png: fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

def export_enaks_deformation_pdf(
    series_def,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    depth_ylim=(0, 35),
    margin_cm=1.0,
    xlim=None  # e.g., (0, 20) if you want fixed range
):
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C5")

    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    tb_left, tb_bottom, tb_width, tb_height = draw_page_frame_and_title_block(
        fig, inner_left, inner_bottom, inner_w, inner_h,
        rapport_nr, figur_nr, tegn, kontr, godkj, dato, logo_path
    )

    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h
    charts_top = inner_top - 0.07
    charts_height = max(0.05, charts_top - charts_bottom)

    left_ax  = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    def setup_xaxis(ax):
        if xlim is not None:
            ax.set_xlim(*xlim)
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='major', linewidth=0.5, alpha=0.4)

    # LEFT: ε_f (%) vs Depth
    for bh, xs, ds, elevs, Z in series_def:
        left_ax.scatter(xs, ds, color=bh_color.get(bh, 'k'), marker='o', s=25, label=f"{bh}, {Z:.1f} m")
    left_ax.set_xlabel(r"Deformasjon ved brudd $\epsilon_f$ (%)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim); left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax); add_box_spines(left_ax)

    # RIGHT: ε_f (%) vs Elevation
    for bh, xs, ds, elevs, Z in series_def:
        right_ax.scatter(xs, elevs, color=bh_color.get(bh, 'k'), marker='o', s=25)
    right_ax.set_xlabel(r"Deformasjon ved brudd $\epsilon_f$ (%)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax); add_box_spines(right_ax)

    # Legend (no title), 4 cols
    handles, labels = [], []
    seen = set()
    for bh, *_ , Z in series_def:
        lab = f"{bh}, {Z:.1f} m"
        if lab in seen: continue
        seen.add(lab)
        handles.append(plt.Line2D([], [], linestyle='', marker='s', markersize=8, color=bh_color.get(bh, 'k')))
        labels.append(lab)

    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)
    if handles:
        fig.legend(handles, labels, loc='upper left',
                   bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
                   bbox_transform=fig.transFigure, ncol=4, frameon=True, fontsize=8,
                   columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6)

    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png: fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))
    
def export_curfc_pdf(
    data_series,
    outfile_pdf,
    outfile_png=None,
    logo_path=None,
    title_info=None,
    depth_ylim=(0, 35),
    margin_cm=1.0
):
   
    # Defaults
    if title_info is None:
        title_info = {}
    rapport_nr = title_info.get("rapport_nr", "XX")
    dato       = title_info.get("dato", "2025-09-22")
    tegn       = title_info.get("tegn", "IGH")
    kontr      = title_info.get("kontr", "JOG")
    godkj      = title_info.get("godkj", "AGR")
    figur_nr   = title_info.get("figur_nr", "C3")

    # Figure: A4 landscape (inches)
    fig_w, fig_h = 11.69, 8.27
    fig = plt.figure(figsize=(fig_w, fig_h))

    # 1 cm margin -> figure fractions
    margin_in = margin_cm / 2.54
    inner_left   = margin_in / fig_w
    inner_right  = 1.0 - margin_in / fig_w
    inner_bottom = margin_in / fig_h
    inner_top    = 1.0 - margin_in / fig_h
    inner_w      = inner_right - inner_left
    inner_h      = inner_top - inner_bottom

    tb_left, tb_bottom, tb_width, tb_height = draw_page_frame_and_title_block(
        fig, inner_left, inner_bottom, inner_w, inner_h,
        rapport_nr, figur_nr, tegn, kontr, godkj, dato, logo_path
    )
   
    # Bottom fixed 0.3 cm above title block top
    charts_bottom = (tb_bottom + tb_height) + (0.3/2.54)/fig_h
    # Lower the top so titles/ticks sit inside page border
    charts_top = inner_top - 0.07
    charts_height = max(0.05, charts_top - charts_bottom)

    # Axes: LEFT depth, RIGHT elevation. Leave a bit of inner margins so labels/ticks stay inside border.
    left_ax = fig.add_axes([inner_left + inner_w*0.06, charts_bottom, inner_w*0.40, charts_height])
    right_ax = fig.add_axes([inner_left + inner_w*0.54, charts_bottom, inner_w*0.38, charts_height])

    # X-axis formatting (as in your good version)
    def setup_xaxis(ax):
        ax.set_xscale('log')
        ax.set_xticks([0.1,0.2,0.5,1,2,5,10,20,30])
        ax.set_xticklabels(['0.1','0.2','0.5','1','2','5','10','20','30'])
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
        ax.grid(True, which='both', linewidth=0.5, alpha=0.4)

    # Colors/markers cycle for series
    colors = ['b','g','r','c','m','y','k']
    markers = ['o','x','s','^']
    cmb = itertools.cycle([(c, m) for c in colors for m in markers])
    style_map = {name: next(cmb) for name, *_ in data_series}

    # LEFT: Depth vs strength
    for workbook_name, x_vals, y_vals, elev_vals, terrain_level in data_series:
        c, m = style_map[workbook_name]
        left_ax.scatter(x_vals, y_vals, label=f"{workbook_name}, {terrain_level:.1f} m",
                        color=c, marker=m, s=25)
    left_ax.set_xlabel("Omrørt skjærstyrke (kPa)")
    left_ax.set_ylabel("Dybde (m)")
    left_ax.set_ylim(*depth_ylim)
    left_ax.invert_yaxis()
    left_ax.yaxis.tick_left(); left_ax.yaxis.set_label_position("left")
    setup_xaxis(left_ax)
    for s in left_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")

    # RIGHT: Elevation vs strength
    for workbook_name, x_vals, y_vals, elev_vals, terrain_level in data_series:
        c, m = style_map[workbook_name]
        right_ax.scatter(x_vals, elev_vals, color=c, marker=m, s=25)
    right_ax.set_xlabel("Omrørt skjærstyrke (kPa)")
    right_ax.set_ylabel("kote (m)")
    right_ax.yaxis.tick_right(); right_ax.yaxis.set_label_position("right")
    setup_xaxis(right_ax)
    for s in right_ax.spines.values():
        s.set_visible(True); s.set_linewidth(1.0); s.set_edgecolor("black")
        
        # --- Vertikal referanselinje (x = 1.27) ---
    left_ax.axvline(1.27, color='k', linestyle='--', linewidth=1.0, zorder=3)

    # ===== Legend: 4 columns, vertically centered to title block =====
    legend_w = (tb_left - (inner_left + inner_w * 0.02)) - inner_w * 0.02
    legend_h = tb_height * 0.60
    legend_x0 = inner_left + inner_w * 0.02
    legend_y0 = (tb_bottom + tb_height/2) - (legend_h/2)

    handles, labels = left_ax.get_legend_handles_labels()
    if handles:
        fig.legend(
            handles, labels,
            loc='upper left',
            bbox_to_anchor=(legend_x0, legend_y0, legend_w, legend_h),
            bbox_transform=fig.transFigure,
            ncol=4, frameon=True, fontsize=8,
            columnspacing=0.8, handletextpad=0.6, borderaxespad=0.6
        )

    # Save
    fig.savefig(outfile_pdf, format="pdf")
    if outfile_png:
        fig.savefig(outfile_png, dpi=300)
    plt.close(fig)
    print(f"Saved: {outfile_pdf}" + (f"\nPreview: {outfile_png}" if outfile_png else ""))

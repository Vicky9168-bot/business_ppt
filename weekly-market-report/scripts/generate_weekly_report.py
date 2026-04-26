#!/usr/bin/env python3
"""
早晨財經速解讀 2026-04-24 市場分析報告 — PPT 生成腳本
風格：Corporate Blue & White Business Theme
主題：費半破萬點 硬體狂飆 軟體失速 | 台股多殺多 股市大換手
內容：費半連17紅站上萬點、硬體vs軟體分化、台股換股行情、台灣外銷訂單創新高、金管會鬆綁
分析日期：2026-04-24
"""

import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from io import BytesIO
from math import pi

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ============================================================
# Config
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '..'))
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'output', 'reports')
TEMP_DIR = os.path.join(PROJECT_ROOT, 'output', 'temp')
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)
OUTPUT_PPTX = os.path.join(OUTPUT_DIR, '20260424_report.pptx')

# Colors — Corporate Blue Theme
CORP_BLUE = RGBColor(0x2B, 0x57, 0x9A)
CORP_BLUE_DARK = RGBColor(0x1A, 0x3A, 0x6C)
CORP_BLUE_LIGHT = RGBColor(0x4A, 0x86, 0xC8)
ACCENT_BLUE = RGBColor(0x5B, 0xA0, 0xD6)
SKY_BLUE = RGBColor(0xD6, 0xE8, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF0, 0xF2, 0xF5)
MID_GRAY = RGBColor(0x99, 0x99, 0x99)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
CHARCOAL = RGBColor(0x2D, 0x2D, 0x2D)
ACCENT_RED = RGBColor(0xE8, 0x4D, 0x4D)
ACCENT_GREEN = RGBColor(0x27, 0xAE, 0x60)
ACCENT_ORANGE = RGBColor(0xF5, 0x9E, 0x0B)
ACCENT_PURPLE = RGBColor(0x8E, 0x44, 0xAD)
BULL_GREEN = RGBColor(0x26, 0xA6, 0x9A)
BEAR_RED = RGBColor(0xEF, 0x53, 0x50)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FONT_CN = 'Microsoft JhengHei'
FONT_EN = 'Arial'
TOTAL_PAGES = 20

# ============================================================
# Matplotlib Setup
# ============================================================

def setup_matplotlib():
    plt.rcParams['font.family'] = ['Heiti TC', 'Arial Unicode MS', 'Hiragino Sans', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['figure.dpi'] = 150
    plt.rcParams['savefig.dpi'] = 250
    plt.rcParams['font.size'] = 14
    plt.rcParams['axes.labelsize'] = 15
    plt.rcParams['axes.titlesize'] = 18


def fig_to_image_stream(fig, dpi=250):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight',
                facecolor='white', edgecolor='none', pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf


# ============================================================
# Helper functions
# ============================================================

def _add_bg_rect(slide, left, top, width, height, color, alpha=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    if alpha is not None:
        _set_shape_alpha(shape, alpha)
    return shape


def _set_shape_alpha(shape, alpha_pct):
    fill_elem = shape.fill._fill
    solid = fill_elem.find(qn('a:solidFill'))
    if solid is not None:
        srgb = solid.find(qn('a:srgbClr'))
        if srgb is not None:
            alpha_elem = srgb.find(qn('a:alpha'))
            if alpha_elem is None:
                from lxml import etree
                alpha_elem = etree.SubElement(srgb, qn('a:alpha'))
            alpha_elem.set('val', str(int((100 - alpha_pct) * 1000)))


def _add_text_box(slide, left, top, width, height, text, font_size=14,
                  font_color=DARK_GRAY, bold=False, alignment=PP_ALIGN.LEFT,
                  font_name=FONT_CN, line_spacing=1.2):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    p.space_after = Pt(0)
    sp = p._pPr
    if sp is None:
        from lxml import etree
        sp = etree.SubElement(p._p, qn('a:pPr'))
    lnSpc = sp.find(qn('a:lnSpc'))
    if lnSpc is None:
        from lxml import etree
        lnSpc = etree.SubElement(sp, qn('a:lnSpc'))
        spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', str(int(line_spacing * 100000)))
    return txBox


def _add_multiline_text_box(slide, left, top, width, height, lines,
                            font_size=14, font_color=DARK_GRAY,
                            font_name=FONT_CN, line_spacing=1.5,
                            alignment=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_data in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if isinstance(line_data, dict):
            p.text = line_data.get('text', '')
            p.font.size = Pt(line_data.get('size', font_size))
            p.font.color.rgb = line_data.get('color', font_color)
            p.font.bold = line_data.get('bold', False)
            p.font.name = line_data.get('font', font_name)
        else:
            p.text = str(line_data)
            p.font.size = Pt(font_size)
            p.font.color.rgb = font_color
            p.font.name = font_name
        p.alignment = alignment
        p.space_after = Pt(2)
    return txBox


def _add_circle(slide, left, top, size, color, text='', font_size=14,
                font_color=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = font_color
        p.font.bold = True
        p.font.name = FONT_EN
        p.alignment = PP_ALIGN.CENTER
    return shape


def _add_rounded_rect(slide, left, top, width, height, color, text='',
                      font_size=13, font_color=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(8)
        tf.margin_right = Pt(8)
        tf.margin_top = Pt(6)
        tf.margin_bottom = Pt(6)
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = font_color
        p.font.name = FONT_CN
        p.alignment = PP_ALIGN.CENTER
    return shape


def _add_header_bar(slide, title_text, subtitle_text=''):
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.1), CORP_BLUE)
    _add_circle(slide, Inches(0.3), Inches(0.15), Inches(0.8), CORP_BLUE_DARK)
    _add_text_box(slide, Inches(1.3), Inches(0.15), Inches(10), Inches(0.55),
                  title_text, font_size=28, font_color=WHITE, bold=True)
    if subtitle_text:
        _add_text_box(slide, Inches(1.3), Inches(0.65), Inches(10), Inches(0.4),
                      subtitle_text, font_size=15, font_color=SKY_BLUE)
    _add_text_box(slide, Inches(10.5), Inches(0.15), Inches(2.5), Inches(0.4),
                  '早晨財經速解讀', font_size=16, font_color=WHITE, bold=True,
                  alignment=PP_ALIGN.RIGHT)
    _add_text_box(slide, Inches(10.5), Inches(0.55), Inches(2.5), Inches(0.35),
                  '2026.04.24', font_size=12, font_color=SKY_BLUE,
                  alignment=PP_ALIGN.RIGHT)


def _add_footer(slide, page_num, total_pages=TOTAL_PAGES):
    _add_bg_rect(slide, Inches(0), Inches(7.1), SLIDE_W, Inches(0.4), LIGHT_GRAY)
    _add_text_box(slide, Inches(0.5), Inches(7.15), Inches(7), Inches(0.3),
                  '早晨財經速解讀 — 2026.04.24 市場分析報告', font_size=10,
                  font_color=MID_GRAY)
    _add_text_box(slide, Inches(10), Inches(7.15), Inches(3), Inches(0.3),
                  f'{page_num} / {total_pages}', font_size=10,
                  font_color=MID_GRAY, alignment=PP_ALIGN.RIGHT, font_name=FONT_EN)


def _add_section_divider(slide, num, title, subtitle, detail=''):
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, CORP_BLUE)
    _add_circle(slide, Inches(1.5), Inches(2), Inches(2.5), WHITE, num,
                font_size=48, font_color=CORP_BLUE)
    _add_text_box(slide, Inches(5), Inches(2.2), Inches(7), Inches(1),
                  title, font_size=40, font_color=WHITE, bold=True, font_name=FONT_EN)
    _add_text_box(slide, Inches(5), Inches(3.3), Inches(7), Inches(0.6),
                  subtitle, font_size=18, font_color=SKY_BLUE)
    _add_bg_rect(slide, Inches(5), Inches(4.1), Inches(4), Inches(0.04), WHITE)
    if detail:
        _add_text_box(slide, Inches(5), Inches(4.4), Inches(7), Inches(0.5),
                      detail, font_size=14, font_color=SKY_BLUE)


# ============================================================
# Chart Generation
# ============================================================

def gen_charts():
    charts = {}
    blue_hex = '#2B579A'
    blue_light_hex = '#4A86C8'
    accent_hex = '#5BA0D6'
    green_hex = '#27AE60'
    red_hex = '#E84D4D'
    orange_hex = '#F59E0B'
    purple_hex = '#8E44AD'

    # ── 1. US Indices 4/23 Performance ──
    fig, ax = plt.subplots(figsize=(11, 3.7))
    indices = ['道瓊工業\n49,310', '標普 500\n7,108', '納斯達克\n24,438', '費城半導體\n10,078']
    daily_chg = [-0.36, -0.41, -0.89, 1.71]
    colors = [green_hex if v > 0 else red_hex for v in daily_chg]
    bars = ax.barh(indices, daily_chg, color=colors, height=0.5, edgecolor='white', linewidth=1.5)
    ax.axvline(x=0, color='#ccc', linewidth=1.2, linestyle='-')
    for bar, val in zip(bars, daily_chg):
        x_pos = bar.get_width() + 0.05 if val > 0 else bar.get_width() - 0.05
        ha = 'left' if val > 0 else 'right'
        label = f'+{val:.2f}%' if val > 0 else f'{val:.2f}%'
        ax.text(x_pos, bar.get_y() + bar.get_height() / 2,
                label, va='center', ha=ha, fontsize=16,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xlim(-1.5, 2.5)
    ax.set_title('美股四大指數 — 費半連17紅突破10,000點！三大指數收黑（2026/04/23）',
                 fontsize=17, fontweight='bold', color=blue_hex, pad=15)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.tick_params(axis='both', labelsize=14)
    # Annotate SOX milestone
    ax.annotate('歷史新高！', xy=(1.71, 3), xytext=(1.5, 2.5),
                arrowprops=dict(arrowstyle='->', color=green_hex, lw=2),
                fontsize=13, color=green_hex, fontweight='bold')
    fig.tight_layout()
    charts['weekly_us'] = fig_to_image_stream(fig)

    # ── 2. 台股成交量與走勢 ──
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4))
    # 台股收盤走勢（近期）
    dates = ['4/17', '4/20', '4/21', '4/22', '4/23', '4/24']
    taiex = [37200, 37800, 38100, 38900, 37600, 38434]
    ax1.plot(dates, taiex, 'o-', color=blue_hex, linewidth=2.5, markersize=8)
    ax1.fill_between(range(len(dates)), taiex, min(taiex)*0.99,
                     alpha=0.15, color=blue_hex)
    ax1.set_xticks(range(len(dates)))
    ax1.set_xticklabels(dates, fontsize=11)
    ax1.set_ylabel('點數', fontsize=12, color=blue_hex)
    ax1.set_title('台股指數走勢（4/17-4/24）', fontsize=15, fontweight='bold',
                  color=blue_hex, pad=10)
    ax1.annotate('38,434\n（+720pt）', xy=(5, 38434), xytext=(3.5, 38700),
                 arrowprops=dict(arrowstyle='->', color=green_hex, lw=1.5),
                 fontsize=12, color=green_hex, fontweight='bold')
    ax1.annotate('多殺多', xy=(4, 37600), xytext=(2.5, 37300),
                 arrowprops=dict(arrowstyle='->', color=red_hex, lw=1.5),
                 fontsize=12, color=red_hex, fontweight='bold')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.set_ylim(36800, 39200)

    # 成交量柱狀圖
    volumes = [6500, 7200, 7800, 9500, 13000, 8200]
    vol_colors = [blue_light_hex]*4 + [red_hex] + [blue_hex]
    bars = ax2.bar(range(len(dates)), volumes, color=vol_colors, width=0.6,
                   edgecolor='white', linewidth=1.5)
    ax2.set_xticks(range(len(dates)))
    ax2.set_xticklabels(dates, fontsize=11)
    ax2.set_ylabel('成交量（億元）', fontsize=12, color=blue_hex)
    ax2.set_title('台股日成交量（億元）', fontsize=15, fontweight='bold',
                  color=blue_hex, pad=10)
    ax2.annotate('歷史新高\n1.3兆', xy=(4, 13000), xytext=(2.5, 11500),
                 arrowprops=dict(arrowstyle='->', color=red_hex, lw=1.5),
                 fontsize=12, color=red_hex, fontweight='bold')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    fig.tight_layout()
    charts['asia_weekly'] = fig_to_image_stream(fig)

    # ── 3. 硬體 vs 軟體分化 (SOX vs IGV) ──
    fig, ax = plt.subplots(figsize=(9, 3.8))
    categories = ['費半 SOX\n連17紅', 'IGV 軟體\nETF', '納斯達克\n指數', 'QQQ\n科技ETF']
    week_chg = [8.5, -7.2, 2.1, 2.8]
    bar_colors = [green_hex if v > 0 else red_hex for v in week_chg]
    bars = ax.bar(range(len(categories)), week_chg, color=bar_colors,
                  width=0.5, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, week_chg)):
        y_pos = val + 0.3 if val >= 0 else val - 0.8
        label = f'+{val:.1f}%' if val > 0 else f'{val:.1f}%'
        ax.text(i, y_pos, label, ha='center', fontsize=14,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(categories)))
    ax.set_xticklabels(categories, fontsize=12)
    ax.set_ylabel('近期漲跌（%）', fontsize=13, color=blue_hex)
    ax.set_title('硬體 vs 軟體極端分化：費半+8.5% vs IGV軟體-7.2%', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axhline(y=0, color='#ccc', linewidth=1)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.set_ylim(-10, 12)
    fig.tight_layout()
    charts['software_pe'] = fig_to_image_stream(fig)

    # ── 4. 資金輪動：AI硬體 vs AI軟體 vs 其他板塊 ──
    fig, ax = plt.subplots(figsize=(9, 4))
    sectors = ['AI硬體\n(SOX)', 'AI軟體\n(IGV)', '能源\n(XLE)', '消費\n(XLY)', '公用事業\n(XLU)']
    fund_flow = [35, -20, 5, -8, -12]
    flow_colors = [green_hex if v > 0 else red_hex for v in fund_flow]
    bars = ax.bar(range(len(sectors)), fund_flow, color=flow_colors,
                  width=0.5, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, fund_flow)):
        y_pos = val + 1.0 if val >= 0 else val - 2.5
        label = f'+{val}' if val > 0 else f'{val}'
        ax.text(i, y_pos, label, ha='center', fontsize=14,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(sectors)))
    ax.set_xticklabels(sectors, fontsize=12)
    ax.set_ylabel('相對資金流入（相對值）', fontsize=13, color=blue_hex)
    ax.set_title('資金輪動：AI硬體獨佔鰲頭，其他板塊全面遭抽離', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axhline(y=0, color='#ccc', linewidth=1)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    fig.tight_layout()
    charts['peg_trend'] = fig_to_image_stream(fig)

    # ── 5. 特斯拉財報 Q1 2026 解析 ──
    fig, ax = plt.subplots(figsize=(10, 3.5))
    tsla_cat = ['EPS 預期\n(Q1 2026)', 'EPS 實際\n(Q1 2026)', '資本支出\n(26年原預期)', '資本支出\n(26年實際)']
    tsla_vals = [0.6, 0.25, 200, 250]
    tsla_colors = [blue_hex, red_hex, blue_light_hex, orange_hex]
    # Use two y-axes
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4))
    # Left: EPS
    eps_cats = ['預期 EPS\n$0.60', '實際 EPS\n$0.25']
    eps_vals = [0.60, 0.25]
    eps_colors = [blue_hex, red_hex]
    bars1 = ax1.bar(range(len(eps_cats)), eps_vals, color=eps_colors, width=0.4,
                    edgecolor='white', linewidth=1.5)
    for bar, val in zip(bars1, eps_vals):
        ax1.text(bar.get_x() + bar.get_width()/2, val + 0.02, f'${val:.2f}',
                 ha='center', fontsize=14, fontweight='bold', color=bar.get_facecolor())
    ax1.set_xticks(range(len(eps_cats)))
    ax1.set_xticklabels(eps_cats, fontsize=12)
    ax1.set_ylabel('每股盈餘（美元）', fontsize=12, color=blue_hex)
    ax1.set_title('特斯拉 Q1 EPS\n連4季低於預期', fontsize=14, fontweight='bold',
                  color=blue_hex, pad=10)
    ax1.annotate('低於預期\n-58%', xy=(1, 0.25), xytext=(0.5, 0.5),
                 arrowprops=dict(arrowstyle='->', color=red_hex, lw=1.5),
                 fontsize=12, color=red_hex, fontweight='bold')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.set_ylim(0, 0.85)
    # Right: Capital expenditure
    capex_cats = ['26年資本支出\n原估$200億', '26年資本支出\n實際$250億']
    capex_vals = [200, 250]
    capex_colors = [blue_hex, orange_hex]
    bars2 = ax2.bar(range(len(capex_cats)), capex_vals, color=capex_colors, width=0.4,
                    edgecolor='white', linewidth=1.5)
    for bar, val in zip(bars2, capex_vals):
        ax2.text(bar.get_x() + bar.get_width()/2, val + 5, f'${val}億',
                 ha='center', fontsize=14, fontweight='bold', color=bar.get_facecolor())
    ax2.set_xticks(range(len(capex_cats)))
    ax2.set_xticklabels(capex_cats, fontsize=12)
    ax2.set_ylabel('資本支出（億美元）', fontsize=12, color=blue_hex)
    ax2.set_title('特斯拉 2026 CAPEX\n增加$50億至$250億', fontsize=14, fontweight='bold',
                  color=blue_hex, pad=10)
    ax2.annotate('+$50億\n+25%', xy=(1, 250), xytext=(0.5, 240),
                 arrowprops=dict(arrowstyle='->', color=orange_hex, lw=1.5),
                 fontsize=12, color=orange_hex, fontweight='bold')
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.set_ylim(0, 310)
    fig.tight_layout()
    charts['taiwan_export'] = fig_to_image_stream(fig)

    # ── 6. 台灣外銷訂單歷史新高 ──
    fig, ax = plt.subplots(figsize=(10, 4.2))
    months = ['2025\n10月', '2025\n11月', '2025\n12月', '2026\n1月', '2026\n2月', '2026\n3月']
    orders = [720, 750, 810, 780, 820, 911]
    order_colors = [blue_light_hex]*5 + ['#1a7a4a']
    bars = ax.bar(range(len(months)), orders, color=order_colors,
                  width=0.55, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, orders)):
        y_pos = val + 10
        label = f'{val}億' if i < 5 else f'{val}億\n歷史新高!'
        color = bar.get_facecolor()
        ax.text(i, y_pos, label, ha='center', fontsize=12 if i < 5 else 13,
                fontweight='bold', color=color)
    ax.set_xticks(range(len(months)))
    ax.set_xticklabels(months, fontsize=12)
    ax.set_ylabel('外銷訂單（億美元）', fontsize=13, color=blue_hex)
    ax.set_title('台灣外銷訂單：3月份首破900億美元大關！年增幅 +65.9%', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axhline(y=900, color='#1a7a4a', linewidth=2, linestyle='--', alpha=0.7)
    ax.text(5.4, 905, '900億里程碑', fontsize=11, color='#1a7a4a', fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.invert_yaxis()
    ax.set_ylim(0, 1050)
    ax.invert_yaxis()
    fig.tight_layout()
    charts['capex_war'] = fig_to_image_stream(fig)

    # ── 7. 產業動能雷達圖 ──
    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    categories_r = ['AI 硬體\n(費半)', '台灣\n外銷訂單', '金管會\n政策利多', '能源股\n築底', '軟體股\n動能']
    N = len(categories_r)
    current = [9, 9, 8, 5, 3]
    future = [9, 9, 9, 7, 5]
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    current += current[:1]
    future += future[:1]
    ax.plot(angles, current, 'o-', color=blue_hex, linewidth=2.5,
            label='當前評分', markersize=8)
    ax.fill(angles, current, alpha=0.15, color=blue_hex)
    ax.plot(angles, future, 's--', color=orange_hex, linewidth=2.5,
            label='3-6月展望', markersize=8)
    ax.fill(angles, future, alpha=0.1, color=orange_hex)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories_r, fontsize=12, fontweight='bold')
    ax.set_ylim(0, 10)
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(['2', '4', '6', '8', '10'], fontsize=10, color='gray')
    ax.legend(fontsize=12, loc='upper right', bbox_to_anchor=(1.3, 1.1))
    ax.set_title('市場動能雷達圖', fontsize=18,
                 fontweight='bold', color=blue_hex, pad=30)
    fig.tight_layout()
    charts['sector_radar'] = fig_to_image_stream(fig)

    # ── 8. 金管會政策鬆綁時間軸 ──
    fig, ax = plt.subplots(figsize=(9, 3.5))
    policies = ['主動ETF持股\n10%→25%', '外債擔保\n換台幣', '台股配息\n美元化', '雙重分配\n制度建立', '外資投信\n進駐台灣']
    impact = [9, 7, 8, 6, 8]
    p_colors = [blue_hex, blue_light_hex, accent_hex, orange_hex, '#1a7a4a']
    bars = ax.bar(range(len(policies)), impact, color=p_colors,
                  width=0.55, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, impact)):
        ax.text(i, val + 0.2, f'{val}/10', ha='center', fontsize=13,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(policies)))
    ax.set_xticklabels(policies, fontsize=11)
    ax.set_ylabel('政策影響力（1-10）', fontsize=13, color=blue_hex)
    ax.set_title('金管會政策鬆綁矩陣：制度×場域×產品三線並進', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.set_ylim(0, 11.5)
    fig.tight_layout()
    charts['capex_ratio'] = fig_to_image_stream(fig)

    # ── 9. 資產配置建議 ──
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(8.5, 3.5))
    labels_c = ['台積電/AI硬體', '主動型ETF', '現金', '能源/防禦', '其他']
    sizes_c = [35, 25, 20, 10, 10]
    colors_c = [blue_hex, blue_light_hex, '#95a5a6', orange_hex, accent_hex]
    wedges, texts, autotexts = ax1.pie(
        sizes_c, labels=labels_c, autopct='%1.0f%%', colors=colors_c,
        startangle=90, pctdistance=0.75, textprops={'fontsize': 11})
    for at in autotexts:
        at.set_fontweight('bold')
        at.set_color('white')
    ax1.set_title('積極型（AI硬體主線）', fontsize=15, fontweight='bold', color=blue_hex, pad=15)

    labels_g = ['現金', '台積電', '防禦科技', '債券/黃金', '其他']
    sizes_g = [35, 25, 20, 10, 10]
    colors_g = ['#95a5a6', blue_hex, blue_light_hex, orange_hex, accent_hex]
    wedges2, texts2, autotexts2 = ax2.pie(
        sizes_g, labels=labels_g, autopct='%1.0f%%', colors=colors_g,
        startangle=90, pctdistance=0.75, textprops={'fontsize': 11})
    for at in autotexts2:
        at.set_fontweight('bold')
        at.set_color('white')
    ax2.set_title('穩健型（等換股完成）', fontsize=15, fontweight='bold', color=blue_hex, pad=15)
    fig.tight_layout()
    charts['allocation'] = fig_to_image_stream(fig)

    # ── 10. 外資賣超 vs 台股走勢 ──
    fig, ax = plt.subplots(figsize=(11, 4.0))
    years = ['2020', '2021', '2022', '2023', '2024', '2025', '2026\nYTD']
    foreign_sell = [-5300, -4500, -12000, 0, -6900, -5900, -4371]
    bar_colors = [red_hex if v < 0 else green_hex for v in foreign_sell]
    bars = ax.bar(range(len(years)), foreign_sell, color=bar_colors,
                  width=0.55, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, foreign_sell)):
        y_pos = val - 350 if val < 0 else val + 100
        label = f'{val:,}億' if val != 0 else '小買'
        ax.text(i, y_pos, label, ha='center', fontsize=11,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(years)))
    ax.set_xticklabels(years, fontsize=12)
    ax.set_ylabel('外資買賣超（億元）', fontsize=13, color=blue_hex)
    ax.set_title('外資連年賣超台股：累計已超過3兆，回補是最大潛在觸媒', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axhline(y=0, color='#ccc', linewidth=1.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    fig.tight_layout()
    charts['asset_ytd'] = fig_to_image_stream(fig)

    # ── 11. 主動型ETF表現 ──
    fig, ax = plt.subplots(figsize=(10, 3.5))
    etf_names = ['992A', '981A', '994A', '主動型\nETF均值', '0050\n台灣50', '大盤\n加權指數']
    etf_ytd = [62, 60, 61, 43, 28, 25]
    etf_colors = ['#1a7a4a', '#1a7a4a', '#1a7a4a', green_hex, blue_light_hex, '#aaa']
    bars = ax.bar(range(len(etf_names)), etf_ytd, color=etf_colors,
                  width=0.55, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, etf_ytd)):
        y_pos = val + 1.5
        ax.text(i, y_pos, f'+{val}%', ha='center', fontsize=13,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(etf_names)))
    ax.set_xticklabels(etf_names, fontsize=12)
    ax.set_ylabel('YTD 漲幅（%）', fontsize=13, color=blue_hex)
    ax.set_title('主動型ETF全面跑贏大盤：3檔突破60%，均值+43%（2026 YTD）', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.set_ylim(0, 75)
    # Bracket for active ETF
    ax.axhline(y=43, color=blue_hex, linewidth=1.5, linestyle='--', alpha=0.5)
    ax.text(5.4, 44, '均值43%', fontsize=10, color=blue_hex)
    fig.tight_layout()
    charts['global_stocks_ytd'] = fig_to_image_stream(fig)

    # ── 12. 油價 & PMI 雙指標 ──
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.0))
    # Oil price trend
    oil_months = ['2月', '3月', '3月底', '4月', '4/23']
    oil_prices = [88, 92, 85, 95, 100]
    oil_colors = [blue_light_hex, blue_light_hex, red_hex, orange_hex, orange_hex]
    ax1.plot(range(len(oil_months)), oil_prices, 'o-', color=orange_hex,
             linewidth=2.5, markersize=9)
    ax1.fill_between(range(len(oil_months)), oil_prices, 80, alpha=0.12, color=orange_hex)
    ax1.set_xticks(range(len(oil_months)))
    ax1.set_xticklabels(oil_months, fontsize=11)
    ax1.set_ylabel('每桶美元', fontsize=12, color=orange_hex)
    ax1.set_title('油價重回$100大關\n中東地緣政治升溫', fontsize=14, fontweight='bold',
                  color=orange_hex, pad=10)
    ax1.axhline(y=100, color=orange_hex, linewidth=2, linestyle='--', alpha=0.7)
    ax1.text(4.1, 100.5, '$100關卡', fontsize=10, color=orange_hex, fontweight='bold')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.set_ylim(78, 110)
    # PMI bar
    pmi_months = ['1月', '2月', '3月', '4月(預)']
    pmi_vals = [49.5, 50.2, 50.8, 51.2]
    pmi_colors = [red_hex, blue_light_hex, blue_hex, blue_hex]
    bars = ax2.bar(range(len(pmi_months)), pmi_vals, color=pmi_colors,
                   width=0.5, edgecolor='white', linewidth=1.5)
    for bar, val in zip(bars, pmi_vals):
        ax2.text(bar.get_x() + bar.get_width()/2, val + 0.1, f'{val}',
                 ha='center', fontsize=13, fontweight='bold',
                 color=bar.get_facecolor())
    ax2.axhline(y=50, color='#ccc', linewidth=2, linestyle='--')
    ax2.text(3.5, 50.05, '榮枯線50', fontsize=10, color='gray')
    ax2.set_xticks(range(len(pmi_months)))
    ax2.set_xticklabels(pmi_months, fontsize=11)
    ax2.set_ylabel('PMI 指數', fontsize=12, color=blue_hex)
    ax2.set_title('美國PMI緩步回升\n企業備貨性需求推動', fontsize=14, fontweight='bold',
                  color=blue_hex, pad=10)
    ax2.spines['top'].set_visible(False)
    ax2.spines['right'].set_visible(False)
    ax2.set_ylim(48, 53)
    fig.tight_layout()
    charts['commodity_ytd'] = fig_to_image_stream(fig)

    # ── 13. 台灣薪資結構 ──
    fig, ax = plt.subplots(figsize=(10, 3.8))
    industries = ['金融業', '資通訊', '電力能源', '製造業', '整體均值', '教育服務', '住宿餐飲']
    avg_salary = [7.4, 6.8, 6.5, 5.8, 9.0, 3.5, 3.1]
    ind_colors = ['#1a7a4a', blue_hex, blue_light_hex, accent_hex, orange_hex, '#aaa', red_hex]
    bars = ax.barh(range(len(industries)), avg_salary, color=ind_colors,
                   height=0.55, edgecolor='white', linewidth=1.5)
    for bar, val in zip(bars, avg_salary):
        x_pos = bar.get_width() + 0.1
        label = f'{val}萬' if val != 9.0 else f'{val}萬（含年終）'
        ax.text(x_pos, bar.get_y() + bar.get_height() / 2,
                label, va='center', ha='left', fontsize=12,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_yticks(range(len(industries)))
    ax.set_yticklabels(industries, fontsize=12)
    ax.set_xlabel('月薪（萬元）', fontsize=13, color=blue_hex)
    ax.set_title('台灣薪資極化：金融/科技 vs 服務業差距擴大（2026.3月）', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axvline(x=4.0, color=orange_hex, linewidth=1.5, linestyle='--', alpha=0.7)
    ax.text(4.05, 6.7, '中位數\n3.9-4萬', fontsize=10, color=orange_hex)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.set_xlim(0, 12)
    ax.invert_yaxis()
    fig.tight_layout()
    charts['sector_rotation'] = fig_to_image_stream(fig)

    # ── 14. 台股換股邏輯 ──
    fig, ax = plt.subplots(figsize=(10, 3.5))
    stock_cat = ['台積電\n(未崩)', '大型權值\n(穩健)', '中型股\n(換手)', '小型股\n(多殺多)', '中石化等\n(獲利了結)']
    stock_change = [2.5, 1.8, -3.5, -6.2, -8.5]
    s_colors = [green_hex, blue_light_hex, red_hex, red_hex, '#8B0000']
    bars = ax.bar(range(len(stock_cat)), stock_change, color=s_colors,
                  width=0.5, edgecolor='white', linewidth=1.5)
    for i, (bar, val) in enumerate(zip(bars, stock_change)):
        y_pos = val + 0.3 if val >= 0 else val - 0.7
        label = f'+{val:.1f}%' if val > 0 else f'{val:.1f}%'
        ax.text(i, y_pos, label, ha='center', fontsize=14,
                fontweight='bold', color=bar.get_facecolor())
    ax.set_xticks(range(len(stock_cat)))
    ax.set_xticklabels(stock_cat, fontsize=12)
    ax.set_ylabel('單日漲跌幅（%）', fontsize=13, color=blue_hex)
    ax.set_title('台股換股行情：資金從中小型股撤出，回流大型權值股', fontsize=17,
                 fontweight='bold', color=blue_hex, pad=15)
    ax.axhline(y=0, color='#ccc', linewidth=1)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#ddd')
    ax.spines['left'].set_color('#ddd')
    ax.set_ylim(-11, 5)
    fig.tight_layout()
    charts['sp500_earnings'] = fig_to_image_stream(fig)

    return charts


# ============================================================
# Slide Builders
# ============================================================

def slide_01_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, CORP_BLUE)
    _add_circle(slide, Inches(10.5), Inches(-0.5), Inches(3), CORP_BLUE_DARK)
    _add_circle(slide, Inches(11.5), Inches(5), Inches(2.5), CORP_BLUE_DARK)
    _add_circle(slide, Inches(-0.8), Inches(5.5), Inches(2), CORP_BLUE_DARK)
    _add_rounded_rect(slide, Inches(1), Inches(1.2), Inches(11.3), Inches(5), WHITE)
    _add_text_box(slide, Inches(1.5), Inches(1.5), Inches(10), Inches(0.8),
                  '早晨財經速解讀 — 市場分析報告', font_size=38, font_color=CORP_BLUE,
                  bold=True, alignment=PP_ALIGN.CENTER)
    _add_text_box(slide, Inches(1.5), Inches(2.5), Inches(10), Inches(0.9),
                  '費半破萬點 硬體狂飆 軟體失速｜台股多殺多 股市大換手',
                  font_size=20, font_color=CORP_BLUE_DARK,
                  bold=True, alignment=PP_ALIGN.CENTER)
    _add_bg_rect(slide, Inches(4.5), Inches(3.5), Inches(4.3), Inches(0.05), CORP_BLUE)
    _add_text_box(slide, Inches(1.5), Inches(3.8), Inches(10), Inches(0.5),
                  '費半連17紅站上10,000點 | SOX vs IGV 極端分化 | 台股成交量歷史新高',
                  font_size=14, font_color=MID_GRAY, alignment=PP_ALIGN.CENTER)
    _add_text_box(slide, Inches(1.5), Inches(4.4), Inches(10), Inches(0.5),
                  '台灣外銷訂單首破900億美元 | 金管會持股上限10%→25% | 主動ETF均值+43%',
                  font_size=14, font_color=MID_GRAY, alignment=PP_ALIGN.CENTER)
    _add_text_box(slide, Inches(1.5), Inches(5.1), Inches(10), Inches(0.4),
                  '2026 年 4 月 24 日', font_size=12,
                  font_color=MID_GRAY, alignment=PP_ALIGN.CENTER)
    _add_text_box(slide, Inches(3.5), Inches(6.5), Inches(6.3), Inches(0.5),
                  '早晨財經速解讀 — 2026.04.24 綜合分析', font_size=18,
                  font_color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)


def slide_02_contents(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '目錄 CONTENTS', '2026.04.24 費半破萬 + 台股換手 + 外銷創高 + 金管會鬆綁')
    _add_footer(slide, 2)

    sections = [
        {'num': '01', 'title': '市場總覽', 'desc': '費半破萬點連17紅\n台股多殺多換股\n成交量1.3兆歷史新高', 'color': CORP_BLUE},
        {'num': '02', 'title': '硬軟分化', 'desc': '硬體狂飆vs軟體失速\n特斯拉財報解析\n比特幣資金輪動', 'color': CORP_BLUE_LIGHT},
        {'num': '03', 'title': '台灣市場', 'desc': '外銷訂單創新高\n金管會三線鬆綁\n主動ETF表現', 'color': ACCENT_BLUE},
        {'num': '04', 'title': '投資策略', 'desc': '硬體主線配置\n換股邏輯分析\n觀察重點展望', 'color': ACCENT_PURPLE},
    ]

    start_x = Inches(0.8)
    for i, sec in enumerate(sections):
        x = start_x + i * Inches(3.2)
        y = Inches(2.0)
        _add_rounded_rect(slide, x, y, Inches(2.9), Inches(4.3), WHITE)
        _add_circle(slide, x + Inches(0.85), y + Inches(0.3), Inches(1.1),
                    sec['color'], sec['num'], font_size=28)
        _add_text_box(slide, x + Inches(0.1), y + Inches(1.6), Inches(2.7),
                      Inches(0.5), sec['title'], font_size=20,
                      font_color=CHARCOAL, bold=True, alignment=PP_ALIGN.CENTER)
        _add_text_box(slide, x + Inches(0.2), y + Inches(2.2), Inches(2.5),
                      Inches(1.5), sec['desc'], font_size=13,
                      font_color=MID_GRAY, alignment=PP_ALIGN.CENTER, line_spacing=1.5)
        _add_bg_rect(slide, x + Inches(0.8), y + Inches(4.0), Inches(1.3),
                     Inches(0.05), sec['color'])


def slide_03_part_one(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_section_divider(slide, '01', 'PART ONE.', '市場總覽 — 費半破萬 台股換手',
                         '費半連17紅+1.71% | 三大指數收黑 | 台股成交1.3兆歷史新高')


def slide_04_us_weekly(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '美股四大指數 — 費半破萬點！連17紅', '2026/04/23 — 道瓊-0.36% 標普-0.41% 那指-0.89% 費半+1.71%')
    _add_footer(slide, 4)

    idx_data = [
        ('道瓊工業', '49,310', '-179 pt', '-0.36%', ACCENT_RED),
        ('標普 500', '7,108', '-29 pt', '-0.41%', ACCENT_RED),
        ('納斯達克', '24,438', '-219 pt', '-0.89%', ACCENT_RED),
        ('費城半導體', '10,078', '+169 pt', '+1.71%', ACCENT_GREEN),
    ]
    for i, (name, price, change, pct, color) in enumerate(idx_data):
        x = Inches(0.4) + i * Inches(3.2)
        _add_rounded_rect(slide, x, Inches(1.4), Inches(3.0), Inches(1.6), WHITE)
        _add_bg_rect(slide, x + Inches(0.05), Inches(1.4), Inches(2.9), Inches(0.06), color)
        _add_text_box(slide, x + Inches(0.2), Inches(1.55), Inches(2.6), Inches(0.35),
                      name, font_size=15, font_color=CHARCOAL, bold=True, alignment=PP_ALIGN.CENTER)
        _add_text_box(slide, x + Inches(0.2), Inches(1.9), Inches(2.6), Inches(0.4),
                      price, font_size=24, font_color=CHARCOAL, bold=True,
                      alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        _add_text_box(slide, x + Inches(0.2), Inches(2.35), Inches(2.6), Inches(0.35),
                      f'{change}（{pct}）', font_size=14, font_color=color,
                      bold=True, alignment=PP_ALIGN.CENTER)

    img = charts['weekly_us']
    slide.shapes.add_picture(img, Inches(1.2), Inches(3.2), Inches(11), Inches(3.7))


def slide_05_taiwan_market(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '台股多殺多換手 — 成交量創歷史新高', '台北收38,434點(+720pt) | 成交1.3兆 | 外資僅賣27億 | 換股非換手')
    _add_footer(slide, 5)

    img = charts['asia_weekly']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.8), Inches(3.5))

    _add_rounded_rect(slide, Inches(8.3), Inches(1.3), Inches(4.7), Inches(3.5), WHITE)
    taiwan_lines = [
        {'text': '台股換股行情解析', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  4/24 台北股市 +720pt = 38,434', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  單日成交歷史新高：1.3兆', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  外資賣超僅27億（幅度小）', 'size': 13, 'color': DARK_GRAY},
        {'text': '  主要是「內資換股」非外資撤退', 'size': 13, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  台積電仍強勢，未崩盤', 'size': 13, 'color': ACCENT_GREEN},
        {'text': '  中小型股：中石化等多殺多', 'size': 13, 'color': ACCENT_RED},
    ]
    _add_multiline_text_box(slide, Inches(8.5), Inches(1.45), Inches(4.3), Inches(3.2),
                            taiwan_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    bottom = [
        {'text': '換股的核心邏輯', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  台股今年成為全球第六大市場，1,000點以內、8,000億以內均屬正常波動。上引線長 = 籌碼換手，非行情結束', 'size': 14, 'color': DARK_GRAY},
        {'text': '  資金從過熱中小型股撤出 → 回流台積電等基本面確定的大型股 = 行情換股升級', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3), bottom)


def slide_06_part_two(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_section_divider(slide, '02', 'PART TWO.', '硬體 vs 軟體 — 極端分化',
                         'SOX連17紅 | IGV軟體大跌 | 特斯拉財報 | 比特幣資金輪動')


def slide_07_hardware_software(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '硬體 vs 軟體極端分化', 'AI基礎建設需求爆發 | 軟體商業模式仍在調整 | 籌碼高度集中')
    _add_footer(slide, 7)

    img = charts['software_pe']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.8))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(4.9), Inches(5.5), WHITE)
    hw_sw_lines = [
        {'text': '分化背後的邏輯', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  費半連17紅，站上1萬點', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  AI基建訂單4-5年結構性循環', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  IGV軟體ETF昨跌近3%', 'size': 14, 'color': ACCENT_RED, 'bold': True},
        {'text': '  AI商業模式仍在調整期', 'size': 13, 'color': DARK_GRAY},
        {'text': '  不確定AI是否顛覆既有軟體業', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  短線警示：費半技術指標', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  17天未跌 = 籌碼有點擁擠', 'size': 13, 'color': ACCENT_ORANGE},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  IGV目前在強勢築底', 'size': 13, 'color': CORP_BLUE},
        {'text': '  抄底資金太多暫時跌不深', 'size': 12, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.5), Inches(5.3),
                            hw_sw_lines)


def slide_08_fund_flow(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '資金輪動 — AI硬體獨佔市場', '單一敘事風險：只看AI硬體 能源/消費/公用事業全遭抽離')
    _add_footer(slide, 8)

    img = charts['peg_trend']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(8), Inches(3.5))

    _add_rounded_rect(slide, Inches(8.6), Inches(1.3), Inches(4.4), Inches(3.5), WHITE)
    flow_lines = [
        {'text': '市場共識脫節警訊', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  市場只剩一個敘事：AI硬體', 'size': 14, 'color': ACCENT_RED, 'bold': True},
        {'text': '  消費/軟體/能源/公用事業', 'size': 13, 'color': DARK_GRAY},
        {'text': '  全面遭資金撤出', 'size': 13, 'color': ACCENT_RED},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  跨資產出現明顯背離', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  利率/油價/中東地緣政治', 'size': 13, 'color': DARK_GRAY},
        {'text': '  市場幾乎全部忽略', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  單邊上攻動能仍強', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  但需注意共識過於集中', 'size': 13, 'color': ACCENT_ORANGE},
    ]
    _add_multiline_text_box(slide, Inches(8.8), Inches(1.45), Inches(4.0), Inches(3.2),
                            flow_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    bottom = [
        {'text': '大股票 vs 中小型股的警示', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  若最終只是大股票的勝利，大型股與中小股持續脫節 → 影響實體經濟。大多數就業仍由小型公司提供', 'size': 14, 'color': DARK_GRAY},
        {'text': '  納指創高 vs IGV軟體脫鉤從2025年下半年開始，是重要的市場結構警訊', 'size': 14, 'color': ACCENT_RED, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3), bottom)


def slide_09_tesla_bitcoin(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '特斯拉財報 & 比特幣資金動態', '特斯拉EPS連4季低於預期 | 散戶每日買入70億 | BTC卡在100日均線')
    _add_footer(slide, 9)

    img = charts['taiwan_export']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.5))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(5.0), Inches(3.5), WHITE)
    tsla_btc_lines = [
        {'text': '特斯拉Q1 2026重點', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  自由現金流回到正值', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  但EPS僅$0.25（預期$0.60）', 'size': 13, 'color': ACCENT_RED},
        {'text': '  連續四季低於市場預期', 'size': 13, 'color': ACCENT_RED},
        {'text': '  CAPEX增$50億至$250億', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  財報後股價跌3.56%', 'size': 14, 'color': ACCENT_RED, 'bold': True},
        {'text': '  散戶日買入約70億美元', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  BTC：卡在100日均線', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  突破75,000-80,000才有機會', 'size': 13, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.6), Inches(3.2),
                            tsla_btc_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    insight = [
        {'text': '散戶投資邏輯：「賣產值的就買」', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  特斯拉&比特幣：逢低買盤持續，但追高意願薄弱。電動車賣差無關緊要，市場等待自駕/robotaxi的營收貢獻', 'size': 14, 'color': DARK_GRAY},
        {'text': '  → 掘金熱邏輯：最後能不能挖到金不重要，賣鏟子的就繼續買。硬體 > 軟體 > 應用服務', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3), insight)


def slide_10_part_three(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_section_divider(slide, '03', 'PART THREE.', '台灣市場 — 全面升溫',
                         '外銷訂單首破900億 | 金管會三線鬆綁 | 主動ETF全面跑贏大盤')


def slide_11_taiwan_exports(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '台灣外銷訂單 — 3月首破900億美元大關', '3月份911億美元 | 年增幅+65.9% | Q1累計2,319億 | 美台逆差全球第一')
    _add_footer(slide, 11)

    img = charts['capex_war']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.8))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(4.9), Inches(5.5), WHITE)
    export_lines = [
        {'text': '台灣外銷訂單亮點', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  3月份：911億美元（歷史首次破900億）', 'size': 13, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  年增幅：+65.9%（爆炸性成長）', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  Q1累計：2,319億，平均增速50%', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '  去年全年增速已維持20-30%', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  美國對台灣進口年增596億', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  全球第一（中國減971億）', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  AI投資未放緩，反而加速', 'size': 13, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  地緣政治衝突未影響AI需求', 'size': 12, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.5), Inches(5.3),
                            export_lines)


def slide_12_fsc_policy(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '金管會政策鬆綁 — 台股新時代', '單一持股上限10%→25% | 外債擔保換台幣 | 股息美元化 | 亞洲資管中心')
    _add_footer(slide, 12)

    img = charts['capex_ratio']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.5))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(5.0), Inches(3.5), WHITE)
    fsc_lines = [
        {'text': '金管會三線並進', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  ① 制度鬆綁', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '  持股上限10%→25%（台積電受益）', 'size': 13, 'color': DARK_GRAY},
        {'text': '  外國債券Triple B+可擔保換台幣', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  ② 場域建設', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '  高雄亞洲資產管理中心建立', 'size': 13, 'color': DARK_GRAY},
        {'text': '  銀行/保險/資管陸續進駐', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  ③ 產品和服務', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '  TISA長期投資帳戶、稅賦優惠', 'size': 13, 'color': DARK_GRAY},
        {'text': '  家族辦公室/跨境投資', 'size': 13, 'color': ACCENT_GREEN},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.6), Inches(3.2),
                            fsc_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    insight = [
        {'text': '核心：台灣從資金輸出地→資金停留地', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  持股上限放寬 → 主動ETF更貼近市場權重 → 台積電ETF回補力道增強 → 帶動指數進一步上揚', 'size': 14, 'color': DARK_GRAY},
        {'text': '  外資可不換匯直接配息 → 降低台股匯率風險 → 增加外資持有台股誘因', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3), insight)


def slide_13_active_etf(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '主動型ETF全面跑贏大盤', '992A/981A/994A 突破60% | 均值+43% | 持股集中AI板塊')
    _add_footer(slide, 13)

    img = charts['global_stocks_ytd']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.5))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(5.0), Inches(3.5), WHITE)
    etf_lines = [
        {'text': '主動ETF為何跑贏？', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  3檔ETF突破60%（992A/981A/994A）', 'size': 13, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '  均值+43%，明顯高於大盤', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  原因：AI相關資金高度集中', 'size': 14, 'color': CORP_BLUE, 'bold': True},
        {'text': '  集中持股 = 貝塔值放大效應', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  注意：台積電持股上限鬆綁', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  10%→25%，未來更大受益', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  熊市時貝塔值放大 = 跌更多', 'size': 12, 'color': ACCENT_RED},
        {'text': '  了解風險，才能正確使用', 'size': 12, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.6), Inches(3.2),
                            etf_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    note_lines = [
        {'text': '「ETF永動機」現象解析', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  台積電股價漲太快 + 持股上限10% → 主動ETF買不夠 → 被迫買0050/0052 → ETF買ETF（永動機現象）', 'size': 14, 'color': DARK_GRAY},
        {'text': '  持股上限放寬後 → 主動ETF可直接配置更多台積電 → 績效更貼近市場 → 永動機現象消失', 'size': 14, 'color': ACCENT_GREEN, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3),
                            note_lines)


def slide_14_part_four(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_section_divider(slide, '04', 'PART FOUR.', '投資策略 — 換股邏輯與配置',
                         '硬體主線確認 | 台股換股完成後的機會 | 油價PMI | 外資觀察')


def slide_15_macro_energy(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '宏觀環境 — 油價回百美元 & PMI緩升', '中東僵局加劇 | XLE能源股築底 | 企業備貨推動PMI 但實質需求偏弱')
    _add_footer(slide, 15)

    img = charts['commodity_ytd']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(7.5), Inches(3.8))

    _add_rounded_rect(slide, Inches(8.1), Inches(1.3), Inches(4.9), Inches(5.5), WHITE)
    macro_lines = [
        {'text': '宏觀三大警訊', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  油價重回$100：通膨隱患', 'size': 14, 'color': ACCENT_RED, 'bold': True},
        {'text': '  美伊談判陷入僵局', 'size': 13, 'color': DARK_GRAY},
        {'text': '  革命衛隊退出談判', 'size': 13, 'color': ACCENT_RED},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  PMI回升：備貨≠實質需求', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
        {'text': '  企業預期通膨而搶貨', 'size': 13, 'color': DARK_GRAY},
        {'text': '  終端需求仍偏弱', 'size': 13, 'color': DARK_GRAY},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  美國實質薪資已轉負', 'size': 14, 'color': ACCENT_RED, 'bold': True},
        {'text': '  就業看似穩但不敢辭職', 'size': 13, 'color': DARK_GRAY},
        {'text': '  消費信心持續下滑', 'size': 13, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(8.3), Inches(1.45), Inches(4.5), Inches(5.3),
                            macro_lines)


def slide_16_stock_rotation(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '台股換股完成後的機會', '外資賣超4,371億 | 外銷訂單增速65.9% | 大型股主導行情')
    _add_footer(slide, 16)

    img = charts['asset_ytd']
    slide.shapes.add_picture(img, Inches(0.2), Inches(1.3), Inches(6.8), Inches(3.5))

    img2 = charts['sp500_earnings']
    slide.shapes.add_picture(img2, Inches(7.2), Inches(1.3), Inches(5.8), Inches(3.5))

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), WHITE)
    note_lines = [
        {'text': '台股行情的三大驅動力', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  ① 外銷訂單：增速65.9%，AI投資持續推動，台灣科技出口成全球最大受益方', 'size': 13, 'color': DARK_GRAY},
        {'text': '  ② 金管會鬆綁：台積電持股上限擴大，ETF回補力道增強，亞洲資管中心逐步建立', 'size': 13, 'color': DARK_GRAY},
        {'text': '  ③ 外資回補潛力：賣超已逾4,371億，一旦外資轉向，將是台股最大正面觸媒', 'size': 14, 'color': ACCENT_ORANGE, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3),
                            note_lines)


def slide_17_part_five(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_section_divider(slide, '05', 'PART FIVE.', '投資策略總結 — 2026.04.24 操作建議',
                         '硬體主線配置 | 資產配置建議 | 風險評估 | 後市觀察重點')


def slide_18_recommendations(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '產業動能雷達 & 配置建議', 'AI硬體主線持續 | 換股後等大型股確認 | 外資觀察是關鍵')
    _add_footer(slide, 18)

    img = charts['sector_radar']
    slide.shapes.add_picture(img, Inches(0.2), Inches(1.3), Inches(5.2), Inches(5.2))

    _add_rounded_rect(slide, Inches(5.6), Inches(1.3), Inches(7.4), Inches(2.3), WHITE)
    tw_rec_lines = [
        {'text': '積極型配置建議', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  AI硬體/台積電 35%（費半主線確認）', 'size': 13, 'color': CHARCOAL},
        {'text': '  主動型ETF 25%（均值+43% 持續跑贏）', 'size': 13, 'color': CHARCOAL},
        {'text': '  現金 30%（等外資回補確認信號）', 'size': 13, 'color': CHARCOAL},
        {'text': '  能源/防禦 10%（油價支撐，築底觀察）', 'size': 13, 'color': CHARCOAL},
    ]
    _add_multiline_text_box(slide, Inches(5.8), Inches(1.4), Inches(7), Inches(2.1),
                            tw_rec_lines)

    _add_rounded_rect(slide, Inches(5.6), Inches(3.8), Inches(7.4), Inches(3.0), WHITE)
    us_rec_lines = [
        {'text': '穩健型配置建議', 'size': 17, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  現金 35%（等換股確認完成後進場）', 'size': 13, 'color': CHARCOAL},
        {'text': '  台積電/大型權值 25%（換股後主力）', 'size': 13, 'color': CHARCOAL},
        {'text': '  防禦型科技ETF 20%（0050/台灣50）', 'size': 13, 'color': CHARCOAL},
        {'text': '  債券/黃金 10%（系統性對沖）', 'size': 13, 'color': CHARCOAL},
        {'text': '  其他 10%', 'size': 13, 'color': CHARCOAL},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  注意：費半17連紅後短線擁擠，別追', 'size': 11, 'color': ACCENT_ORANGE, 'bold': True},
    ]
    _add_multiline_text_box(slide, Inches(5.8), Inches(3.9), Inches(7), Inches(2.8),
                            us_rec_lines)


def slide_19_allocation(prs, charts):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, LIGHT_GRAY)
    _add_header_bar(slide, '資產配置圖解 & 後市觀察重點', '2026.04.24 — AI硬體主線確立，外資回補仍是最大觸媒')
    _add_footer(slide, 19)

    img = charts['allocation']
    slide.shapes.add_picture(img, Inches(0.3), Inches(1.3), Inches(8.5), Inches(3.5))

    _add_rounded_rect(slide, Inches(9.0), Inches(1.3), Inches(4.0), Inches(3.5), WHITE)
    risk_lines = [
        {'text': '後市觀察重點', 'size': 16, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  1. 費半籌碼是否鬆動', 'size': 13, 'color': CHARCOAL, 'bold': True},
        {'text': '  17連紅後短線擁擠待觀察', 'size': 12, 'color': MID_GRAY},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  2. 台股換股完成確認', 'size': 13, 'color': CHARCOAL, 'bold': True},
        {'text': '  大型股是否接棒中小型股', 'size': 12, 'color': MID_GRAY},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  3. 油價持守$100關卡', 'size': 13, 'color': CHARCOAL, 'bold': True},
        {'text': '  公債殖利率是否連帶反應', 'size': 12, 'color': MID_GRAY},
        {'text': '', 'size': 4, 'color': DARK_GRAY},
        {'text': '  4. 外資回補信號出現', 'size': 13, 'color': CHARCOAL, 'bold': True},
        {'text': '  4,371億賣超待回補', 'size': 12, 'color': MID_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(9.2), Inches(1.4), Inches(3.6), Inches(3.3),
                            risk_lines)

    _add_rounded_rect(slide, Inches(0.3), Inches(5.1), Inches(12.7), Inches(1.7), CORP_BLUE)
    summary = [
        {'text': '核心策略總結', 'size': 18, 'color': WHITE, 'bold': True},
        {'text': '', 'size': 5, 'color': WHITE},
        {'text': '  費半破萬點確認AI硬體主線；台股換股 = 行情升級非結束。外銷訂單+65.9%是台股最強基本面支撐', 'size': 15, 'color': WHITE},
        {'text': '  外資4,371億賣超回補是最大潛在觸媒，金管會鬆綁持續加油，台股長期多頭格局未變', 'size': 14, 'color': SKY_BLUE},
    ]
    _add_multiline_text_box(slide, Inches(0.6), Inches(5.25), Inches(12), Inches(1.3), summary)


def slide_20_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg_rect(slide, Inches(0), Inches(0), SLIDE_W, SLIDE_H, CORP_BLUE)
    _add_circle(slide, Inches(10.5), Inches(-0.5), Inches(3), CORP_BLUE_DARK)
    _add_circle(slide, Inches(-0.8), Inches(5.5), Inches(2), CORP_BLUE_DARK)

    _add_text_box(slide, Inches(1), Inches(0.5), Inches(11), Inches(0.8),
                  '2026.04.24 重點回顧 & 後市展望', font_size=32, font_color=WHITE,
                  bold=True, alignment=PP_ALIGN.CENTER)
    _add_bg_rect(slide, Inches(5), Inches(1.3), Inches(3.3), Inches(0.04), WHITE)

    _add_rounded_rect(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(3.2), WHITE)
    summary_lines = [
        {'text': '本集五大重點回顧', 'size': 18, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  1. 費半破萬點連17紅：+1.71%收10,078點，三大指數收黑。硬體爆發vs軟體失速，AI基建需求4-5年結構循環確認', 'size': 14, 'color': DARK_GRAY},
        {'text': '  2. 台股多殺多換股：成交1.3兆創歷史新高，外資僅賣27億，資金從中小型股流向大型權值。行情升級非結束', 'size': 14, 'color': DARK_GRAY},
        {'text': '  3. 台灣外銷訂單創紀錄：3月份911億美元首破900億大關，年增+65.9%，Q1累計2,319億，AI投資加速不減', 'size': 14, 'color': DARK_GRAY},
        {'text': '  4. 金管會三線並進：持股上限10%→25%，外債換台幣，美元計價股息，亞洲資管中心建立加速', 'size': 14, 'color': DARK_GRAY},
        {'text': '  5. 主動ETF全面跑贏：992A/981A/994A突破60%，均值+43%，持股集中AI板塊放大貝塔效應', 'size': 14, 'color': DARK_GRAY},
    ]
    _add_multiline_text_box(slide, Inches(0.8), Inches(1.7), Inches(11.7), Inches(3.0),
                            summary_lines)

    _add_rounded_rect(slide, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.8), WHITE)
    next_lines = [
        {'text': '後市觀察重點', 'size': 18, 'color': CORP_BLUE, 'bold': True},
        {'text': '', 'size': 5, 'color': DARK_GRAY},
        {'text': '  費半連17紅後籌碼擁擠程度（短線整理空間）；台股換股能否順利完成，大型股接棒訊號', 'size': 14, 'color': CHARCOAL},
        {'text': '  油價$100是否持守（通膨風險 → 公債殖利率影響）；外資4,371億賣超是否開始回補', 'size': 14, 'color': CHARCOAL},
        {'text': '  台灣外銷訂單是否維持高增速；金管會政策落地效果與亞洲資管中心實質進駐進度', 'size': 14, 'color': CHARCOAL},
    ]
    _add_multiline_text_box(slide, Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.6),
                            next_lines)

    _add_text_box(slide, Inches(2), Inches(7.0), Inches(9.3), Inches(0.4),
                  '資料來源：早晨財經速解讀 | 分析日期：2026-04-24 | 本報告僅供參考，非投資建議',
                  font_size=11, font_color=SKY_BLUE, alignment=PP_ALIGN.CENTER)


# ============================================================
# Main Build
# ============================================================

def build_presentation():
    setup_matplotlib()
    print('Generating charts...')
    charts = gen_charts()

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    print('Building slides...')

    slide_01_cover(prs)
    print('  [1/20] Cover')

    slide_02_contents(prs)
    print('  [2/20] Contents')

    slide_03_part_one(prs)
    print('  [3/20] Part 1 divider')

    slide_04_us_weekly(prs, charts)
    print('  [4/20] US indices — SOX 10,000 milestone')

    slide_05_taiwan_market(prs, charts)
    print('  [5/20] Taiwan market / volume record')

    slide_06_part_two(prs)
    print('  [6/20] Part 2 divider')

    slide_07_hardware_software(prs, charts)
    print('  [7/20] Hardware vs Software divergence')

    slide_08_fund_flow(prs, charts)
    print('  [8/20] Fund flow rotation')

    slide_09_tesla_bitcoin(prs, charts)
    print('  [9/20] Tesla earnings & Bitcoin')

    slide_10_part_three(prs)
    print('  [10/20] Part 3 divider')

    slide_11_taiwan_exports(prs, charts)
    print('  [11/20] Taiwan export orders record')

    slide_12_fsc_policy(prs, charts)
    print('  [12/20] FSC policy deregulation')

    slide_13_active_etf(prs, charts)
    print('  [13/20] Active ETF performance')

    slide_14_part_four(prs)
    print('  [14/20] Part 4 divider')

    slide_15_macro_energy(prs, charts)
    print('  [15/20] Macro — oil price & PMI')

    slide_16_stock_rotation(prs, charts)
    print('  [16/20] Taiwan stock rotation analysis')

    slide_17_part_five(prs)
    print('  [17/20] Part 5 divider')

    slide_18_recommendations(prs, charts)
    print('  [18/20] Investment recommendations')

    slide_19_allocation(prs, charts)
    print('  [19/20] Asset allocation')

    slide_20_summary(prs)
    print('  [20/20] Summary & outlook')

    prs.save(OUTPUT_PPTX)
    print(f'\n✅ Saved: {OUTPUT_PPTX}')


if __name__ == '__main__':
    build_presentation()

---
name: weekly-market-report
description: >
  Generate a comprehensive weekly financial market analysis PPTX presentation from SRT subtitle files.
  Use when the user asks to create a weekly market report, weekly stock analysis, 財金週報, 股市週報,
  or wants to analyze SRT files from financial YouTube channels (財金號角, 股癌投資) and produce a
  professional PPTX slide deck. Triggers on 週報, weekly report, market analysis PPTX, SRT 分析簡報,
  財金簡報, or when SRT files are present in 財金股市分析/ directory and user requests a presentation.
---

# 財金股市週報生成

## Overview

Generate a 20-slide Corporate Blue & White PPTX weekly market report by extracting data from financial YouTube channel SRT subtitle files, synthesizing market data, and producing professional charts with matplotlib.

## Workflow

### Step 1: Locate and Read SRT Files

Scan the project's SRT directories for the current week's files:

```
財金股市分析/
├── 財金號角/          ← 早晨財經速解讀（每日）
├── 股癌投資/          ← 股癌 Podcast（不定期）
└── [其他頻道]/        ← 依新增來源擴充
```

Read all SRT files from the target week. For data extraction patterns, see [references/data_extraction_guide.md](references/data_extraction_guide.md).

### Step 2: Extract Market Data

From each SRT, extract:
- **US indices**: DJIA, S&P500, NASDAQ, SOX — closing price and weekly change %
- **Asia markets**: TAIEX, Nikkei, KOSPI — weekly performance
- **Crypto**: BTC, ETH — price and key events
- **Macro indicators**: PMI, employment, CPI, Fed policy
- **Sector trends**: Identify 4-6 hot sectors with catalysts
- **Key events**: Policy changes, elections, earnings, geopolitical

### Step 3: Update the Generation Script

Read and modify the existing `scripts/generate_weekly_report.py` with the week's extracted data. Key sections to update:

1. **Docstring** — Update week number and date range
2. **Market data dicts** — Replace hardcoded values with current week's numbers
3. **Chart data** — Update all chart generation functions with new data
4. **Text content** — Update bullet points, analysis, and commentary on each slide
5. **Sector analysis** — Adjust radar chart dimensions and featured sectors
6. **Strategy section** — Update allocation recommendations based on current conditions

### Step 4: Generate PPTX

Run the script:

```bash
python3 src/main/python/generate_weekly_report.py
```

Output: `output/reports/weekly_report.pptx`

Dependencies: `python-pptx`, `matplotlib`, `numpy`

### Step 5: Verify and Commit

1. Confirm the PPTX was generated at `output/reports/weekly_report.pptx`
2. Git add with `-f` flag (output/ may be in .gitignore): `git add -f output/reports/weekly_report.pptx`
3. Commit and push to GitHub

## Slide Structure (20 slides)

| # | Section | Content |
|---|---------|---------|
| 1 | Cover | Title, date range, disclaimer |
| 2 | Contents | 5-part table of contents |
| 3-4 | Part 1: Market Overview | US index weekly bar chart, Asia market comparison |
| 5-7 | Part 2: Key Events | Trump/policy impact, Japan election, BTC crash chart |
| 8-10 | Part 3: Macro Economy | PMI gauge, Value vs Growth rotation, AI job displacement |
| 11-13 | Part 4: Sector Trends | LEO satellite & cooling tech, Capex war chart, Sector radar |
| 14-16 | Part 5: Strategy | Top picks with rationale, Asset allocation pie, Risk assessment |
| 17-18 | Appendix | Data sources, Drawdown comparison chart |
| 19 | Summary | Key takeaways and next week outlook |
| 20 | Closing | Disclaimer and contact |

## Design Specifications

- **Theme**: Corporate Blue & White (`#2B579A` primary)
- **Slide size**: Widescreen 16:9 (13.333" x 7.5")
- **Chinese font**: Microsoft JhengHei (PPTX) / Heiti TC (matplotlib)
- **English font**: Arial
- **Charts**: matplotlib → PNG BytesIO → embedded in slides

For complete color scheme and chart implementation patterns, see [references/chart_patterns.md](references/chart_patterns.md).

## Resources

### scripts/
- `generate_weekly_report.py` — Main PPTX generation script. Read this file, update data and text content for the current week, then execute to produce the weekly report.

### references/
- `data_extraction_guide.md` — SRT data extraction process, data categories, and structured data templates
- `chart_patterns.md` — Chart types, matplotlib configuration, Corporate Blue color scheme, and PPTX embedding patterns

# CLAUDE.md - 財金股市分析系統

> **Documentation Version**: 1.0
> **Last Updated**: 2026-04-19
> **Project**: 財金股市分析自動化系統
> **Description**: 將 YouTube / Podcast SRT 字幕轉換為專業財金投資分析 PPTX 簡報
> **Tech Stack**: Python (python-pptx, matplotlib) + Node.js (PptxGenJS)

---

## 🚨 CRITICAL RULES - READ FIRST

### 🔄 RULE ACKNOWLEDGMENT REQUIRED
> Before starting ANY task, respond with:
> "✅ CRITICAL RULES ACKNOWLEDGED - I will follow all prohibitions and requirements listed in CLAUDE.md"

### ❌ ABSOLUTE PROHIBITIONS
- **NEVER** create files in root directory → use `src/` or `output/`
- **NEVER** write PPTX output to root → always use `output/reports/`
- **NEVER** create duplicate scripts (generate_v2.py, enhanced_xyz.py) → extend existing files
- **NEVER** hardcode SRT file paths → use arguments or config
- **NEVER** use `grep`, `find`, `cat`, `head`, `tail`, `ls` shell commands → use Grep, Glob, Read, Bash tools
- **NEVER** create documentation (.md) unless explicitly requested

### ✅ MANDATORY REQUIREMENTS
- **COMMIT** after every completed report generation
- **GITHUB BACKUP** after every commit: `git push origin main`
- **READ scripts first** before editing — Edit/Write tools require prior Read
- **USE TASK AGENTS** for operations >30 seconds
- **SEARCH FIRST** before creating new scripts — check for existing implementations

---

## 🏗️ PROJECT STRUCTURE

```
財金股市分析/
├── src/
│   └── main/
│       └── python/                         # 報告生成腳本
│           ├── generate_weekly_report.py           # 週報生成（主流程）
│           ├── generate_market_analysis_slides.py  # Q1 市場分析簡報
│           ├── generate_stock_analysis_pptx.py     # 個股分析 PPTX
│           └── generate_optical_stock_comparison.py # 光通訊選股對比
│
├── output/                                 # 所有輸出檔案
│   ├── reports/                            # 輸出 PPTX（主要產出）
│   ├── slides/                             # PNG 幻燈片
│   ├── charts/                             # 圖表 PNG
│   └── temp/                              # 暫存檔案
│
├── 財金號角/                               # 早晨財經速解讀 SRT 字幕
├── 股癌投資/                               # 股癌 Podcast SRT 字幕
│   └── output/                            # 股癌系列 PPTX 輸出
│
├── weekly-market-report.skill             # Claude Skill 定義
├── CLAUDE.md                              # 本文件
└── .claude/settings.local.json           # Claude Code 權限設定
```

---

## 🔄 WORKFLOW

### 標準週報流程

```
SRT 字幕檔（財金號角/ 或 股癌投資/）
    ↓
Python 腳本解析字幕內容
    ├── 提取市場數據與分析觀點
    ├── 準備圖表數據
    └── 生成 matplotlib PNG 圖表
    ↓
python-pptx 組合 PPTX
    └── 輸出至 output/reports/YYYYMMDD_report.pptx

備用路線（Node.js）：
SRT → Node.js PptxGenJS → PPTX
```

### Skill 使用
- 使用 `/weekly-market-report` skill 啟動標準週報生成流程
- Skill 定義位於 `weekly-market-report.skill`

---

## 📂 KEY SCRIPTS

| 腳本 | 用途 | 輸入 | 輸出 |
|------|------|------|------|
| `generate_weekly_report.py` | 週報生成（主流程）| SRT 字幕 | `output/reports/YYYYMMDD_report.pptx` |
| `generate_market_analysis_slides.py` | 季度市場分析簡報 | 手動數據 | `output/reports/YYYYQ*_*.pptx` |
| `generate_stock_analysis_pptx.py` | 個股深度分析 | SRT / 手動 | `output/reports/YYYYMMDD_主題.pptx` |
| `generate_optical_stock_comparison.py` | 光通訊選股對比 | 手動數據 | `output/reports/optical_*.pptx` |

---

## 🏷️ NAMING CONVENTIONS

### SRT 來源檔
- 財金號角：`YYYY_M_D(星期)_標題關鍵字.srt`
- 股癌投資：`EP{號碼} _ {emoji}.srt`

### 輸出 PPTX
- 週報：`output/reports/YYYYMMDD_report.pptx`
- 主題報告：`output/reports/YYYYMMDD_主題名稱.pptx`
- 季報：`output/reports/YYYYQ{季}_市場分析.pptx`

### 腳本命名
- 格式：`generate_{類型}_{主題}.py`
- 禁止使用：`_v2`, `_new`, `_enhanced`, `_improved`

---

## 🚀 COMMON COMMANDS

```bash
# 生成週報（指定 SRT 檔案）
python src/main/python/generate_weekly_report.py

# 生成個股分析
python src/main/python/generate_stock_analysis_pptx.py

# 生成市場分析簡報
python src/main/python/generate_market_analysis_slides.py

# Git 備份完整流程
git add output/reports/ src/
git commit -m "新增 YYYYMMDD 財金週報 PPTX：標題"
git push origin main

# 安裝 Python 依賴
pip install python-pptx matplotlib numpy
```

---

## 🎯 PRE-TASK COMPLIANCE CHECK

Before starting ANY task:

- [ ] ✅ I acknowledge all critical rules in CLAUDE.md
- [ ] Will output go to `output/reports/`? (not root)
- [ ] Is there an existing script to extend? (search first)
- [ ] Will this take >30 seconds? → use Task agents
- [ ] Is this 3+ steps? → use TodoWrite breakdown

---

**Focus**: Single source of truth per script. Extend existing, don't duplicate.

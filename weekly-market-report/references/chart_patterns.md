# 圖表生成模式參考

## 使用的圖表類型

週報使用 matplotlib 生成以下圖表，存為 PNG 後嵌入 PPTX：

### 1. 水平長條圖（美股指數週漲跌）
- 用於展示多個指數的漲跌比較
- 綠色表示上漲，紅色表示下跌
- 標籤顯示百分比數值

### 2. 分組長條圖（亞洲市場對比）
- 並列本週 vs 前週漲跌
- 適合展示 2-4 個市場的對比

### 3. 折線圖（BTC 價格走勢）
- 模擬日線走勢
- 搭配關鍵事件標註（如「崩盤」標記）
- 使用紅色虛線標示重要支撐/阻力位

### 4. 儀表板風格指標（ISM PMI）
- 使用半圓環或進度條風格
- 區分「收縮」(< 50) vs「擴張」(> 50) 區間
- 顏色從紅→黃→綠漸變

### 5. 雷達圖（產業評分）
- 評估各產業的多維度表現
- 維度：成長性、政策支持、資金面、技術面、風險
- 適合 4-6 個產業同時比較

### 6. 圓餅圖（資產配置建議）
- 顯示建議投資組合比例
- 標籤包含類別名稱和百分比

### 7. 瀑布圖/堆疊圖（Capex 支出比較）
- 比較各大科技公司的資本支出
- 區分 AI 相關 vs 非 AI 支出

## Matplotlib 中文字型設定

```python
plt.rcParams['font.family'] = ['Heiti TC', 'Arial Unicode MS', 'Hiragino Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
```

macOS 優先使用 `Heiti TC`，備選 `Arial Unicode MS`。

## 圖表嵌入 PPTX 流程

```python
from io import BytesIO

def fig_to_image_stream(fig, dpi=250):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight',
                facecolor='white', edgecolor='none', pad_inches=0.15)
    plt.close(fig)
    buf.seek(0)
    return buf

# 嵌入到投影片
stream = fig_to_image_stream(fig)
slide.shapes.add_picture(stream, left, top, width, height)
```

## Corporate Blue 色彩方案

| 變數名稱 | Hex | 用途 |
|----------|-----|------|
| CORP_BLUE | #2B579A | 主色、標題背景 |
| CORP_BLUE_DARK | #1A3A6C | 深色強調 |
| CORP_BLUE_LIGHT | #4A86C8 | 圖表次色 |
| ACCENT_BLUE | #5BA0D6 | 輔助藍色 |
| SKY_BLUE | #D6E8F7 | 淺背景 |
| ACCENT_RED | #E84D4D | 下跌/警示 |
| ACCENT_GREEN | #27AE60 | 上漲/正面 |
| ACCENT_ORANGE | #F59E0B | 警告/中性 |
| ACCENT_PURPLE | #8E44AD | 特殊標記 |

## 投影片尺寸

- 寬螢幕 16:9：13.333" x 7.5"（`Inches(13.333)` x `Inches(7.5)`）
- python-pptx PPTX 字型：`Microsoft JhengHei`（中文）、`Arial`（英文）

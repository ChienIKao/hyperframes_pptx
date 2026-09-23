# Markdown 撰寫規格

`build_deck.py` 讀的是一般 Markdown，加上少量約定。

## 標題階層 → 投影片結構（預設）

| Markdown | 產生 |
|---|---|
| `# 標題` | **封面** |
| `## 章節 \| English` | **大綱**中的一項 + **章節頁**（右側列出該章的子章節） |
| `### 子章節` | 章節頁子目錄中的一項 + 頂部導覽列第二排 |
| `#### 投影片標題 \| 副標` | 一張**內容投影片** |
| `#####` 以下 | 投影片內的粗體小標 |
| `---`（在投影片內） | 強制換頁，下一頁沿用同標題 |

- 任何標題都可以用 ` | ` 接副標：章節的英文名（`## 緒論 | Introduction`）、投影片的灰色副標。
- `###` 底下直接寫內容（沒有 `####`）→ 以子章節名稱為標題產生一張投影片。所以三層的稿子也能用。
- 內容投影片標題會自動組成「**子章節 – 投影片標題**」（例：「研究背景 – 消費習慣的轉變」）。兩者相同時只顯示一次。格式可用 `title_format` 改。
- 模版模式（使用 .pptx 模版）沒有導覽列與子目錄：`##` 章節頁、`###`/`####` 都變成內容頁。

### 其他階層配置

```yaml
subsection_level: 0   # 不要子章節：## 章節、### 投影片
slide_level: 3
```

```yaml
section_level: 0      # 完全扁平：## 就是投影片
subsection_level: 0
slide_level: 2
```

## Front matter（全部可省略）

```yaml
---
title: 簡報標題              # 沒有 `#` 時使用
subtitle: 英文或副標題
author: 報告人：講者姓名
date: 2026 / 01 / 01
theme: blockframe            # themes/<slug>.json
template: my-company         # 或改用 .pptx 模版（與 theme 擇一）
eyebrow: 期末專題報告          # 封面左上標籤
badge: 2026                  # 封面星爆裡的字（canvas 主題）
cover_image: assets/cover.png
logo_left: assets/lab.png    # 導覽列左右 logo
logo_right: assets/school.png
outline: true                # 大綱頁
outline_title: 大綱
outline_en: OUTLINE
recap: false                 # true：每個子章節開始前重放一次章節頁並標出目前位置
title_format: "{sub} – {title}"
closing: Q&A                 # 結尾頁文字；false 不產生
closing_label: THANK YOU
max_table_rows: 9            # 表格超過此列數自動拆頁（表頭重複）
section_level: 2
subsection_level: 3
slide_level: 4
---
```

以 `#` 開頭的 front matter 行是註解。

## 內容元素

| 寫法 | 效果 |
|---|---|
| `- 項目`（縮排 2 或 4 格為子層） | 條列 |
| `1. 項目` | 編號條列 |
| 一般段落 | 無項目符號的文字 |
| `**粗體**`、`*斜體*`、`` `code` ``、`[文字](url)` | 行內格式 |
| `==關鍵詞==` | 色塊強調（螢光筆效果）＋粗體 |
| `> 引言` | 引言；整頁只有一段引言時成為引言卡片頁 |
| `> [!NOTE] 文字` | 底部提示框。種類：`NOTE`（註）、`TIP`、`IMPORTANT`（重點）、`WARNING`／`CAUTION`（注意）、`SUMMARY`（小結） |
| `資料來源：…` 開頭的段落 | 頁尾小字出處（也認 `來源：`、`參考資料：`、`Source:`、`Ref:`） |
| `![圖說](assets/a.png)`（單獨一行） | 加框圖片＋圖說標籤 |
| Markdown 表格 | 原生表格；○ × ✓ 與數字自動置中 |
| ```` ```bash ```` 等程式碼區塊 | 深色程式碼卡片 |
| ```` ```cards ```` ```` ```steps ```` ```` ```timeline ```` ```` ```flow ```` ```` ```stats ```` ```` ```compare ```` ```` ```chart ```` | **資訊圖表元件** → 見 `infographics.md` |
| `<!-- 講者備註 -->` | 講者備註（可多行） |

## 版面自動選擇（canvas 主題）

| 投影片內容 | 版面 |
|---|---|
| 只有文字 | 全寬文字（字級上限較大） |
| 只有段落且含 `==強調==`（≤4 段） | 宣言頁：置中大字 |
| 只有一段引言 | 引言卡片頁 |
| 文字 + 1 個圖／表／元件 | 左文右圖 |
| 只有圖／表／元件 | 全寬置中 |
| 提示框 | 固定在內容區底部 |

## 自動拆頁

依實際文字框大小估算：文字縮到 14pt 仍放不下就對半拆，後頁標題加「（續）」；拆點只落在頂層條列之間。多出來的圖表元件各自成頁。想自己控制就用 `---`。

## 素材

- 圖片路徑相對於 `deck.md`（或 `--assets` 指定的目錄）。
- 支援 PNG/JPG/GIF/BMP/TIFF；**SVG 不支援**，先轉 PNG。

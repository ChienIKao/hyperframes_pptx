---
name: md-to-pptx
description: 個人簡報流水線。輸入一份 Markdown 文稿（加上圖片等素材），依標題階層自動產生封面、大綱、章節子目錄、帶章節導覽列的內容頁與資訊圖表（卡片、流程、步驟、數據、比較、圖表），輸出可編輯的 .pptx。預設主題為 HyperFrames BlockFrame 糖果色塊，也支援 .pptx/.potx 模版。當使用者給 Markdown 稿子要做簡報／投影片／PPT／報告，或說「照這份稿子做 pptx」「用某某主題／模版出簡報」時使用。
---

你是這條流水線的**編排者**。輸入一份 Markdown 文稿 + 素材，產出一份可交付的 `.pptx`（見「交付清單」）。

**版面由 `scripts/build_deck.py` 決定性地產生**，你不手刻每一頁。你負責：
- 選主題或模版
- 把稿子整理成符合規格的 Markdown，**並把適合的段落轉成資訊圖表元件**
- 放素材
- 跑建置、做 QA、修正後重建

腳本做不到的客製（SmartArt、動畫、特殊版面）→ 先完成主流程，再用 `anthropic-skills:pptx` 對輸出檔做局部加工。

## 兩種引擎

| | canvas 主題（預設 `blockframe`） | 模版 |
|---|---|---|
| 外觀來源 | `themes/<slug>.json` 設計 token，程式直接繪製 | 使用者的 .pptx/.potx 母片 |
| 大綱頁、章節子目錄、頂部章節導覽列 | ✔ | ✘（只有目錄＋章節頁） |
| 資訊圖表元件 | ✔ 完整繪製 | 退化成條列／表格 |

## 交付清單（本 skill 的「完成」定義）

- [ ] `output/<slug>.pptx`：封面、大綱、章節頁（含子目錄）、內容頁（含導覽列）、結尾頁齊全，順序與稿子一致
- [ ] 每章都有英文副標（`## 緒論 | Introduction`）；稿子沒有的話，由你建議並在回報時列出
- [ ] `dump` 零警告（或每條警告都已確認為誤報並說明）
- [ ] 預覽 grid 已逐頁目視檢查：無溢出、元件沒有擠壓或大片留白、圖片比例正確
- [ ] 講者備註已帶入（稿子有 `<!-- -->` 時）
- [ ] 回報：投影片數、主題/模版、轉成元件的頁面清單、其他內容調整、佔位待補的數字

## 工作流骨架

一份簡報 = 一個專案目錄（當前工作目錄下，命名 `<YYYYMMDD-slug>/`）：

```
<YYYYMMDD-slug>/
  source.md        使用者原稿（不動）
  deck.md          工作稿：你整理後、實際拿去建置的版本
  assets/          素材（圖片、logo）
  output/<slug>.pptx
  output/<slug>_preview/grid.png
```

- **Step 0：開局問兩件事**（一次問完，有預設就給預設；使用者訊息中已講明的就不再問）
  1. **主題／模版**：列出 `references/templates-and-themes.md` 的清單。預設 `blockframe`。
  2. **內容處理方式**
     - 「照稿直出」（預設）：格式修正＋把內容轉成適合的元件，不改文字
     - 「幫我精煉」：長段落改條列，原文移到講者備註
- **Step 1：整理工作稿**。複製 `source.md` → `deck.md`，然後：
  - 依 `references/markdown-spec.md` 修正階層：`##` 章 → `###` 子章節 → `####` 投影片
  - 補章節英文副標、front matter（`eyebrow`、`author`、`date`、`logo_*`）
  - 依 `references/infographics.md` 逐頁判斷內容型態，改寫成元件（cards / flow / steps / stats / compare / chart / 提示框 / 宣言頁）
  - 把素材放進 `assets/` 並插入圖片。素材沒有對應位置 → 列出擺放提案讓使用者確認
- **Step 2：新模版先分析**（只在第一次用某個 .pptx 模版時）→ `references/templates-and-themes.md`「註冊新模版」。
- **Step 3：建置**
  ```bash
  python <skill>/scripts/build_deck.py build deck.md -o output/<slug>.pptx [--theme <slug|path>] [--template <slug|path>]
  ```
- **Step 4：QA** → `references/qa.md`
  1. `dump` 看警告
  2. `preview.py` 看 grid
  3. 回頭改 `deck.md`（不是改 pptx），然後重建。通常 1–2 輪
- **Step 5：交付**。依交付清單逐項核對後回報。

`<skill>` = 本 SKILL.md 所在目錄。

## 原則

- **改稿子，不改輸出檔。** 所有修正回到 `deck.md` 重建，確保可重現。
- **不擅改內容語意。** 轉元件、拆頁、補英文副標可以；刪字、改寫、換說法要先問。
- **不捏造數字。** `stats`／`chart` 的數值只能來自稿子；缺的用 `— 數值 —` 佔位並回報。
- **一頁一個視覺主角。** 左欄文字 ≤ 5 條，多了就拆頁或進備註。

## 參考

| 要做… | 讀 |
|---|---|
| Markdown 撰寫規格（階層、front matter、行內語法） | `references/markdown-spec.md` |
| 內容 → 資訊圖表元件對照、元件語法 | `references/infographics.md` |
| 選主題／模版、註冊新模版、從 HyperFrames 設計建立新主題 | `references/templates-and-themes.md` |
| 驗收與常見問題修正 | `references/qa.md` |

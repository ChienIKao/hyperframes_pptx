# 模版與主題

兩種外觀來源，二選一（有指定模版時以模版為準）：

| | 主題（theme） | 模版（template） |
|---|---|---|
| 是什麼 | 一個 JSON：設計 token（配色、字型、邊框、陰影） | 使用者的 .pptx / .potx（含母片、版面配置） |
| 放哪 | `themes/<slug>.json` | `templates/<slug>/` |
| 引擎 | `"engine": "canvas"` → 全部由程式繪製，有大綱、子目錄、導覽列、資訊圖表 | 套版面配置的佔位框 |
| 何時用 | 預設；想要設計感、完整的章節導覽 | 公司／學校有指定母片 |

> **開局時把以下兩張表完整列給使用者選。** 預設：主題 `blockframe`。

## 目前可用的主題

| slug | 引擎 | 風格 |
|---|---|---|
| **blockframe**（預設） | canvas | HyperFrames BlockFrame 新粗野主義：黑色粗邊框、零模糊硬陰影、五色糖果色塊（粉／藍／綠／黃／奶油）逐章輪替，直角、傾斜裝飾、星爆、斜紋塊。字型：Arial Black、Bahnschrift、微軟正黑體 |
| default | 基本 | 深藍封面／章節頁＋白底內容頁，橘色點綴。無導覽列 |

## 目前可用的模版

| slug | 說明 |
|---|---|
| （尚無，使用者提供後依下方流程註冊） | |

## 從 HyperFrames 設計建立新的 canvas 主題

HyperFrames 的設計規格在 `~/.claude/skills/hyperframes-creative/frame-presets/<slug>/FRAME.md`（例如 capsule、coral、editorial-forest）。移植步驟：

1. 複製 `themes/blockframe.json` 成 `themes/<slug>.json`。
2. 讀 FRAME.md 的 front matter，把值對應過來：
   - `colors` → `colors.palette`（章節輪替色用 `chapter_cycle`）、`ground`、`text`、`muted`
   - `borders` → `stroke.border`／`stroke.thin`：px ÷ 2 ≈ pt（1920px 寬 ≈ 960pt）
   - `shadows` → `stroke.shadow`／`stroke.thin_shadow`，同樣 px ÷ 2；要柔和、無硬陰影的風格就設 0
   - `typography` → `fonts`：Google 字型多半沒裝在本機，換成 Windows 內建的近似字型（襯線：Georgia／Cambria；無襯線：Arial／Segoe UI；等寬：Consolas）。中文一律 `Microsoft JhengHei`
3. 用 `examples/demo.md` 試建一次，看 grid 確認風格。
4. 把新主題加進上表。

裝飾元素（星爆、斜紋塊、傾斜卡片）目前是 BlockFrame 的語彙；移植到風格差很多的設計時，要在 `canvas_deck.py` 加主題開關，或者接受這些裝飾。

---

## 註冊新模版

1. 建目錄並放入檔案：`templates/<slug>/template.pptx`（或 `.potx`）。
2. 分析版面：
   ```bash
   python <skill>/scripts/build_deck.py inspect templates/<slug>/template.pptx --write-config templates/<slug>/config.json
   python <skill>/scripts/preview.py templates/<slug>/template.pptx     # 模版本身若有範例頁，可看其外觀
   ```
   `inspect` 會列出每個版面配置（名稱、佔位框類型與位置）與自動判定的角色對應。
3. 檢查並修正 `config.json` 的 `layouts` —— 角色對應到版面名稱（或索引）：

   | 角色 | 用途 | 需要的佔位框 |
   |---|---|---|
   | `cover` | 封面 | 標題 + 副標 |
   | `agenda` | 目錄 | 標題 + 內文 |
   | `section` | 章節分隔頁 | 標題 + （內文：放編號/副標） |
   | `content` | 一般內容 | 標題 + 1 個內文 |
   | `two_content` | 左文右圖 | 標題 + 2 個內文 |
   | `title_only` | 純圖/表 | 標題 |
   | `closing` | 結尾頁 | 標題 |

   自動判定依版面名稱關鍵字（中英文皆可）+ 佔位框結構。模版版面名稱不標準時，一定要手動確認。
4. 視需要調整 `config.json` 的其他欄位：
   ```json
   {
     "file": "template.pptx",
     "layouts": { "cover": "標題投影片", "content": "標題及內容", "...": "..." },
     "sizes": { "title": null, "body_max": 24, "body_min": 14, "split_below": 16 },
     "text_ratio": 0.45,
     "section_label": "{n:02d}",
     "theme": { "fonts": { "ea": "Microsoft JhengHei" } }
   }
   ```
   - `sizes.title: null` = 沿用模版字級。
   - `theme` 在模版模式只建議放字型（例如模版沒設中文字型時）；不要覆寫模版配色。
5. 用範例稿試跑一次（`examples/sample.md`），看 grid，確認每種角色都正常 → 把模版加進上面的「目前可用的模版」表。

模版內原有的投影片會在建置時全部移除，只保留母片與版面配置。

## 建立新主題

複製 `themes/default.json` 改值即可。欄位：

| 欄位 | 說明 |
|---|---|
| `fonts.heading` / `fonts.body` | 西文字型 |
| `fonts.ea`（或 `ea_heading` / `ea_body`） | 中文（東亞）字型，例如 `Microsoft JhengHei`、`Noto Sans TC` |
| `fonts.code` | 程式碼字型 |
| `colors.title/text/muted/accent` | 淺底頁的標題、內文、次要文字、強調色 |
| `colors.*_on_dark` | 深底頁（背景亮度低時自動切換）的對應顏色 |
| `colors.code_bg/code_text` | 程式碼區塊 |
| `backgrounds.cover/section/closing/content` | 各角色背景色 |
| `sizes.*` | 字級（pt）：`title`、`cover_title`、`subtitle`、`body_max`、`body_min` |
| `title_align` | `left` / `center`（內容頁標題） |
| `title_bold` | 標題是否粗體 |

色碼不加 `#`。字型要選使用者電腦（與觀眾電腦）有的；跨平台分享用 `Microsoft JhengHei` 最保險。

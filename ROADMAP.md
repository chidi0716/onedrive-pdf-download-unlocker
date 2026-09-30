# Roadmap / 待辦清單

整理自 1.2.0 發佈後的兩輪審查（文件／介面、程式碼／運行邏輯），尚未實作。

標記：量 S／M／L = 小／中／大；🧪 = 改完需實機測試；📊 = 動手前需先取得數據。

> **開工前先取得的數據**：一次匯出的 console `timingsMs`（nav／render／capture），以及匯出期間在檢視器 frame 的 console 執行 `document.fonts.status` 的結果。這兩個決定 #5、#6 怎麼改。

---

## 階段 1：1.2.1 修正版（錯誤、文件、效能）

| # | 項目 | 位置 | 量 | |
|---|---|---|---|---|
| 1 | README／CHANGELOG 宣稱「頁內按鈕已驗證」「Chrome／Edge 不會閃」未經實測，改成「預期」或實測後再寫 | README×2、CHANGELOG | S | |
| 2 | 頁內按鈕啟動時焦點落在頂層頁面，方向鍵可能送錯 frame → 點擊後對檢視器 iframe 呼叫 `.focus()` | content.js | S | 🧪 |
| 3 | 頁內按鈕可能被截進 PDF、且蓋住右下角縮放列 → 匯出中隱藏、位置上移、加 ✕ | content.js | S | 🧪 |
| 4 | `dismissNotification` 對任何 `[role=alertdialog]`／`[id^=BaseCallout]` 按第一個按鈕，可能誤觸其他對話框 → 比對通知文字，或改用 Escape／點空白處 | slides.js | S | 🧪 |
| 5 | 每張可能固定多等 200ms 字型（`document.fonts.status` 永不為 loaded 時）→ 開始時等一次即可 | slides.js（waitSlideReady） | S | 🧪📊 |
| 6 | 開始前固定卡 3～5 秒：`waitForStable` 等「全部網路」靜止（自動存檔讓它幾乎必等滿 3s）、slides.js 400ms 與 background.js 700ms 重複等提示列、`waitLayoutQuiet` 至少 1.5s → 精簡 | slides.js（exportSlides 開頭）、background.js（beginSlideExport） | S | 🧪 |
| 7 | 中途失敗時整份作廢 → 失敗也存「部分 PDF」並標示停在第幾張 | slides.js | M | |
| 8 | 使用者按掉偵錯提示列時只顯示 `export not started` → 改為「已取消匯出」並正常收尾 | slides.js、background.js | S | |
| 9 | `slidesExportTitle` 仍寫「2 倍解析度」（已改為三檔畫質） | i18n.js | S | |
| 10 | 刪除沒人用的 `slidesNeedPermission` 字串 | i18n.js | S | |
| 11 | debugger 提示寫死「Chrome」→ 改「瀏覽器」（Edge 使用者） | i18n.js | S | |
| 12 | PPT 頁面開 popup 時訊息矛盾（偵測到投影片 + 請重新整理 PDF 預覽）；使用說明只有 PDF 流程 → 投影片頁面隱藏空狀態／重新掃描，說明依頁面切換 | popup.js、i18n.js | M | |
| 13 | popup 的 `startSlides` 不管成功失敗都顯示「已開始」 | popup.js | S | |
| 14 | `looksLikePpt` 的 `powerpoint` 關鍵字太寬，檔名含 PowerPoint 的 PDF 頁會被誤判 | content.js | S | |
| 15 | README 安裝說明的檔名與 Release 附件（帶版本號）不一致 | README×2 | S | |
| 16 | PRIVACY.md「最後更新」仍為 2026-06-26（內容已加 debugger 說明） | PRIVACY.md | S | |
| 17 | CHANGELOG 開發紀錄中舊的「已知限制：右上角白色小方塊原因不明」— 已確認是通知並已處理；兩個「已知限制」段落並存易混淆 | CHANGELOG | S | |

## 階段 2：體驗與穩健性

| # | 項目 | 位置 | 量 | |
|---|---|---|---|---|
| 18 | 檢視器偵測改事件驅動：slides.js ready 時主動通知 background → 轉給頂層 frame（frameId 0）。現在是每 0.8 秒廣播、只輪詢 15 秒，超過 15 秒才載好的簡報不會出現按鈕 | slides.js、background.js、content.js | M | 🧪 |
| 19 | 深色模式按鈕對比不足：白字配 #4f8cff 約 3.2:1、hover #6ea0ff 約 2.6:1（AA 需 4.5:1）→ 填色按鈕改用較深的藍 | popup.html | S | |
| 20 | 「⭐ 可能是目標」與「📦 大型檔案」標籤顏色相同，失去區辨 | popup.html | S | |
| 21 | 從 popup 啟動的匯出，頁內按鈕不會變成「匯出中…」（只在自己被點時監看 busy） | content.js | S | |
| 22 | slides.js 語言只在載入時讀一次，popup 切換語言後狀態列不跟著變 | slides.js | S | |
| 23 | `buildPdf` 把所有 JPEG 再複製進一個大陣列，峰值記憶體加倍 → 改用 `new Blob([...chunks])` | slidepdf.js、slides.js | S | |
| 24 | `captureSlide` 在 fromSurface:false 逾時後，原請求仍在跑又發表面截圖，可能兩個同時進行 | background.js | S | |

## 階段 3：重構（使用者看不到差別）

| # | 項目 | 量 |
|---|---|---|
| 25 | 重複程式抽成共用 `util.js`（載入方式同 i18n.js）：`sanitizeFilename` ×3（且略有差異）、base64 解碼 ×3、大檔分塊傳輸 ×2（content.js／popup.js 完全相同） | M |
| 26 | 把「微軟 DOM 知識」集中到 `slides-viewer.js`：`viewPanel`、`thumbnails`、`slidePosition`、`currentSlideBox`、`wrapperHasContent`、`slideTextFrom`、hideHints 選擇器、`dismissNotification`。微軟改版時只修這一個檔 | M |
| 27 | 清死碼／過時註解：`waitForStable`、`lastResourceAt`（#6 之後）、`endSlideExport` 裡的隔離清理與 `#__odpdf_hide_top`、開頭「2x 擷取」、`slideInput` 的「PageDown」、background.js「debugger 為選用權限」 | S |
| 28 | 散落的等待時間（200、250、60、5000…）集中成一個 `TIMING` 設定 | S |
| 29 | Firefox manifest 改由打包腳本從 `manifest.json` 自動產生（拿掉 debugger、換 background、去掉 slides content script），避免版本不同步 | S |

## 階段 4：工程流程

| # | 項目 | 量 |
|---|---|---|
| 30 | 單元測試：`buildPdf`、`jpegSize`、`sanitizeFilename` 等純函式；自動驗證產出的 PDF 有效（qpdf／pdf.js） | M |
| 31 | 打包腳本改跨平台（node），`build.ps1` 在 macOS 無法執行；zip 檔名自動帶 manifest 版本號 | S |
| 32 | GitHub Actions：推 tag 自動語法檢查 → 測試 → 打包兩個 zip → 附到 Release | M |

## 階段 5：1.3.0 新功能

| # | 項目 | 量 |
|---|---|---|
| 33 | 可選取／可複製的隱形文字層：採 OCR 可搜尋 PDF 的「無字形字型」做法（內嵌一個所有字形皆空白的極小字型 + ToUnicode 對照表，文字以 render mode 3 隱形疊在圖上），位置由 DOM `Range.getClientRects()` 換算成投影片座標。不需內嵌中文字型即可正確複製／搜尋中文，檔案幾乎不增加 | L |
| 34 | dsf 高解析度輸出：以 `Emulation.setDeviceMetricsOverride` 提高 deviceScaleFactor，突破螢幕原生解析度（也可取消目前「最高」畫質被原生解析度限制的問題） | M |

## 之後

| # | 項目 |
|---|---|
| 35 | manifest `description` 補上投影片匯出（上架商店時需要） |

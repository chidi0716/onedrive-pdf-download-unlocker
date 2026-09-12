# Changelog

本專案版本紀錄遵循 [Keep a Changelog](https://keepachangelog.com/) 格式，版本號採用 [Semantic Versioning](https://semver.org/)。

## [1.1.0] - 2026-09-13

> 測試版（pre-release）。大型檔案下載與 Firefox 支援尚待實機測試，歡迎回報結果。

### 新增

- **Firefox 支援**：新增 `manifest.firefox.json`（將 Chrome 專用的 `background.service_worker` 換成 Firefox 需要的 `background.scripts`，並加上 `browser_specific_settings.gecko` 設定）。搭配 `build.ps1` 打包腳本，一鍵產出 Chrome 與 Firefox 兩個 zip。

### 修正

- **大型檔案下載失敗（issue #1：「Message exceeded maximum allowed size of 64MiB」）**：舊版把整份檔案 base64 後塞進單一則 `chrome.runtime` 訊息，超過訊息大小上限（約 64 MiB，base64 又會膨脹約 33%）就失敗。改為將檔案位元組暫存後以分塊（每塊 8 MB）方式傳輸，任何單則訊息都遠低於上限，不再受此限制。小型檔案維持原本的單則快速路徑。
- **Firefox 相容性（issue #2）**：`webRequest` 監聽的 `extraHeaders` 旗標為 Chrome 專用、在 Firefox 會拋錯，改為 try/catch，Firefox 自動退回不帶該旗標。

## [1.0.0] - 2026-06-26

### 新增

- 偵測 OneDrive / SharePoint 預覽頁面中真正帶有檔案內容的網路請求，並浮出下載按鈕，一鍵以原始檔名下載。
- 多語言介面（英文 / 中文，popup 即時切換）。
- 深色/淺色主題自適應。
- 鍵盤可及性（Tab / Enter / 空白鍵操作下載按鈕與關閉鈕）。
- 同分頁切換檔案時自動重置候選檔案狀態。
- 找不到候選檔案時顯示「未偵測到 PDF」的停用狀態，避免看起來像沒反應。

### 修正

- 修正 `manifest.json` 的 `name`、`description`、`action.default_title` 與實際功能/README 不一致的問題（舊文字描述的是「偵測最大 PDF-like 回應」的舊版偵測邏輯，且 `default_title` 只有中文）。

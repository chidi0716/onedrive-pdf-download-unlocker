# Changelog

本專案版本紀錄遵循 [Keep a Changelog](https://keepachangelog.com/) 格式，版本號採用 [Semantic Versioning](https://semver.org/)。

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

# 隱私權政策 / Privacy Policy

**OneDrive / SharePoint PDF Download Unlocker**

最後更新 / Last updated: 2026-06-26

---

## 中文版

### 這個擴充功能做什麼

這個擴充功能在你開啟 OneDrive 或 SharePoint 的 PDF 預覽分頁時，自動偵測該分頁載入的檔案內容請求，並讓你用一個按鈕把檔案另存到本機的「下載」資料夾，取代手動用瀏覽器開發工具或 PowerShell 抓取網址的流程。

### 我們不收集、不上傳任何資料

這個擴充功能的所有運作都只發生在你自己的瀏覽器裡：

- 偵測檔案請求、抓取檔案內容、重組檔案、觸發下載，全部都在你的電腦本機完成。
- 我們**不會**把你瀏覽的網址、檔案內容、檔案名稱、帳號資訊或任何其他資料傳送到任何外部伺服器、第三方服務，或開發者自己的伺服器（因為根本沒有這樣的伺服器存在）。
- 沒有任何分析、追蹤、廣告或遠端記錄功能。

### 本機儲存的資料

擴充功能會用瀏覽器內建的 `chrome.storage.local`（純粹存在你自己的電腦上，不會同步到雲端、不會離開你的裝置）儲存一項設定：

- 你選擇的顯示語言（英文或中文）。

除此之外不會儲存任何其他內容；分頁關閉後，該分頁偵測到的檔案請求資訊也會自動清除，不會留存。

### 使用到的權限與用途

| 權限 | 用途 |
|---|---|
| `webRequest` / `webNavigation` | 偵測 OneDrive / SharePoint 頁面載入的檔案內容請求，找出真正的 PDF/文件資料 |
| `downloads` | 把抓到的檔案內容存成本機檔案 |
| `activeTab` / `scripting` | 在目前分頁裡讀取檔案標題、觸發瀏覽器原生的下載動作 |
| `storage` | 只用來記住你選的顯示語言 |
| `host_permissions`（限定 sharepoint.com、sharepointonline.com、onedrive.live.com、officeapps.live.com、1drv.ms、svc.ms、mcas.ms、login.microsoftonline.com、login.windows.net、spoprod-a.akamaihd.net） | 只在這些微軟 OneDrive / SharePoint 相關網域上運作（`svc.ms` 是實際傳送檔案內容的後端服務網域；`mcas.ms` 是部分企業/學校啟用的安全代理網域；其餘是登入與靜態資源網域），不會影響或讀取其他網站 |

---

## English

### What this extension does

This extension watches the network requests made by OneDrive / SharePoint PDF preview pages, detects the request that actually carries the file's content, and lets you save it to your local Downloads folder with one click — replacing a manual DevTools / PowerShell workflow.

### We do not collect or transmit any data

Everything this extension does happens entirely inside your own browser:

- Detecting file requests, fetching file bytes, rebuilding the file, and triggering the download all happen locally on your machine.
- We do **not** send the URLs you visit, file contents, filenames, account information, or any other data to any external server, third-party service, or a server operated by the developer (no such server exists).
- There is no analytics, tracking, advertising, or remote logging of any kind.

### Data stored locally

The extension uses the browser's built-in `chrome.storage.local` (which lives only on your own device and is never synced or transmitted) to remember exactly one setting:

- Your chosen display language (English or Chinese).

Nothing else is stored. Detected file-request information for a tab is automatically cleared once that tab is closed.

### Permissions used and why

| Permission | Purpose |
|---|---|
| `webRequest` / `webNavigation` | Detect the network request on OneDrive/SharePoint pages that carries the actual PDF/document content |
| `downloads` | Save the fetched file content as a local file |
| `activeTab` / `scripting` | Read the document title in the current tab and trigger the browser's native download action |
| `storage` | Remember only your chosen display language |
| `host_permissions` (limited to sharepoint.com, sharepointonline.com, onedrive.live.com, officeapps.live.com, 1drv.ms, svc.ms, mcas.ms, login.microsoftonline.com, login.windows.net, spoprod-a.akamaihd.net) | The extension only operates on these Microsoft OneDrive/SharePoint domains (`svc.ms` actually serves file content; `mcas.ms` is a security proxy domain some organizations enable; the rest are sign-in and static-asset domains) and does not access or affect any other website |

# 客戶業績達成率儀表板

純靜態網頁，用來查看**每位業務底下每個客戶**的業績與達成率。可直接架在 GitHub Pages 上，不需要任何後端或資料庫。

這是 `sales-tracker`（業務 × 產品系列）的姊妹站，資料來源同一份業績總表 Excel，只是維度改成「業務 × 客戶」。

## 功能

- **業務篩選**：選一位業務，即看到他名下所有客戶的目標、累計業績、達成率、進度狀態
- **摘要卡片**：年度目標、累計業績、年度達成率、進度狀態（超前／接近／落後）
- **客戶業績排行**：累計業績 Top 15（顏色代表進度狀態）
- **各通路達成率**：醫學中心／區域醫院／地區醫院／診所… 達成率比較
- **客戶一覽表**：可依產品系列、通路篩選，搜尋客戶或體系名稱，點欄位標題排序
- 「全部業務」時為全公司彙總；選定單一系列時完整顯示該系列（含無目標客戶）

> 達成率口徑與 `sales-tracker` 一致：未選系列時只計入「有目標」的產品系列，
> 讓分子（業績）與分母（目標）口徑相同。

## 如何更新資料

資料每月一次，從業績總表 Excel 重新產生。最簡單的方式：

**把 Excel 檔「拖曳」到 `更新客戶業績.bat` 圖示上放開**（不要雙擊 bat）。
它會自動讀取 Excel → 重新產生 `data.js` →（若已設定好 GitHub）自動上傳。

手動方式：

```
python convert_xlsx.py "G:\My Drive\業績\2026\新業代業績總表(含target)_20260528 (忠季).xlsx"
git add data.js
git commit -m "更新客戶業績資料"
git push
```

- 年度與資料截止日會自動從**檔名裡的 8 位數日期**抓取（例：`..._20260528...` → 2026 / 2026-05-28）
- 範圍：處別 Div1 / Div2 / Div3，不含直營與經銷
- 來源分頁固定為 Excel 的「業績 DATA」，客戶欄為「客戶名稱」

## 部署到 GitHub Pages

1. 在 GitHub 建立一個新 repository（例如 `customer-tracker`）
2. 把這個資料夾的檔案 push 上去：

   ```
   git init
   git add .
   git commit -m "init"
   git remote add origin https://github.com/<你的帳號>/customer-tracker.git
   git branch -M main
   git push -u origin main
   ```

3. 到 repo 的 **Settings → Pages**，Source 選 **Deploy from a branch**，
   Branch 選 `main`、資料夾選 `/ (root)`，按 Save
4. 約一分鐘後，網站會出現在 `https://<你的帳號>.github.io/customer-tracker/`

> **隱私提醒**：GitHub Pages 是公開的，任何知道網址的人都看得到客戶與業績資料。
> 客戶層級資料更敏感，**強烈建議**把 repository 設為 Private、或只在內部分享網址。

## 本機預覽

直接用瀏覽器打開 `index.html` 即可（資料用 `<script>` 載入，不受 file:// 限制）。

## 技術說明

- 純 HTML / CSS / JavaScript，無需建置工具
- 圖表使用 [Chart.js](https://www.chartjs.org/)（CDN 載入）
- 響應式設計，手機也能瀏覽
- 目前資料規模：34 位業務、約 920 個客戶

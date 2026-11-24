# 地圖簡報 PPT 自動產生

GitHub Actions（`.github/workflows/deploy.yml`）在每次網站更新時執行：

1. `make_pptx.py` 讀 `data/years.json` 列出的每個年度資料 `data/<年>.json`。
2. `buildyear.py` 以 `spec_base.json.gz`（原互動式地圖簡報的完整結構，含 Morph 轉場、滑過亮起、圖片）為底，
   填入該年度統計：各區圖表、總場次、亮起顏色（有資料橘色／無資料灰白）、局限空間類型、是否偏遠（115 年起）、
   缺失項目統計、缺失五項以上公司表；圓餅圖標籤由 `labelfix.py` 自動排開避免重疊。
3. `rebuild.py` 組成 PPTX。
4. `verify_pptx.py` 直接讀 PPTX 的 XML，逐項比對每張圖表（含內嵌 Excel）、總場次、表格、亮起顏色、換頁連結
   與網站資料；有任何不符就不提供該年度下載。
5. 成功的年度輸出到網站的 `pptx/<年>.pptx`，清單在 `pptx/index.json`，年度選擇頁據此顯示「下載 PPT」。

若資料出現原簡報沒有的新分類，該年度會產生失敗（記錄在 `pptx/index.json` 的 `failed`），網站其他部分照常更新。

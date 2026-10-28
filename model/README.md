# Carvana（CVNA）個股模型

命題＋營收線＋收支平衡＋信用風險。Excel 是計算事實的來源（全部活公式），HTML 只負責呈現。

## 成品
- `dist/20261029_Carvana個股模型_v0.2.xlsx`：工作表 README → In → Q → Thesis → PL → CF → Credit → Val → Recon → Sens → Sources
- `dist/20261029_Carvana個股模型_v0.2.html`：單頁自含展示層（每個數字標註來源儲存格）

## 建置
```
python3 model/scripts/sens.py --jobs 6                      # 敏感度（改輸入或公式後重跑）
CVNA_ROWMAP=model/data/rowmap_v02.json python3 model/build_carvana.py /tmp/raw.xlsx
soffice --headless --calc --convert-to xlsx --outdir /tmp/out /tmp/raw.xlsx
cp /tmp/out/raw.xlsx model/dist/<日期>_Carvana個股模型_v<版本>.xlsx
python3 model/scripts/make_html_json.py                     # 產生 model/html.json
```
HTML 由 `YinchenChang/company-models` 的共用產生器 `tools/company_html/build_company_html.py` 讀 `html.json` 與 `dist/*.xlsx` 產生（把本資料夾放到 company-models 根目錄下即可執行，並以 `check_company_html.py` 檢核）。

## 資料
- `data/carvana_facts_20261028.json`：事實總帳（SEC 8-K 股東信、10-Q、10-K、XBRL、ABS 424B5）
- `data/sens_v02.json`：敏感度結果；`data/rowmap_v02.json`：列號對照

架構比照 company-models 的 Oracle v0.4（`docs/plan/個股模型原則_命題加營收分部.md`）與 CoreWeave 收支模型的收支精神。

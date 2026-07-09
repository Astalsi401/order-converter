# 專案說明

這是一支單檔腳本工具：把各電商平台（yahoo購物中心、shopee、shopline、rakuten、酷澎）匯出的訂單報表，換算成公司內部進貨/利潤報表格式，供匯入其他系統使用。核心邏輯都在 `orderConverter.py`，`setting.py` 是設定用的 tkinter GUI（手續費率、來源檔密碼）。

**計算規則的唯一真實來源是 `readme.md`**。修改 `orderConverter.py` 的計算邏輯時，務必同步更新 `readme.md` 對應平台的「計算流程」章節；反之若被要求依 `readme.md` 實作，也要先確認來源檔案實際欄位是否真的存在該欄位名稱（`readme.md` 常常是從別的平台章節複製修改，欄位名稱可能是筆誤或沿用舊名，須以 `待轉檔/*.xlsx` 實際欄位為準，有疑問就用 `AskUserQuestion` 問使用者，不要自己猜）。

# 目錄結構與資料敏感性

- `待轉檔/`、`import/`、`設定/`、`examples/`、`tmp.py`、`*.tsv` 都在 `.gitignore` 中，**只有 `orderConverter.py`、`readme.md`、`setting.py`、`install-packages/`、`.gitignore` 會被 commit**。
- `待轉檔/*.xlsx` 是真實訂單（含姓名、電話、地址），`設定/settings.pkl` 含解密密碼、`設定/商品總表.xlsx` 是進貨價/廠商資料 —— 這些都不能外流或貼到訊息以外的地方。`examples/*.xlsx` 才是可以自由查看/引用的匿名化樣本。
- 想確認實際欄位名稱、資料格式時，直接用 Python/pandas 讀 `待轉檔/*.xlsx`（或 `examples/*.xlsx`）比對，不要憑空假設。

# 執行方式

- 轉檔：`python orderConverter.py`（讀 `待轉檔/`，寫到 `import/`，log 寫在 `設定/run.log`）。
- 修改手續費率/密碼：`python setting.py`（寫入 `設定/settings.pkl`，`orderConverter.py` 執行時會讀取）。
- 沒有自動化測試。改完程式後要跑一次 `python orderConverter.py`，並抽幾筆訂單手算核對（商品總金額、成交手續費、金流費用、運費、利潤），確認跟 `readme.md` 公式一致。

# `orderConverter.py` 架構

`Converter.run()` 固定跑這個順序，新增平台或改邏輯前先搞懂每一步在幹什麼：

1. `concat_fr()` — 讀來源檔、依 `OutputColumns.rename` 把原始欄位改成統一命名（`self.oc.xxx`）。
2. `preprocess()` — 建付款代號、`count`（同訂單內第幾件商品，`0`=第一件）、日期格式、郵遞區號取前3碼、電話補`****`、算出初始「商品總金額」。**平台特有的差異用 `if self.oc.fr == ColumnType().xxx:` 包起來，共用邏輯不要動。**
3. `cols_basic_price()` — 算成交手續費、金流費用；商品總金額若需要依商品權重分攤到每一列（多品項訂單），也在這裡做（`groupby(code)[price].transform('first') / groupby(code)[tmp].sum() * tmp`）。
4. `ship_calculate()` — 算「運費(箱子+包材+運費)」內部成本欄；shopline/rakuten/coupang 這幾個平台若來源檔有實際「運費」欄位，會把它拆成獨立一列（`產品編號`=`888888888`）代表運費品項。
5. `product_detail()` — 依「產品編號」合併 `設定/商品總表.xlsx`（進貨價、廠商、預設倉庫、負責PM），算入庫出貨撿貨費、進貨小計。
6. `cost_calculate()` — 算「訂單金額」；依「預設倉庫」是否為`原廠出貨`/`公司倉`把撿貨費、訂單處理費、運費歸零；非第一件商品的訂單金額/運費/訂單處理費/隱碼服務費/折扣都清 0（`self.order_costs`，用 `getattr`/欄位存在檢查過濾掉不適用該平台的欄位，避免 `AttributeError`/`KeyError`）。
7. `profit_calculate()` — 用該訂單第一列的訂單金額算利潤、利潤百分比，非第一件商品清 0。
8. `to_excel()` — 補齊 `fin_cols`、四捨五入、輸出。

`OutputColumns` 用 `if/elif self.fr == ColumnType().xxx` 分平台設定：`fee`（成交手續費欄位名稱，各平台不同）、`profit_cols`（利潤公式要扣掉哪些欄位）、`fin_cols`（匯出欄位順序）、`rename`（原始欄位→統一欄位名稱對照）；若該平台有拆運費列（見上）還要設 `ship_cols`（拆列後只保留哪些欄位，其餘變 NaN）。**不是每個平台都有 `discount`/`service_fee`/`ship_cols` 等屬性 —— 存取前用 `hasattr()` 或欄位是否存在於 `self.df.columns` 檢查，酷澎就是因為沒有這些屬性/欄位而在早期版本炸掉過。**

`multi_condition()` 是共用的「依欄位值對照表決定另一欄位值」機制（例如依付款方式決定付款代號/金流費用），靠 `set(df).issubset(set(self.df.columns))` 判斷該平台有沒有對照所需的欄位（例如沒有「付款方式」欄位的平台就不會套用）；如果平台缺欄位又想套固定值（酷澎沒有付款方式/配送方式欄位，固定當信用卡、固定運費140），就直接在對應方法裡用 `if self.oc.fr == ColumnType().xxx:` 寫死，不要硬塞假資料進 `multi_condition` 的對照表。

# 新增平台的檢查清單

1. 在 `SourceFiles`、`ColumnType` 加平台代號。
2. 在 `OutputColumns.__init__` 加一個 `elif` 分支：`fee`、`profit_cols`、`fin_cols`、`rename`（若有拆運費列要加 `ship_cols`）。
3. 用 pandas 讀該平台實際的 `待轉檔/*.xlsx`（或 `examples/*.xlsx`）確認欄位名稱、日期格式、金額欄位是否為「訂單重複值」還是「逐商品加總值」——這會決定 `preprocess()`/`cost_calculate()` 該用哪一種訂單金額計算方式（單列直接讀值 vs `groupby().transform(sum)`），選錯會重複計算或漏算。
4. 檢查來源檔有沒有「付款方式」「配送方式」欄位可以走既有的 `multi_condition` 對照表；沒有的話要跟使用者確認固定值怎麼定，不要自己假設。
5. 在 `main()` 建立對應 `Converter(...)`，加進迴圈清單。
6. 更新 `readme.md` 對應章節，並用真實/範例資料跑一次、手算核對利潤公式。

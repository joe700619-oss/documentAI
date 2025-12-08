# 文件產生系統 - 更新說明

## 檔案結構

```
documentAI/
├── createdocument.py          # 主程式 (已更新)
├── update_charter.py          # 儲存格更新模組 (新建)
├── test_find_and_replace.py   # 測試檔案 (保留)
├── data.json                  # 資料來源
├── documents_required_list.json  # 文件設定 (含 type 欄位)
├── templates/                 # 範本資料夾
├── history/                   # 歷史文件資料夾
└── output/                    # 輸出資料夾
```

---

## 主要修改內容

### 1. 新建 `update_charter.py` 模組

這個模組從 `test_find_and_replace.py` 的功能提煉而來，提供可重用的函數：

**主要功能：**
- `process_charter_updates(document, data_content, verbose=True)` - 處理所有儲存格更新
- `find_and_replace_by_config(document, table_config, verbose=True)` - 根據設定替換儲存格
- `format_replacement_content(table_config)` - 格式化替換內容
- `get_target_tables(data_content)` - 取得需要更新的表格設定

**支援的資料格式：**
- 營業項目 (business_items 含 code)
- 章程修訂日期 (business_items 含 edition_date)
- 股東名冊 (shareholders)

---

### 2. 更新 `createdocument.py`

**Import 設定：**
```python
import update_charter  # 儲存格內容替換模組
```

**歷史模式 (type="H") 處理流程：**
```
1. 從 history 資料夾找到符合的檔案 (比對底線前的文件名稱)
   例如: "委任書_測試公司.docx" → 比對 "委任書"

2. 步驟 1: 執行文字替換
   - 根據 data.json 中的 replacements 欄位進行文字替換

3. 步驟 2: 執行儲存格更新
   - 調用 update_charter.process_charter_updates()
   - 根據 data.json 中的 tableData 設定
   - 找到符合 target_start_text 的儲存格並替換內容

4. 輸出: 文件名稱_公司名稱.docx
   例如: "公司章程_測試有限公司.docx"
```

**輸出訊息範例：**
```
[H-處理] 公司章程 (來源: 公司章程_測試公司.docx)
  步驟 1: 執行文字替換...
  步驟 2: 檢查儲存格更新...
  - 已更新 3 個儲存格
[H-成功] 已產生: 公司章程_測試有限公司.docx
```

---

## 使用方式

### 執行主程式
```bash
python createdocument.py
```

### 單獨測試 update_charter 模組
```bash
python update_charter.py
```

---

## 範例：documents_required_list.json

```json
{
    "documents_required_list": [
        {
            "Registration_Type": "遷址(不同縣市)",
            "org_type": "股份有限公司",
            "Required_Documents": [
                {"name": "申請書", "type": "T"},
                {"name": "委任書", "type": "T"},
                {"name": "公司章程", "type": "H"},
                {"name": "股東會議事錄", "type": "T"},
                {"name": "董事會議事錄", "type": "T"},
                {"name": "變更登記表", "type": "T"}
            ]
        }
    ]
}
```

**說明：**
- `type: "T"` → 使用 templates 資料夾中的範本
- `type: "H"` → 使用 history 資料夾中的歷史文件 + 文字替換 + 儲存格更新

---

## 優點

1. **模組化設計**：`update_charter.py` 獨立於主程式，易於維護和測試
2. **可重用性**：其他腳本也可以 import `update_charter` 使用
3. **清晰的處理流程**：歷史模式包含兩個明確的步驟
4. **詳細的輸出訊息**：每個步驟都有明確的狀態回報

---

## 維護建議

- 如需新增其他儲存格替換格式，只需修改 `update_charter.py` 中的 `format_replacement_content()` 函數
- `createdocument.py` 不需要修改，只要匯入模組即可使用最新功能
- 建議保留 `test_find_and_replace.py` 作為獨立測試工具

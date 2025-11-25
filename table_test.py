import json
import os
from docx import Document
from docx.shared import Cm


# --- 工具函數：cm 轉 twips ---
def cm_to_twips(cm):
    """將公分轉換為 Word 的內部單位 twips (二十份點)"""
    return int(cm * 567)

# --- 工具函數：設定整張表格欄寬（真正生效） ---
def set_col_widths(table, widths_cm):
    """
    通過操作底層 XML，強制設定表格所有欄位的寬度。
    """
    # 禁止自動調整欄寬
    table.allow_autofit = False
    
    tbl = table._tbl
    tblGrid = tbl.tblGrid

    # 確保 tblGrid 元素存在
    if len(tblGrid.gridCol_lst) < len(widths_cm):
        raise ValueError("Column count mismatch in underlying XML structure.")

    for i, cm_val in enumerate(widths_cm):
        twips = cm_to_twips(cm_val)
        # 設置 gridCol 元素的寬度
        gridCol = tblGrid.gridCol_lst[i]
        gridCol.w = twips
        
        # 設定現有列的儲存格寬度 (包含標題列)
        for row in table.rows:
            row.cells[i].width = Cm(cm_val)

# Embedded data
data = {
    "document_title": "董事、監察人名單",
    "directors": [
        {
            "id": "001",
            "position": "董事長",
            "name": "李耿佑",
            "id_number": "F128873285",
            "shares": 20000,
            "address": "(220) 新北市板橋區湳興里 4 鄰南雅西路二段 7 巷 18 之 4 號"
        },
        {
            "id": "002",
            "position": "監察人",
            "name": "李耿甫",
            "id_number": "A123456789",
            "shares": 0,
            "address": "台北市信義區忠孝東路一段1號"
        }
    ]
}

# 請確保範本文件路徑正確且樣式存在
document = Document(r"C:\Users\joe70\PythonProject\documentAI\templates\設立登記表.docx")

# 標題
document.add_paragraph(data.get("document_title", ""), style="Normal")

# --------------------------------------
# 建立表格（5 欄）
# --------------------------------------
table = document.add_table(rows=1, cols=5)
table.style = "director_table"

# 依照您提供的欄寬設定
column_widths = [0.75, 2.01, 5.61, 6.03, 3.49]
set_col_widths(table, column_widths)

# 填充標頭行
hdr = table.rows[0].cells
hdr[0].text = "編號"
hdr[1].text = "職稱"
hdr[2].text = "姓名"
hdr[3].text = "身分證字號"
hdr[4].text = "持股數"


# --------------------------------------
# 填入每位董事資料
# --------------------------------------
for d in data["directors"]:

    # 1. 第一列：基本資料
    row = table.add_row().cells
    
    # Explicitly set width for new row cells
    for i, w in enumerate(column_widths):
        row[i].width = Cm(w)
        
    row[0].text = d["id"]
    row[1].text = d["position"]
    row[2].text = d["name"]
    row[3].text = d["id_number"]
    row[4].text = str(d["shares"])

    # 2. 第二列：地址
    addr_row = table.add_row().cells
    
    # Explicitly set width for new row cells
    for i, w in enumerate(column_widths):
        addr_row[i].width = Cm(w)
        
    # 2a. 垂直合併編號欄
    row[0].merge(addr_row[0]) 
    
    # 2b. 設置地址內容
    addr_row[1].text = d["address"]
    
    # 2c. 水平合併
    # Merge addr_row[1] to addr_row[4]
    addr_row[1].merge(addr_row[4])


# 輸出結果
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, "output.docx")
document.save(output_path)
print(f"文件已產生：{output_path}")
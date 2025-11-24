import json
import os
from docx import Document

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

document = Document("設立登記表.docx")

# 標題
document.add_heading(data.get("document_title", ""), level=1)

# 建立表格（6 欄）
table = document.add_table(rows=1, cols=5)
table.style = "director_table"

hdr = table.rows[0].cells
hdr[0].text = "編號"
hdr[1].text = "職稱"
hdr[2].text = "姓名"
hdr[3].text = "身分證字號"
hdr[4].text = "持股數"


# 填入資料
for d in data["directors"]:
    # 第一列：基本資料
    row = table.add_row().cells
    row[0].text = d["id"]
    row[1].text = d["position"]
    row[2].text = d["name"]
    row[3].text = d["id_number"]
    row[4].text = str(d["shares"])
   

    # 第二列：地址（合併 6 欄）
    addr_row = table.add_row().cells
    addr_row[0].text = ""
    addr_row[1].text = d["address"]

    # 合併儲存格（左 → 右）
    for i in range(1, 5):
        addr_row[1].merge(addr_row[i])

# 輸出結果
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, "output.docx")
document.save(output_path)
print(f"文件已產生：{output_path}")
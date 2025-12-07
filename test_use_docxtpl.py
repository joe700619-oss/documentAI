"""
測試 docxtpl 處理表格的功能
使用 data.json 中的 directors 資料填入設立登記表範本
"""

import json
from pathlib import Path
from docxtpl import DocxTemplate

def main():
    # 設定路徑
    base_dir = Path(r"c:\Users\joe70\PythonProject\documentAI")
    template_path = base_dir / "templates" / "設立登記表.docx"
    data_path = base_dir / "data.json"
    output_path = base_dir / "output" / "設立登記表_測試.docx"
    
    # 確認檔案存在
    if not template_path.exists():
        print(f"錯誤: 找不到範本檔案 {template_path}")
        return
    
    if not data_path.exists():
        print(f"錯誤: 找不到資料檔案 {data_path}")
        return
    
    # 讀取 data.json
    print(f"讀取資料: {data_path}")
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 準備 context
    # 將 basicInformation 展平到根層級，方便直接存取
    context = {}
    if 'basicInformation' in data:
        context.update(data['basicInformation'])
    
    # 加入 directors 資料 (用於表格迴圈)
    if 'director_table' in data and 'directors' in data['director_table']:
        context['directors'] = data['director_table']['directors']
        print(f"找到 {len(context['directors'])} 筆董事/監察人資料")
    else:
        print("警告: 找不到 directors 資料")
        context['directors'] = []
    
    # 印出 context 內容供除錯
    print("\n=== Context 內容 ===")
    for key, value in context.items():
        if isinstance(value, list):
            print(f"{key}: (共 {len(value)} 筆)")
            for i, item in enumerate(value):
                print(f"  [{i}] {item}")
        else:
            print(f"{key}: {value}")
    print("=" * 30)
    
    # 載入範本並渲染
    print(f"\n載入範本: {template_path}")
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        # 確保輸出目錄存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        doc.save(output_path)
        print(f"\n成功! 已產生: {output_path}")
        
    except Exception as e:
        print(f"\n錯誤: {e}")
        import traceback
        traceback.print_exc()

    # 提供範本語法說明
    print("\n" + "=" * 50)
    print("【docxtpl 表格語法說明】")
    print("=" * 50)
    print("""
在 Word 範本中，若要產生動態表格列，請使用以下語法：

1. **單一變數替換** (用於一般文字):
   {{ companyName }}
   {{ chairmanName }}

2. **表格迴圈** (用於產生多列):
   在表格的資料列中使用:
   
   第一欄: {%tr for d in directors %}{{ d.id }}
   第二欄: {{ d.position }}
   第三欄: {{ d.name }}
   第四欄: {{ d.id_number }}
   第五欄: {{ d.shares }}
   最後一欄: {{ d.address }}{%tr endfor %}

   重要提示：
   - {%tr for d in directors %} 和 {%tr endfor %} 必須放在表格儲存格內
   - {%tr ... %} 標籤用於表格列迴圈
   - 整個 for 迴圈從 {%tr for ... %} 開始，到 {%tr endfor %} 結束
   - 這兩個標籤必須在同一列的不同儲存格中

3. **正確範例 (表格結構)**:
   
   | 編號 | 職稱 | 姓名 | 身分證號 | 持股數 | 地址 |
   |------|------|------|----------|--------|------|
   | {%tr for d in directors %}{{ d.id }} | {{ d.position }} | {{ d.name }} | {{ d.id_number }} | {{ d.shares }} | {{ d.address }}{%tr endfor %} |

   ↑ 這一列會依據 directors 陣列自動產生多列

4. **條件判斷**:
   {% if directors %}
   有董事資料
   {% endif %}
""")

if __name__ == "__main__":
    main()
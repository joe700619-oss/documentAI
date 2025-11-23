import json
import os
from docx import Document

BASE_DIR = r"C:\Users\joe70\PythonProject\documentAI" 
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMPLATE_FILE = os.path.join(TEMPLATE_DIR, "business_Scope_table.docx")

testData={
    "business_Scope_table": [
        {"code": "J101010", "description": "建築物清潔服務業"},
        {"code": "JA03010", "description": "洗衣業"},
        {"code": "JZ99990", "description": "未分類其他服務業"}
    ]
}

def fill_docx_table_template(template_path, data, output_path):
    """
    讀取 Word 模板並填充表格資料。

    :param template_path: Word 模板路徑
    :param data: 資料字典
    :param output_path: 輸出文件路徑
    """
    try:
        # 1. 載入資料
        # data 已經是字典了，不需要再讀取檔案
        # table_data = data["table_data"] 
        # 使用者提供的 testData 直接就是 table_data 的結構，或者我們假設它是
        table_data = data
        
        # 2. 載入 Word 文件
        if not os.path.exists(template_path):
            print(f"❌ 找不到模板文件: {template_path}")
            return

        document = Document(template_path)

        # 3. 遍歷文件中的所有表格
        for table in document.tables:
            # 遍歷表格中的每一列
            for i, row in enumerate(table.rows):
                # 檢查第一個單元格是否包含表格迴圈標記
                first_cell_text = row.cells[0].text.strip()
                
                # 使用正規表達式或簡單字串查找來識別標記
                if first_cell_text.startswith('{#each'):
                    # 提取列表名稱，例如從 "{#each business_Scope_table}{{code}}" 中取出 "business_Scope_table"
                    list_name = first_cell_text.split('}')[0].replace('{#each', '').strip()
                    
                    if list_name in table_data:
                        template_row = table.rows[i]
                        list_to_iterate = table_data[list_name]
                        
                        # 4. 準備替換模板變數的字典
                        # 提取模板行中所有變數的名稱，例如：{{code}} 和 {{description}}
                        variable_placeholders = {}
                        for cell in template_row.cells:
                            # 找出 {{...}} 標記
                            for p in cell.paragraphs:
                                text = p.text
                                start = text.find('{{')
                                end = text.find('}}')
                                if start != -1 and end != -1:
                                    var_name = text[start + 2: end].strip()
                                    variable_placeholders[cell] = var_name # 儲存 {cell: variable_name}
                                    
                        # 5. 複製和填充行
                        new_rows = []
                        for item_data in list_to_iterate:
                            # 複製樣板行
                            new_row = table.add_row()
                            new_rows.append(new_row)
                            
                            # 填充新行中的單元格
                            for col_index, cell in enumerate(template_row.cells):
                                target_cell = new_row.cells[col_index]
                                
                                # 找到當前單元格的原始文本 (包含變數)
                                original_text = cell.text
                                
                                # 遍歷數據並替換
                                replaced_text = original_text
                                for var_key, var_value in item_data.items():
                                    replaced_text = replaced_text.replace(f'{{{{{var_key}}}}}', str(var_value))
                                
                                target_cell.text = replaced_text
                        
                        # 6. 移除原始模板行
                        # 在 docx 庫中，直接刪除行比較麻煩，通常的做法是清空內容並隱藏，但最簡單的方法是把模板行內容設為空
                        for cell in template_row.cells:
                            cell.text = ''
                            for paragraph in cell.paragraphs:
                                p = paragraph._element
                                p.getparent().remove(p)
                        
                        # 也可以使用更高階的函式庫如 docxtpl 來處理更複雜的模板邏輯

        # 7. 儲存結果
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        document.save(output_path)
        print(f"✅ 文件成功生成至: {output_path}")

    except Exception as e:
        print(f"❌ 處理文件時發生錯誤: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    output_file = os.path.join(OUTPUT_DIR, "business_Scope_table_Output.docx")
    fill_docx_table_template(TEMPLATE_FILE, testData, output_file)
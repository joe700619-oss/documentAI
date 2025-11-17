import os 
import json 
import re 
import ast 
from docx import Document 
from docx.shared import Pt # 引入 Pt 單位，確保字體大小能正確複製
from docx.oxml.ns import qn # <-- 新增此行

# -----------------------------------------------------
# 設定路徑 (根據您的環境變數設置，請自行調整)
# -----------------------------------------------------
# 假設您的主程式目錄已設定好 BASE_DIR
BASE_DIR = r"C:\Users\joe70\PythonProject\documentAI" 
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# **注意：使用您目前的模板**
TEMPLATE_FILE = os.path.join(TEMPLATE_DIR, "board_minutes_template.docx")

# -----------------------------------------------------
# AI 回傳的 JSON（使用您提供的測試資料）
# -----------------------------------------------------
# 此資料已被確定為正確的 AI 萃取結果
data = {
    "files": [
        {
            "file_name": "原人股份有限公司_20251119_董監改選及增資.docx",
            "file_type": "董事會議事錄"
        },
        {
            "file_name": "原人股份有限公司_20251119_董監改選及增資.docx",
            "file_type": "委託書"
        },
        {
            "file_name": "原人股份有限公司_20251119_董監改選及增資.docx",
            "file_type": "股東臨時會會議事錄"
        }
    ],
    "summaries": [
        "原人股份有限公司董事會議事錄採取了在114年11月10日上午10時於公司會議室進行的會議，與會成員為徐峻祥和施皓文，徐峻祥為主席。",
        "委託書顯示，原人股份有限公司委託黃勝平會計師辦理增資發行新股及改選董監事的事宜，並賦予相關的必要權限。",
        "114年股東臨時會會議事錄記載了公司的臨時會議狀況，具體內容未詳述。"
    ],
    "merged_change": "原人股份有限公司在114年11月10日上午10時於公司會議室進行的董事會議中，出席成員包含徐峻祥和施皓文， 並由徐峻祥擔任主席。因應董事長須替換的需求，董事長由徐峻祥更換為Peter。",
    "registration_data": {
        "companyName": "原人股份有限公司",
        "registrationNumber": "60299784",
        "zipcode": "403", 
        "address": "臺中市西區臺灣大道二段2號11樓之6",
        "chairperson": "Peter",
        "change_type": "董事長替換",
        "new_directors": [], 
        "removal_directors": ["徐峻祥"],  
        "directors_table":[
            {"name": "Peter", 
            "position": "director",
            "id": "A123456789",
            "address":"台中市西區五權路1-67號11樓之5",
            "shares":"500,000"
            },
            {"name": "Joe", 
            "position": "chairperson",
            "id": "B123456789",
            "address":"台中市西區五權路1-67號11樓之5",
            "shares":"500,000"
            }
        ]
    } 
}

# -----------------------------------------------------
# 終極穩健替換函數 (包含格式保留)
# -----------------------------------------------------

def _replace_in_paragraph(para, mapping):
    """處理單個 Paragraph 物件中的佔位符替換邏輯 (終極穩健替換 + 格式保留 + CJK字體修正)"""
    full_text = para.text
    
    # 內部輔助函數：從 Run XML 中獲取字體設定
    def _get_font_settings(run):
        settings = {}
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                # 獲取西文字體 (w:ascii) 和中文字體 (w:eastAsia)
                settings['ascii'] = rFonts.get(qn('w:ascii'))
                settings['eastAsia'] = rFonts.get(qn('w:eastAsia'))
        return settings

    # 內部輔助函數：將字體設定套用到新的 Run XML
    def _set_font_settings(run, settings):
        if settings:
            rPr = run._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            
            if settings.get('ascii'):
                rFonts.set(qn('w:ascii'), settings['ascii'])
                rFonts.set(qn('w:hAnsi'), settings['ascii']) # 兼容設定
            
            if settings.get('eastAsia'):
                rFonts.set(qn('w:eastAsia'), settings['eastAsia'])


    for placeholder, value in mapping.items():
        if placeholder in full_text:
            # 1. 執行替換
            new_text = full_text.replace(placeholder, str(value))
            
            # --- 核心修正：格式保留和穩定替換 ---
            
            style = None
            is_bold, is_italic, is_underline = None, None, None
            font_name, font_size = None, None
            
            # **新增**：中文字體設定字典
            cjk_font_settings = {}

            # 2. 儲存格式屬性 (使用第一個 Run 作為格式來源)
            if para.runs:
                source_run = para.runs[0] 
                
                style = source_run.style 
                is_bold = source_run.bold
                is_italic = source_run.italic
                is_underline = source_run.underline
                
                # 標準字體複製
                if source_run.font.name:
                    font_name = source_run.font.name
                if source_run.font.size:
                    font_size = source_run.font.size
                    
                # *** 關鍵：複製 CJK 字體設定 ***
                cjk_font_settings = _get_font_settings(source_run)


            # 3. 刪除所有現有 runs (穩定替換的核心)
            while para.runs:
                run_element = para.runs[0]._element
                run_element.getparent().remove(run_element)
        
            # 4. 插入一個新的 run
            new_run = para.add_run(new_text)
            
            # 5. 重新套用儲存的格式
            if style:
                new_run.style = style
            
            if is_bold is not None: new_run.bold = is_bold
            if is_italic is not None: new_run.italic = is_italic
            if is_underline is not None: new_run.underline = is_underline
            
            if font_name: new_run.font.name = font_name
            if font_size: new_run.font.size = font_size

            # *** 關鍵：應用 CJK 字體設定 ***
            _set_font_settings(new_run, cjk_font_settings)

            full_text = new_text
            
def replace_all_placeholders(doc, mapping):
    """遞歸替換文件中的佔位符 (調用穩健替換函數)"""
    # 1. 處理所有段落 (頂層)
    for para in doc.paragraphs:
        _replace_in_paragraph(para, mapping)

    # 2. 處理所有表格內的段落
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_in_paragraph(para, mapping)


# -----------------------------------------------------
# 主程式
# -----------------------------------------------------
def generate_document():
    
    if not os.path.exists(TEMPLATE_FILE):
        print(f"🚨 錯誤：找不到模板文件 {TEMPLATE_FILE}")
        return

    print("讀取模板：", TEMPLATE_FILE)
    doc = Document(TEMPLATE_FILE)

    reg = data["registration_data"]

    # 替換映射：假設您已將範本中的 {{...}} 替換為純英文標籤 (例如：companyName)
    replacements = {
        "companyName": reg.get("companyName", ""),
        "registrationNumber": reg.get("registrationNumber", ""),
        "zipcode": reg.get("zipcode", ""), 
        "address": reg.get("address", ""),
        "chairperson": reg.get("chairperson", ""),
        # 如果範本中有其他標籤，請在此處添加
    }

    # 使用穩健替換函數
    replace_all_placeholders(doc, replacements)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    # 更改輸出檔名以避免與舊的錯誤文件衝突
    output_path = os.path.join(OUTPUT_DIR, "test_output_final_formatted.docx")
    doc.save(output_path)

    print("✔ 成功輸出 →", output_path)


if __name__ == "__main__":
    generate_document()
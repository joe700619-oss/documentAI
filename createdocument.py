"""
文件產生工具
- 模式 1: 使用歷史資料 (history 資料夾) + 文字替換
- 模式 2: 使用範本 (templates 資料夾) + 變數/表格替換
"""
import json
from pathlib import Path
from docx import Document
from docxtpl import DocxTemplate


def replace_text_in_paragraph(paragraph, replacements):
    """替換段落中的文字"""
    for replacement in replacements:
        old_text = replacement.get('old_text', '')
        new_text = replacement.get('new_text', '')
        if old_text and old_text in paragraph.text:
            # 需要處理 runs 以保留格式
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)


def replace_text_in_table(table, replacements):
    """替換表格中的文字"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)


def replace_text_in_document(doc_path, replacements, output_path):
    """使用 python-docx 替換文件中的文字"""
    doc = Document(doc_path)
    
    # 替換段落中的文字
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    
    # 替換表格中的文字
    for table in doc.tables:
        replace_text_in_table(table, replacements)
    
    # 替換頁首頁尾
    for section in doc.sections:
        # 頁首
        for paragraph in section.header.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)
        for table in section.header.tables:
            replace_text_in_table(table, replacements)
        # 頁尾
        for paragraph in section.footer.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)
        for table in section.footer.tables:
            replace_text_in_table(table, replacements)
    
    doc.save(output_path)


def get_required_documents():
    """共用函數: 詢問變更種類並取得必要文件清單"""
    json_path = r"c:\Users\joe70\PythonProject\documentAI\documents_required_list.json"
    
    # 讀取 JSON 檔案
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"找不到檔案: {json_path}")
        return None
    except json.JSONDecodeError:
        print(f"檔案格式錯誤: {json_path}")
        return None

    # 詢問變更種類
    print("\n請輸入本次變更種類 (多種變更請用 ',' 分隔)")
    user_input = input("例如 '遷址(同縣市), 變更負責人': ").strip()

    # 處理輸入：支援全形逗號轉半形，並分割
    user_input = user_input.replace('，', ',')
    input_types = [t.strip() for t in user_input.split(',') if t.strip()]

    if not input_types:
        print("未輸入任何內容")
        return None

    # 建立 mapping (排除 default)
    registration_map = {}
    default_docs = []
    for item in data.get("documents_required_list", []):
        reg_type = item.get("Registration_Type")
        if reg_type == "default":
            default_docs = item.get("Required_Documents", [])
        else:
            registration_map[reg_type] = item.get("Required_Documents", [])

    # 依照取得的種類，取回需要文件(Required_Documents)並移除重複
    all_required_docs = []
    seen_docs = set()
    found_types = []
    not_found_types = []

    for type_name in input_types:
        if type_name in registration_map:
            found_types.append(type_name)
            docs = registration_map[type_name]
            for doc in docs:
                if doc not in seen_docs:
                    seen_docs.add(doc)
                    all_required_docs.append(doc)
        else:
            not_found_types.append(type_name)

    # 如果完全沒有找到任何種類，使用 default
    if not found_types:
        print(f"\n找不到您輸入的所有種類: {', '.join(not_found_types)}")
        all_required_docs = default_docs
        print("\n目前系統支援的種類如下:")
        for key in registration_map.keys():
            print(f"  - {key}")
    else:
        print(f"\n已識別的變更種類: {', '.join(found_types)}")
        if not_found_types:
            print(f"警告: 找不到以下種類: {', '.join(not_found_types)}")

    # 顯示必要文件清單
    print(f"\n合併後的必要文件清單 (已移除重複):")
    print("-" * 40)
    for doc in all_required_docs:
        print(f"[x] {doc}")
    print("-" * 40)

    return {
        'all_required_docs': all_required_docs,
        'found_types': found_types,
        'not_found_types': not_found_types,
        'registration_map': registration_map
    }


def process_history_mode(data_content):
    """模式 1: 處理歷史資料"""
    base_dir = Path(r"c:\Users\joe70\PythonProject\documentAI")
    history_dir = base_dir / "history"
    output_dir = base_dir / "output"
    
    # 1. 詢問變更種類並顯示必要文件清單
    result = get_required_documents()
    if result is None:
        return
    
    # 確保目錄存在
    if not history_dir.exists():
        print(f"\n錯誤: 找不到 history 資料夾: {history_dir}")
        return
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 2. 取得替換規則
    replacements = data_content.get('replacements', [])
    if not replacements:
        print("\n警告: data.json 中沒有找到 replacements 設定")
        return
    
    print(f"\n載入 {len(replacements)} 組替換規則:")
    for r in replacements:
        print(f"  '{r.get('old_text')}' → '{r.get('new_text')}'")
    
    # 3. 找到 history 中所有的 docx 檔案
    docx_files = list(history_dir.glob("*.docx"))
    
    if not docx_files:
        print(f"\n在 history 資料夾中找不到任何 .docx 檔案")
        return
    
    print(f"\n找到 {len(docx_files)} 個歷史檔案:")
    for f in docx_files:
        print(f"  - {f.name}")
    
    # 4. 處理文字替換
    print(f"\n--- 開始處理文字替換 ---")
    success_count = 0
    error_count = 0
    
    for doc_path in docx_files:
        output_path = output_dir / doc_path.name
        try:
            replace_text_in_document(doc_path, replacements, output_path)
            print(f"[成功] {doc_path.name} → {output_path.name}")
            success_count += 1
        except Exception as e:
            print(f"[錯誤] 處理 {doc_path.name} 失敗: {e}")
            error_count += 1
            import traceback
            traceback.print_exc()
    
    print(f"\n--- 處理完成 ---")
    print(f"成功: {success_count} 份")
    print(f"錯誤: {error_count} 份")


def process_template_mode(data_content):
    """模式 2: 處理範本"""
    # 1. 詢問變更種類並取得必要文件清單
    result = get_required_documents()
    if result is None:
        return
    
    all_required_docs = result['all_required_docs']
    found_types = result['found_types']
    
    # 如果沒有找到任何種類（使用了 default），就不產生文件
    if not found_types:
        print("\n由於無法判斷變更種類，不產生文件。")
        return
    
    print("\n開始產生文件...")

    # 準備 context
    context = {}
    
    # 1. 將 basicInformation 展平放入根目錄
    if 'basicInformation' in data_content:
        context.update(data_content['basicInformation'])
        print(f"已載入基本資料: {list(data_content['basicInformation'].keys())}")
    
    # 2. 處理表格資料 - 動態搜尋 tableData 內所有表格
    if 'tableData' in data_content:
        table_data = data_content['tableData']
        print(f"找到 tableData，包含 {len(table_data)} 個表格定義")
        
        for table_name, table_content in table_data.items():
            if isinstance(table_content, dict):
                for key, value in table_content.items():
                    if isinstance(value, list):
                        context[key] = value
                        print(f"  - 已載入 {len(value)} 筆 {key} 資料 (來自 {table_name})")
                    elif key not in ['document_title']:
                        context[key] = value
    
    # 3. 處理其他可能的表格資料（向下相容舊格式）
    for key, value in data_content.items():
        if key in ['basicInformation', 'tableData', 'replacements']:
            continue
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if isinstance(sub_value, list) and sub_key not in context:
                    context[sub_key] = sub_value
                    print(f"已載入 {len(sub_value)} 筆 {sub_key} 資料")
        elif isinstance(value, list) and key not in context:
            context[key] = value
            print(f"已載入 {len(value)} 筆 {key} 資料")

    # 設定目錄
    templates_dir = Path(r"c:\Users\joe70\PythonProject\documentAI\templates")
    output_dir = Path(r"c:\Users\joe70\PythonProject\documentAI\output")
    
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n--- 開始處理 {len(all_required_docs)} 份文件 ---")
    success_count = 0
    skip_count = 0
    error_count = 0

    for doc_name in all_required_docs:
        template_path = templates_dir / f"{doc_name}.docx"
        
        if not template_path.exists():
            print(f"[略過] 找不到範本: {template_path.name}")
            skip_count += 1
            continue
        
        try:
            doc = DocxTemplate(template_path)
            doc.render(context)
            
            output_path = output_dir / f"{doc_name}.docx"
            doc.save(output_path)
            print(f"[成功] 已產生: {output_path.name}")
            success_count += 1
        except Exception as e:
            print(f"[錯誤] 產生 {doc_name} 失敗: {e}")
            error_count += 1
            import traceback
            traceback.print_exc()
    
    print(f"\n--- 處理完成 ---")
    print(f"成功: {success_count} 份")
    print(f"略過 (無範本): {skip_count} 份")
    print(f"錯誤: {error_count} 份")


def main():
    print("=" * 50)
    print("文件產生工具")
    print("=" * 50)
    print("\n請選擇模式:")
    print("  1. 使用歷史資料 (從 history 資料夾讀取，進行文字替換)")
    print("  2. 使用範本 (從 templates 資料夾讀取，使用變數/表格替換)")
    print("")
    
    mode = input("請輸入 1 或 2: ").strip()
    
    # 讀取 data.json
    data_json_path = r"c:\Users\joe70\PythonProject\documentAI\data.json"
    try:
        with open(data_json_path, 'r', encoding='utf-8') as f:
            data_content = json.load(f)
    except Exception as e:
        print(f"讀取 data.json 失敗: {e}")
        return
    
    if mode == "1":
        print("\n>>> 已選擇: 歷史資料模式")
        process_history_mode(data_content)
    elif mode == "2":
        print("\n>>> 已選擇: 範本模式")
        process_template_mode(data_content)
    else:
        print(f"\n無效的選擇: {mode}")
        print("請輸入 1 或 2")


if __name__ == "__main__":
    main()

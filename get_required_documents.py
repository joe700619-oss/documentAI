import json
from typing import List, Dict

def get_required_documents(allowed_types: List[str], documents_required_list_path: str = "documents_required_list.json") -> List[Dict[str, str]]:
    """
    【純 Python 實現】根據 allowed_types 從 documents_required_list.json 取得所有需要的 required_documents
    此函數不使用 LLM，直接從 JSON 檔案中讀取和過濾資料
    
    Args:
        allowed_types: 變更類型列表 (例如: ["遷址(同縣市)", "變更負責人"])
        documents_required_list_path: documents_required_list.json 的路徑
        
    Returns:
        required_documents 列表
        - 如果只有一個 registration_type: 回傳該類型的所有 required_documents
        - 如果有多個 registration_type: 只回傳名稱包含"議事錄"、"同意書"或"公司章程"的文件(去重)
    """
    try:
        # 載入 documents_required_list.json
        with open(documents_required_list_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        documents_list = data.get("documents_required_list", [])
        
        # 收集所有匹配的 required_documents
        all_required_docs = []
        
        for reg_type in allowed_types:
            # 找到對應的 registration_type
            for item in documents_list:
                if item.get("registration_type") == reg_type:
                    required_docs = item.get("required_documents", [])
                    all_required_docs.extend(required_docs)
                    break
        
        # 如果只有一個 registration_type，直接回傳所有文件
        if len(allowed_types) == 1:
            return all_required_docs
        
        # 如果有多個 registration_type，過濾文件
        # 保留包含"議事錄"、"同意書"或"公司章程"的文件
        filtered_docs = []
        for doc in all_required_docs:
            doc_name = doc.get("name", "")
            if "議事錄" in doc_name or "同意書" in doc_name or "公司章程" in doc_name:
                filtered_docs.append(doc)
        
        # 移除重複的文件 (根據 name 去重)
        unique_docs = []
        seen_names = set()
        
        for doc in filtered_docs:
            doc_name = doc.get("name", "")
            if doc_name not in seen_names:
                seen_names.add(doc_name)
                unique_docs.append(doc)
        
        return unique_docs
        
    except FileNotFoundError:
        print(f"Error: File '{documents_required_list_path}' not found.")
        return []
    except json.JSONDecodeError as e:
        print(f"Error: Failed to parse JSON: {e}")
        return []
    except Exception as e:
        print(f"Error: {e}")
        return []


# 測試範例
if __name__ == "__main__":
    # 測試案例1: 只有一個 registration_type
    print("=== 測試案例1: 單一變更類型 ===")
    test_types_single = ["遷址(同縣市)"]
    result_single = get_required_documents(test_types_single)
    print(f"輸入: {test_types_single}")
    print(f"結果 ({len(result_single)} 個文件):")
    for doc in result_single:
        print(f"  - {doc['name']} (type: {doc['type']})")
    
    print("\n" + "="*50 + "\n")
    
    # 測試案例2: 多個 registration_type
    print("=== 測試案例2: 多個變更類型 ===")
    test_types_multiple = ["遷址(不同縣市)", "重新改選董監事", "變更負責人"]
    result_multiple = get_required_documents(test_types_multiple)
    print(f"輸入: {test_types_multiple}")
    print(f"結果 ({len(result_multiple)} 個文件，僅顯示包含'議事錄'、'同意書'或'公司章程'的文件):")
    for doc in result_multiple:
        print(f"  - {doc['name']} (type: {doc['type']})")

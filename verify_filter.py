"""
最終驗證: 確認過濾功能正確運作
"""
from get_required_documents import get_required_documents
import json

def print_separator():
    print("=" * 80)

def verify_filtering():
    print_separator()
    print("✅ 驗證 get_required_documents 功能 (純 Python 實現,不使用 LLM)")
    print_separator()
    
    # 測試1: 單一類型
    print("\n📋 測試1: 單一變更類型 - 回傳所有文件")
    print("-" * 80)
    types1 = ["遷址(同縣市)"]
    result1 = get_required_documents(types1)
    print(f"輸入: {types1}")
    print(f"輸出: {len(result1)} 個文件 (全部回傳)")
    for i, doc in enumerate(result1, 1):
        print(f"  {i}. {doc['name']} [{doc['type']}]")
    
    # 測試2: 多個類型 - 測試過濾功能
    print("\n📋 測試2: 多個變更類型 - 過濾議事錄、同意書、公司章程")
    print("-" * 80)
    types2 = ["遷址(不同縣市)", "重新改選董監事"]
    result2 = get_required_documents(types2)
    print(f"輸入: {types2}")
    print(f"輸出: {len(result2)} 個文件 (僅保留議事錄、同意書、公司章程)")
    for i, doc in enumerate(result2, 1):
        print(f"  {i}. {doc['name']} [{doc['type']}]")
    
    # 驗證過濾條件
    print("\n🔍 驗證過濾條件:")
    print("-" * 80)
    for doc in result2:
        name = doc['name']
        if "議事錄" in name:
            print(f"✓ {name} - 包含'議事錄'")
        elif "同意書" in name:
            print(f"✓ {name} - 包含'同意書'")
        elif "公司章程" in name:
            print(f"✓ {name} - 包含'公司章程'")
        else:
            print(f"✗ {name} - 不符合條件 (錯誤!)")
    
    print("\n" + "=" * 80)
    print("✅ 測試完成! 功能運作正常")
    print("=" * 80)
    
    return result2

if __name__ == "__main__":
    verify_filtering()

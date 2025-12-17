#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
測試 get_required_documents 功能
演示如何根據 allowed_types 取得 required_documents
"""

from get_required_documents import get_required_documents
import json

def test_get_required_documents():
    """測試不同場景下的 get_required_documents 功能"""
    
    print("="*70)
    print("測試 get_required_documents 功能")
    print("="*70)
    
    # 測試案例1: 只有一個 registration_type
    print("\n【測試案例1】單一變更類型")
    print("-"*70)
    test_types_1 = ["遷址(同縣市)"]
    result_1 = get_required_documents(test_types_1)
    print(f"輸入的變更類型: {test_types_1}")
    print(f"回傳的文件數量: {len(result_1)}")
    print("回傳的文件清單:")
    for i, doc in enumerate(result_1, 1):
        print(f"  {i}. {doc['name']} (類型: {doc['type']})")
    
    # 測試案例2: 多個 registration_type
    print("\n【測試案例2】多個變更類型 (過濾議事錄、同意書與公司章程)")
    print("-"*70)
    test_types_2 = ["遷址(不同縣市)", "重新改選董監事", "變更負責人"]
    result_2 = get_required_documents(test_types_2)
    print(f"輸入的變更類型: {test_types_2}")
    print(f"回傳的文件數量: {len(result_2)} (僅包含'議事錄'、'同意書'或'公司章程')")
    print("回傳的文件清單:")
    for i, doc in enumerate(result_2, 1):
        print(f"  {i}. {doc['name']} (類型: {doc['type']})")
    
    # 測試案例3: 設立登記
    print("\n【測試案例3】設立登記")
    print("-"*70)
    test_types_3 = ["設立"]
    result_3 = get_required_documents(test_types_3)
    print(f"輸入的變更類型: {test_types_3}")
    print(f"回傳的文件數量: {len(result_3)}")
    print("回傳的文件清單:")
    for i, doc in enumerate(result_3, 1):
        print(f"  {i}. {doc['name']} (類型: {doc['type']})")
    
    # 測試案例4: 包含設立的多個類型
    print("\n【測試案例4】包含設立的多個類型 (過濾議事錄、同意書與公司章程)")
    print("-"*70)
    test_types_4 = ["設立", "遷址(同縣市)"]
    result_4 = get_required_documents(test_types_4)
    print(f"輸入的變更類型: {test_types_4}")
    print(f"回傳的文件數量: {len(result_4)} (僅包含'議事錄'、'同意書'或'公司章程')")
    print("回傳的文件清單:")
    for i, doc in enumerate(result_4, 1):
        print(f"  {i}. {doc['name']} (類型: {doc['type']})")
    
    print("\n" + "="*70)
    print("測試完成!")
    print("="*70)

if __name__ == "__main__":
    test_get_required_documents()

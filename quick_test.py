from get_required_documents import get_required_documents
import json

print("=" * 70)
print("測試1: 單一變更類型 (遷址同縣市)")
print("=" * 70)
result1 = get_required_documents(["遷址(同縣市)"])
print(f"文件數量: {len(result1)}")
for i, doc in enumerate(result1, 1):
    print(f"  {i}. {doc['name']}")

print("\n" + "=" * 70)
print("測試2: 多個變更類型 (遷址不同縣市 + 重新改選董監事)")
print("過濾條件: 議事錄、同意書、公司章程")
print("=" * 70)
result2 = get_required_documents(["遷址(不同縣市)", "重新改選董監事"])
print(f"文件數量: {len(result2)}")
for i, doc in enumerate(result2, 1):
    print(f"  {i}. {doc['name']}")

print("\n" + "=" * 70)
print("完整輸出 (JSON格式):")
print("=" * 70)
print(json.dumps(result2, ensure_ascii=False, indent=2))

# 這是修改 data 格式的說明文件
# 請手動將以下內容複製到 test_script.py 的 data["table_data"] 部分

"""
將原本的:
    "board_minutes_table": [
      {
        "item": "1",
        "AGENDA_TITLE": "訂定公司章程案",
        "AGENDA_DESC": "依公司法第129條規定，擬定公司章程，如附章程。",
        "AGENDA_RESOLUTION": "經主席徵詢全體發起人無異議照案通過。"
      }
    ],

改為:
    "board_minutes_table": {
      "tabel_ANCHOR": "[TABLE_BOARD_MINUTES_ANCHOR]",
      "column_widths": [1.5, 2.0, 14.5],
      "agenda_rows": [
        {"項目": "1", "案由": "訂定公司章程案"},
        {"項目": "", "說明": "依公司法第129條規定,擬定公司章程,如附章程。"},
        {"項目": "", "決議": "經主席徵詢全體發起人無異議照案通過。"}
      ]
    },

同樣地,將:
    "shareholders_minutes_table": [
      {
        "item": "1",
        "AGENDA_TITLE": "訂定公司章程案",
        "AGENDA_DESC": "依公司法第129條規定，擬定公司章程，如附章程。",
        "AGENDA_RESOLUTION": "經主席徵詢全體發起人無異議照案通過。"
      }
    ],

改為:
    "shareholders_minutes_table": {
      "tabel_ANCHOR": "[TABLE_SHAREHOLDERS_MINUTES_ANCHOR]",
      "column_widths": [1.5, 2.0, 14.5],
      "agenda_rows": [
        {"項目": "1", "案由": "訂定公司章程案"},
        {"項目": "", "說明": "依公司法第129條規定,擬定公司章程,如附章程。"},
        {"項目": "", "決議": "經主席徵詢全體股東無異議照案通過。"}
      ]
    },
"""

# 然後在 generate_document() 函式中,找到處理文件的部分,添加以下邏輯:

"""
在 for doc_name in required_docs: 迴圈中,添加:

        # 檢查是否為董事會議事錄或股東會議事錄
        if "董事會議事錄" in doc_name:
            board_minutes_data = data["table_data"].get("board_minutes_table")
            if board_minutes_data and isinstance(board_minutes_data, dict):
                process_minutes_table(doc, board_minutes_data)
        elif "股東會議事錄" in doc_name or "發起人會議事錄" in doc_name:
            shareholders_minutes_data = data["table_data"].get("shareholders_minutes_table")
            if shareholders_minutes_data and isinstance(shareholders_minutes_data, dict):
                process_minutes_table(doc, shareholders_minutes_data)
        else:
            # 其他文件使用原有的處理方式
            replace_all_placeholders(doc, replacements)
"""

print("✅ 說明文件已創建")
print("📝 請按照以上說明手動修改 test_script.py")

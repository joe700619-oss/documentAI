#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
整合腳本 - 自動修改 test_script.py
"""

import re

# 讀取文件
with open('test_script.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 步驟1: 修改 board_minutes_table
old_board = '''    "board_minutes_table": [
      {
        "item": "1",
        "AGENDA_TITLE": "訂定公司章程案",
        "AGENDA_DESC": "依公司法第129條規定，擬定公司章程，如附章程。",
        "AGENDA_RESOLUTION": "經主席徵詢全體發起人無異議照案通過。"
      }
    ],'''

new_board = '''    "board_minutes_table": {
      "tabel_ANCHOR": "[TABLE_BOARD_MINUTES_ANCHOR]",
      "column_widths": [1.5, 2.0, 14.5],
      "agenda_rows": [
        {"項目": "1", "案由": "訂定公司章程案"},
        {"項目": "", "說明": "依公司法第129條規定,擬定公司章程,如附章程。"},
        {"項目": "", "決議": "經主席徵詢全體發起人無異議照案通過。"}
      ]
    },'''

content = content.replace(old_board, new_board)

# 步驟2: 修改 shareholders_minutes_table
old_shareholders = '''    "shareholders_minutes_table": [
      {
        "item": "1",
        "AGENDA_TITLE": "訂定公司章程案",
        "AGENDA_DESC": "依公司法第129條規定，擬定公司章程，如附章程。",
        "AGENDA_RESOLUTION": "經主席徵詢全體發起人無異議照案通過。"
      }
    ],'''

new_shareholders = '''    "shareholders_minutes_table": {
      "tabel_ANCHOR": "[TABLE_SHAREHOLDERS_MINUTES_ANCHOR]",
      "column_widths": [1.5, 2.0, 14.5],
      "agenda_rows": [
        {"項目": "1", "案由": "訂定公司章程案"},
        {"項目": "", "說明": "依公司法第129條規定,擬定公司章程,如附章程。"},
        {"項目": "", "決議": "經主席徵詢全體股東無異議照案通過。"}
      ]
    },'''

content = content.replace(old_shareholders, new_shareholders)

# 步驟3: 修改處理邏輯
old_logic = '''            # 2. 表格填充 (如果是董事會議事錄，且有議程資料)
            # 這裡簡單判斷，如果文件內有該 Tag 才執行
            # 或是根據 doc_name 判斷也可以，但 Tag 判斷較通用
            if find_table_by_header(doc, AGENDA_HEADER_TAG):
                 add_agendas_to_table(doc, AGENDA_HEADER_TAG, agenda_data)'''

new_logic = '''            # 2. 表格填充 - 使用新的 insert 方式
            if "董事會議事錄" in doc_name:
                # 使用新的 insert 方式處理董事會議事錄
                board_minutes_data = data["table_data"].get("board_minutes_table")
                if board_minutes_data and isinstance(board_minutes_data, dict):
                    process_minutes_table(doc, board_minutes_data)
            elif "股東會議事錄" in doc_name or "發起人會議事錄" in doc_name:
                # 使用新的 insert 方式處理股東會/發起人會議事錄
                shareholders_minutes_data = data["table_data"].get("shareholders_minutes_table")
                if shareholders_minutes_data and isinstance(shareholders_minutes_data, dict):
                    process_minutes_table(doc, shareholders_minutes_data)'''

content = content.replace(old_logic, new_logic)

# 寫回文件
with open('test_script.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✅ 整合完成!")
print("📝 已修改:")
print("   1. board_minutes_table 格式")
print("   2. shareholders_minutes_table 格式")
print("   3. 文件處理邏輯")

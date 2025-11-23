<<<<<<< HEAD
import os 
import json 
import re 
import ast # 引入 ast 模組用於穩定的 JSON (Python dict-like) 解析
from dotenv import load_dotenv 
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS 
from langchain_openai import OpenAIEmbeddings, ChatOpenAI 
from docx import Document # 重新引入 docx 模組

# 路徑設定
HISTORY_DIR = "history_docs"
# 請再次確認您的範本檔名和路徑是正確的
TEMPLATE_PATH = "templates/board_minutes_template.docx" 
OUTPUT_DIR = "output"

load_dotenv()
print(f"OPENAI_API_KEY Loaded: {bool(os.environ.get('OPENAI_API_KEY'))}")


# -----------------------------
# STEP 1：讀取 Word（含表格）
# -----------------------------
def extract_text_from_docx(path):
    """提取 Word 文件的文本內容 (用於 RAG)"""
    document = Document(path)
    texts = []

    for p in document.paragraphs:
        texts.append(p.text)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texts.append(p.text)

    return "\n".join(texts)


def load_history_docs():
    docs = []
    print(f"檢查歷史文件於目錄：{HISTORY_DIR}")
    if not os.path.exists(HISTORY_DIR):
        print(f"🚨 錯誤：找不到歷史文件目錄 {HISTORY_DIR}")
        return docs
        
    for filename in os.listdir(HISTORY_DIR):
        if filename.endswith(".docx") or filename.endswith(".doc"):
            try:
                text = extract_text_from_docx(os.path.join(HISTORY_DIR, filename))
                if text.strip():
                    docs.append({"filename": filename, "text": text})
                else:
                    print(f"文件 {filename} 內容為空。")
            except Exception as e:
                # 提示：如果文件路徑錯誤 (PackageNotFoundError)，會在這裡報錯
                print(f"文件 {filename} 讀取失敗: {e}")
                
    return docs


def build_vector_db(docs):
    splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    texts, metadatas = [], []

    for doc in docs:
        chunks = splitter.split_text(doc["text"])
        for chunk in chunks:
            if chunk.strip():
                texts.append(chunk)
                metadatas.append({"filename": doc["filename"]})

    if not texts:
        print("🚨 警告：沒有可供向量化的文本內容。")
        return None

=======
import os
from dotenv import load_dotenv

# 呼叫 load_dotenv()，它會自動尋找並載入 .env 檔案中的變數到環境中
load_dotenv()

# --- 新增這行來測試 ---
# --- 新增這行來測試 ---
api_key_check = os.environ.get('OPENAI_API_KEY')
print(f"OPENAI_API_KEY is loaded: {bool(api_key_check)}")
# ---
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from docx import Document

HISTORY_DIR = "history_docs"
TEMPLATE_PATH = "templates/變更登記表_範本.docx"
OUTPUT_DIR = "output"

# -----------------------------
# STEP 1：讀取歷史文件並建立向量資料庫
# -----------------------------
def load_history_docs():
    docs = []
    for filename in os.listdir(HISTORY_DIR):
        if filename.endswith(".docx"):
            text = extract_text_from_docx(os.path.join(HISTORY_DIR, filename))
            docs.append({"filename": filename, "text": text})
    return docs

def extract_text_from_docx(path):
    document = Document(path)
    return "\n".join([p.text for p in document.paragraphs])

def build_vector_db(docs):
    splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    texts = []
    metadatas = []
    
    for doc in docs:
        chunks = splitter.split_text(doc["text"])
        for chunk in chunks:
            texts.append(chunk)
            metadatas.append({"filename": doc["filename"]})
    
>>>>>>> 8e51595c7031e762eba481fc172cceab0a832b0d
    embeddings = OpenAIEmbeddings()
    db = FAISS.from_texts(texts, embeddings, metadatas=metadatas)
    return db

<<<<<<< HEAD

# ------------------------------------
# STEP 2：搜尋 + 過濾符合公司的歷史紀錄
# ------------------------------------
def search_history(db, company):
    if db is None:
        return []
    results = db.similarity_search(company, k=20)
    return results


def filter_by_company(results, company):
    """只保留與該公司有關的歷史紀錄"""
    filtered = []
    for r in results:
        if company in r.page_content:
            filtered.append(r)
    return filtered


def summarize_history(results):
    """列出給使用者選擇"""
    summary_list = []
    for r in results:
        summary_list.append(f"- {r.metadata['filename']}")
    return "\n".join(summary_list)


# ------------------------------------
# STEP 3：AI → 回傳 JSON (包含本次變更合併邏輯)
# ------------------------------------
def extract_key_info(selected_results, company, change):
    model = ChatOpenAI(model="gpt-4o-mini")

    merged_text = "\n\n".join([
        f"【{r.metadata['filename']}】\n{r.page_content}"
        for r in selected_results
    ])

    prompt = f"""
你是一個專業的登記資訊抽取系統。
你的任務是基於「歷史紀錄」和「本次變更指令」來生成完整的資料。
請優先使用「本次變更指令」來更新或填寫相關欄位。
請依照以下要求輸出純 JSON，不得夾帶說明文字：

◎輸入內容為「公司歷史變更資料」+「本次變更指令」  
◎請輸出：
1. files: 本次參考的文件列表（含推斷的文件類型）
2. summaries: 每份文件的摘要
3. merged_change: 把「歷史內容」與「本次變更」合併後的最終版本
4. registration_data: 可直接填進變更登記表的欄位資訊（可留空）

以下是歷史文件內容：
{merged_text}

本次變更說明：{change}

請輸出此 JSON 格式：
{{
  "files": [],
  "summaries": [],
  "merged_change": "",
  "registration_data": {{
        "company_name": "",
        "registration_number": "",
        "zipcode": "",
        "address": "",
        "chairperson": "",
        "change_type": "",
        "new_directors": [],
        "removal_directors": []
  }}
}}
"""

    result = model.invoke(prompt)
    return result.content.strip()


# ------------------------------------
# STEP 4：清洗 JSON + 填入 Word (已修正：包含 JSON & Word 替換修正)
# ------------------------------------
def clean_json_text(text):
    """
    強化清洗 AI 回傳的 JSON，處理 Markdown 標記、最外層括號，並移除非打印字符 U+00A0。
    """
    # 1. 移除 ```json ... ``` (如果 AI 仍然輸出 Markdown 格式)
    text = re.sub(r"```.*?```", "", text, flags=re.S)
    
    # 2. 替換頑固的 Non-breaking Space (U+00A0)，這是導致解析錯誤的元兇
    text = text.replace(u'\xa0', ' ') 

    # 3. 只保留最外層的 { ... }
    match = re.search(r"\{.*\}", text, flags=re.S)
    
    # 4. 移除 JSON 結尾多餘的逗號
    if match:
        cleaned_text = match.group(0).strip()
        cleaned_text = re.sub(r',\s*([\]\}])', r'\1', cleaned_text) 
        return cleaned_text
    else:
        return "{}"


def _replace_in_paragraph(para, mapping):
    """處理單個 Paragraph 物件中的佔位符替換邏輯 (穩健替換)"""
    full_text = para.text
    
    for placeholder, value in mapping.items():
        # 如果佔位符存在於段落的完整文本中
        if placeholder in full_text:
            # 進行替換
            new_text = full_text.replace(placeholder, value)
            
            # 確保段落有 run
            if len(para.runs) > 0:
                # 只保留第一個 run，並將新文本賦值給它
                first_run = para.runs[0]
                first_run.text = new_text

                # 移除其餘所有的 run，以確保格式和內容的準確性
                while len(para.runs) > 1:
                    run_element_to_remove = para.runs[-1]._element
                    run_element_to_remove.getparent().remove(run_element_to_remove)
                
                # 更新 full_text，以便如果段落包含多個佔位符時能進行連續替換
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


def fill_template(json_data, output_name="本次變更登記表.docx"):
    """使用 ast.literal_eval 作為 JSON 解析回退，並執行 Word 替換"""
    
    # 再次檢查範本路徑，防止 PackageNotFoundError
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 Word 範本：'{TEMPLATE_PATH}'。請檢查路徑和檔名。")
        
    doc = Document(TEMPLATE_PATH)

    cleaned = clean_json_text(json_data)
    
    # 穩健 JSON 解析：優先用 json.loads，失敗則用 ast.literal_eval
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        print("⚠️ JSON 解析失敗，嘗試使用 ast.literal_eval 處理非標準格式...")
        try:
            # 嘗試轉成 Python 字典，並重新序列化回標準 JSON
            py_dict = ast.literal_eval(cleaned)
            data = json.loads(json.dumps(py_dict))
            print("✅ 成功使用 ast.literal_eval 解析！")
        except (SyntaxError, ValueError, TypeError) as e:
            print(f"❌ 無法解析 AI 輸出為有效字典：{e}")
            print("將使用空資料進行替換，請檢查 AI 輸出。")
            data = {} 

    reg = data.get("registration_data", {})
    if not reg:
        print("🚨 警告：AI 輸出的 JSON 中缺少 registration_data 欄位或內容為空。")

    mapping = {
        "{{公司名稱}}": reg.get("company_name", ""),
        "{{公司地址}}": reg.get("address", ""),
        "{{郵遞區號}}": reg.get("zipcode", ""),
        "{{代表人}}": reg.get("chairperson", ""),
        "{{變更類型}}": reg.get("change_type", ""),
        # 您可以根據範本和 AI 輸出，在這裡添加更多需要的欄位
    }

    replace_all_placeholders(doc, mapping)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
=======
# ------------------------------------
# STEP 2：讓 AI 找出最近/最相關的歷史紀錄
# ------------------------------------
def search_history(db, query):
    results = db.similarity_search(query, k=5)
    return results

def summarize_history(results):
    model = ChatOpenAI(model="gpt-4o-mini")
    texts = "\n\n".join([f"【{r.metadata['filename']}】\n{r.page_content}" for r in results])
    
    prompt = f"""
以下是歷史文件內容的摘要，請以「可供使用者選擇的列表格式」整理出重要變更內容：

{texts}
"""
    summary = model.invoke(prompt)
    return summary.content

# ------------------------------------
# STEP 3：依照使用者選定的版本，AI 整理關鍵資訊
# ------------------------------------
def extract_key_info(selected_results):
    model = ChatOpenAI(model="gpt-4o-mini")
    txt = "\n\n".join([r.page_content for r in selected_results])
    
    prompt = f"""
以下是使用者選定的歷史文件內容：

{txt}

請萃取以下資訊：
1. 上次的董事/監察人名單（含職稱）
2. 上次異動摘要
3. 與本次變更可能相關的決議格式
4. 需要填入政府變更登記表的欄位資料

請以 JSON 格式輸出。
"""
    result = model.invoke(prompt)
    return result.content

# ------------------------------------
# STEP 4：將資訊填入 Word 範本
# ------------------------------------
def fill_template(json_data, output_name="本次變更登記表.docx"):
    doc = Document(TEMPLATE_PATH)

    # 非正式示範：替換 {{placeholder}}
    for p in doc.paragraphs:
        if "{{資料}}" in p.text:
            p.text = p.text.replace("{{資料}}", json_data)

>>>>>>> 8e51595c7031e762eba481fc172cceab0a832b0d
    output_path = os.path.join(OUTPUT_DIR, output_name)
    doc.save(output_path)
    return output_path

<<<<<<< HEAD

=======
>>>>>>> 8e51595c7031e762eba481fc172cceab0a832b0d
# ------------------------------------
# 主流程
# ------------------------------------
if __name__ == "__main__":
<<<<<<< HEAD
    # Step 1: 使用者輸入
    company = input("請輸入公司名稱：")
    change = input("請描述本次變更（例如：變更董事）: ")

    # Step 2: 找歷史資料
    print("\n📌 STEP 1：讀取歷史紀錄...")
    history_docs = load_history_docs()
    
    if not history_docs:
        print("🚨 錯誤：HISTORY_DIR 資料夾中沒有找到任何文件，或文件讀取失敗。")
        exit()

    print("\n📌 STEP 2：建構向量資料庫並搜尋該公司的歷史紀錄...")
    db = build_vector_db(history_docs)
    
    if db is None:
        print("🚨 錯誤：無法建構向量資料庫，請檢查歷史文件內容或安裝所需套件。")
        exit()
        
    results = search_history(db, company)
    filtered = filter_by_company(results, company)

    print("\n🔍 找到以下歷史紀錄：")
    
    if not filtered:
        print("⚠️ 警告：沒有找到該公司相關的歷史文件。")
        chosen_results = []
    else:
        print(summarize_history(filtered))
        
        # Step 3: 讓使用者選擇
        selected = input("\n請輸入你要參考的檔案名稱（用逗號分隔）：\n")
        selected_list = [s.strip() for s in selected.split(",")]
        chosen_results = [r for r in filtered if r.metadata["filename"] in selected_list]


    print("\n📌 STEP 3：AI 分析文件並合併變更內容...")
    raw_json = extract_key_info(chosen_results, company, change)

    print("\n🧩 AI 回傳 JSON（原始內容）：\n", raw_json)

    # Step 4: 產生 Word 文件
    confirm = input("\n是否要產生新的變更登記表？(Y/N)：").upper()
    if confirm == "Y":
        print("\n📌 STEP 4：套用 Word 範本...")
        try:
            output_path = fill_template(raw_json)
            print(f"\n🎉 完成！文件已輸出至：{output_path}")
        except FileNotFoundError as e:
            print(f"\n❌ 錯誤：{e}")
            print("請確認 templates 資料夾和範本檔名是否正確！")
    else:
        print("\n❌ 已取消，不產生文件。")
=======
    company = input("請輸入公司名稱：")
    change = input("請描述本次變更（例如：變更董事）:")

    print("\n📌 STEP 1：讀取歷史紀錄...")
    history_docs = load_history_docs()
    db = build_vector_db(history_docs)

    print("\n📌 STEP 2：AI 搜尋相似歷史紀錄...")
    results = search_history(db, f"{company} {change}")
    summary = summarize_history(results)
    print("\n🔍 找到以下歷史紀錄：\n")
    print(summary)

    # 讓使用者選擇
    selected = input("\n請輸入你要參考的檔案名稱（用逗號分隔）：\n")
    selected_list = [s.strip() for s in selected.split(",")]

    selected_results = [r for r in results if r.metadata["filename"] in selected_list]

    print("\n📌 STEP 3：AI 萃取關鍵資料...")
    json_data = extract_key_info(selected_results)
    print("\n🧩 AI 萃取的資料：\n", json_data)

    print("\n📌 STEP 4：套用 Word 範本...")
    output_path = fill_template(json_data)

    print(f"\n🎉 完成！文件已輸出至：{output_path}")
>>>>>>> 8e51595c7031e762eba481fc172cceab0a832b0d

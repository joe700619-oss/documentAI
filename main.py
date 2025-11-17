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
    
    embeddings = OpenAIEmbeddings()
    db = FAISS.from_texts(texts, embeddings, metadatas=metadatas)
    return db

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

    output_path = os.path.join(OUTPUT_DIR, output_name)
    doc.save(output_path)
    return output_path

# ------------------------------------
# 主流程
# ------------------------------------
if __name__ == "__main__":
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

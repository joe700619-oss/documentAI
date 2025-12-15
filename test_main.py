import os
import json
from getbasicInformationfromMOEA import BasicInformationAPI
import re
from dotenv import load_dotenv

from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_core.documents import Document
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser

# PDF / DOCX
from pypdf import PdfReader
import docx

# Unstructured
from unstructured.partition.auto import partition

load_dotenv()


# ==========================================================
# 文件類型定義
# ==========================================================
DOCUMENT_TYPES = {
    "articles_of_association",
    "meeting_minutes",
    "director_roster",
    "shareholder_roster",
    "change_registration_form",
    "establishment_registration_form",
    "unknown"
}


# ==========================================================
# 文件分類器
# ==========================================================
def detect_document_type(filename: str, text: str):
    name = filename.lower()
    head = text[:8000]

    if "變更" in name and "登記" in name:
        return "change_registration_form"

    if "設立" in name and "登記" in name:
        return "establishment_registration_form"

    if "董事" in name and "名單" in name:
        return "director_roster"

    if "股東" in name and "名單" in name:
        return "shareholder_roster"

    if "章程" in name:
        return "articles_of_association"

    if "會議" in name or "議事錄" in name:
        return "meeting_minutes"

    # content fallback
    if "變更登記" in head and head.count("董事") >= 3:
        return "change_registration_form"

    if "設立登記" in head and head.count("董事") >= 3:
        return "establishment_registration_form"

    if "出席" in head and "決議" in head:
        return "meeting_minutes"

    if re.search(r"第.{1,3}條", head):
        return "articles_of_association"

    return "unknown"


# ==========================================================
# 文件載入
# ==========================================================
def load_documents(folder):
    docs = []

    for f in os.listdir(folder):
        path = os.path.join(folder, f)
        if not os.path.isfile(path):
            continue

        content = ""

        try:
            if path.endswith(".pdf") or path.endswith(".docx"):
                elements = partition(
                    filename=path,
                    strategy="hi_res",
                    max_characters=4000
                )
                content = "\n\n".join(str(el) for el in elements)
            else:
                with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                    content = fh.read()
        except Exception:
            # fallback
            try:
                if path.endswith(".pdf"):
                    reader = PdfReader(path)
                    content = "\n".join(p.extract_text() for p in reader.pages if p.extract_text())
                elif path.endswith(".docx"):
                    doc = docx.Document(path)
                    content = "\n".join(p.text for p in doc.paragraphs)
            except Exception:
                continue

        if not content.strip():
            continue

        dtype = detect_document_type(f, content)

        docs.append(Document(
            page_content=content,
            metadata={"source": f, "doc_type": dtype}
        ))

    return docs


# ==========================================================
# 資料抽取（整合 MOEA 與 文件）
# ==========================================================
def select_registration_type(llm, user_request: str, allowed_types: list, current_info: dict):
    """
    Step 2: 讓 LLM 從 allowed_types 中選擇符合 user_request 的變更類型。
    """
    allowed_str = json.dumps(allowed_types, ensure_ascii=False, indent=2)
    prompt = ChatPromptTemplate.from_template("""
你是一個專業的法律與公司登記顧問。
請根據【使用者需求】與【目前公司資料】，從【合法變更類型列表】中選擇一個或多個最符合的變更類型。

特別注意：
1. 若涉及遷址：
   請比較【使用者需求】中的新地址與【目前公司資料】中的舊地址。
   - 若縣市名稱相同（例如都在台北市），請選擇 "遷址(同縣市)"。
   - 若縣市名稱不同（例如從台中市搬到台北市），請選擇 "遷址(不同縣市)"。
2. 若無法從目前資料判斷，請依常理推斷或選擇較寬鬆的選項。

【使用者需求】：{user_request}

【目前公司資料】：
{current_info}

【合法變更類型列表】：
{allowed_types}

請直接回傳一個 JSON Array，包含所有符合的類型字串。
例如：["變更負責人", "遷址(同縣市)"]
若無匹配，回傳 []。
只回傳 JSON，不要有其他說明。
""")
    chain = prompt | llm | StrOutputParser()
    result = chain.invoke({
        "user_request": user_request,
        "current_info": json.dumps(current_info, ensure_ascii=False, indent=2),
        "allowed_types": allowed_str
    })
    
    try:
        cleaned_result = result.strip()
        if cleaned_result.startswith("```json"):
            cleaned_result = cleaned_result[7:]
        if cleaned_result.startswith("```"):
            cleaned_result = cleaned_result[3:]
        if cleaned_result.endswith("```"):
            cleaned_result = cleaned_result[:-3]
        return json.loads(cleaned_result.strip())
    except Exception as e:
        print(f"Error parsing registration type selection: {e}")
        return []

def extract_company_data(llm, full_text: str, moea_data: dict, user_request: str, selected_types: list, target_schema: dict):
    """
    Step 5: 依照 format_example (target_schema) 回傳資料。
    需填寫 application_reason, registration_type。
    需整合 MOEA (locked) 與 Document (editable)。
    """

    moea_info_str = json.dumps(moea_data, ensure_ascii=False, indent=2)
    schema_str = json.dumps(target_schema, ensure_ascii=False, indent=2)
    selected_types_str = ", ".join(selected_types)

    prompt = ChatPromptTemplate.from_template("""
你是一個法律文件資料抽取引擎。

任務：
請根據以下來源資訊，產出符合【目標 JSON 格式】的資料。

來源資訊：
1. 【使用者需求】：{user_request}
2. 【選定的變更類型】：{selected_types}
3. 【MOEA 權威資料】 (Source of Truth)：包含已知的公司基本資料與欄位鎖定規則(locked/editable)。
4. 【文件內容】：變更登記表或設立登記表全文。

⚠️ 資料整合與填寫規則：
1. **registration_type**：必須填入【選定的變更類型】（若有多個用逗號分隔）。
2. **application_reason**：請根據【使用者需求】摘要填寫（例如："更換負責人為王大明"）。
3. **MOEA 鎖定欄位**：若 MOEA 資料中某欄位標記為 "locked"，且使用者需求未明確要求修改該欄位，請直接使用 MOEA 的值。
4. **缺失與補足**：若 MOEA 資料缺漏或標記為 "editable"，請從【文件內容】中尋找最新資訊填入。
5. **格式一致性**：輸出的 JSON 結構必須與【目標 JSON 格式】完全一致。

【MOEA 權威資料】
{moea_info}

【目標 JSON 格式 (範例)】
{target_schema}

【文件內容】
{content}

請依照上述目標格式輸出 JSON 資料 (只輸出 JSON，不要說明文字)：
""")

    chain = prompt | llm | StrOutputParser()
    result = chain.invoke({
        "content": full_text,
        "moea_info": moea_info_str,
        "selected_types": selected_types_str,
        "user_request": user_request,
        "target_schema": schema_str
    })

    try:
        cleaned_result = result.strip()
        if cleaned_result.startswith("```json"):
            cleaned_result = cleaned_result[7:]
        if cleaned_result.startswith("```"):
            cleaned_result = cleaned_result[3:]
        if cleaned_result.endswith("```"):
            cleaned_result = cleaned_result[:-3]
        return json.loads(cleaned_result.strip())
    except Exception as e:
        return {"error": "JSON parse failed", "raw": result}


# ==========================================================
# 主流程
# ==========================================================
class DebugWorkflow:

    def __init__(self):
        self.history_dir = "history"
        self.history_cases_dir = "history_cases"
        self.api_key = os.getenv("OPENAI_API_KEY")

        self.llm = ChatOpenAI(
            model="gpt-4o",
            temperature=0,
            api_key=self.api_key
        )
        self.moea_api = BasicInformationAPI()

        # 載入合法變更類型
        self.allowed_registration_types = []
        try:
            with open("documents_required_list.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data.get("documents_required_list", []):
                    rtype = item.get("Registration_Type")
                    if rtype and rtype != "default":
                        self.allowed_registration_types.append(rtype)
        except Exception as e:
            print(f"Warning: Failed to load documents_required_list.json: {e}")

        # 載入目標 JSON Schema
        self.target_schema = {}
        try:
            with open("company_registration_form.json", "r", encoding="utf-8") as f:
                self.target_schema = json.load(f)
        except Exception as e:
            print(f"Warning: Failed to load company_registration_form.json: {e}")

    def run(self):
        # 1. 載入文件並過濾 (只保留變更登記表/設立登記表)
        # 依照需求：先在資料夾中，只有找出變更登記表，或是設立登記表...
        print("\n=== STEP 1: searching Documents ===")
        all_docs = load_documents(self.history_dir) + load_documents(self.history_cases_dir)
        
        target_types = ["change_registration_form", "establishment_registration_form"]
        filtered_docs = [d for d in all_docs if d.metadata["doc_type"] in target_types]

        if not filtered_docs:
            print("未找到「變更登記表」或「設立登記表」，流程結束。")
            return

        print(f"Found {len(filtered_docs)} registration form(s).")
        full_text = "\n\n".join(d.page_content for d in filtered_docs)

        # 2. 取得 MOEA 基本資料 (提前至此以便判斷變更類型)
        print("\n=== STEP 2: Fetching MOEA Data ===")
        # 嘗試從文本中抓取統一編號
        tax_id_match = re.search(r"統一編號[：:\s]*(\d{8})", full_text)
        if not tax_id_match:
            tax_id_match = re.search(r"\b(\d{8})\b", full_text)
        
        tax_id = tax_id_match.group(1) if tax_id_match else "60299784" # Fallback
        print(f"Detected Tax ID: {tax_id}")

        print("Fetching MOEA Data...")
        moea_facts = self.moea_api.get_company_facts(tax_id)
        
        # 建構 Context
        moea_context = {
            "values": moea_facts,
            "policy": {
                "companyName": "locked",
                "authorizedCapital": "locked",
                "companyAddress": "locked",
                "chairmanName": "locked",
                "business_items": "locked", # 假設 API 回傳目前最新的，設為 locked (除非使用者要改)
            }
        }

        # 3. 使用者輸入 & LLM 選擇變更類型
        print("\n=== STEP 3: User Input & Classification ===")
        user_input_request = input("\n請輸入您的變更需求 (例如：更換負責人為王大明、遷址...): ").strip()
        
        print("Identifying Registration Type...")
        selected_types = select_registration_type(
            self.llm, 
            user_input_request,
            self.allowed_registration_types,
            current_info=moea_facts 
        )
        print(f"Selected Types: {selected_types}")

        # 4. & 5. 填入資料並回傳 JSON
        print("\n=== STEP 4 & 5: Filling Form & Generating JSON ===")
        print("Extracting & Merging Data...")
        result_json = extract_company_data(
            self.llm, 
            full_text, 
            moea_context, 
            user_input_request,
            selected_types,
            self.target_schema
        )

        print("\n=== Final JSON Output ===")
        print(json.dumps(result_json, ensure_ascii=False, indent=2))
        
        # Save to file
        with open("final_llm_output.json", "w", encoding="utf-8") as f:
            json.dump(result_json, f, ensure_ascii=False, indent=2)
        print("\n(已將結果儲存至 final_llm_output.json)")

if __name__ == "__main__":
    DebugWorkflow().run()

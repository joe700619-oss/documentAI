import time

print("Starting imports...")

try:
    t0 = time.time()
    print("Importing langchain_openai...")
    from langchain_openai import ChatOpenAI, OpenAIEmbeddings
    print(f"langchain_openai imported in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"Failed to import langchain_openai: {e}")

try:
    t0 = time.time()
    print("Importing langchain_community.vectorstores FAISS...")
    from langchain_community.vectorstores import FAISS
    print(f"FAISS imported in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"Failed to import FAISS: {e}")

try:
    t0 = time.time()
    print("Importing langchain_text_splitters...")
    from langchain_text_splitters import RecursiveCharacterTextSplitter
    print(f"langchain_text_splitters imported in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"Failed to import langchain_text_splitters: {e}")

try:
    t0 = time.time()
    print("Importing langchain_core...")
    from langchain_core.documents import Document
    print(f"langchain_core imported in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"Failed to import langchain_core: {e}")

try:
    t0 = time.time()
    print("Importing unstructured...")
    from unstructured.partition.auto import partition
    print(f"unstructured imported in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"Failed to import unstructured: {e}")

print("Imports done.")

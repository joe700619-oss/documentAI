from docx import Document
import os

path = r"C:\Users\joe70\PythonProject\documentAI\wordformat\word範本_patched.docx"
print(f"File exists: {os.path.exists(path)}")
try:
    doc = Document(path)
    print("Successfully opened document")
except Exception as e:
    print(f"Error: {e}")

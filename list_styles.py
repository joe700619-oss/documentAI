from docx import Document

path = r"C:\Users\joe70\PythonProject\documentAI\templates\設立登記表.docx"
doc = Document(path)

print("Paragraph styles:")
for style in doc.styles:
    if style.type.name == 'PARAGRAPH':
        print(f"- {style.name} (id: {style.style_id})")

import zipfile
import os

path = r"C:\Users\joe70\PythonProject\documentAI\wordformat\word範本.docx"
temp_path = r"C:\Users\joe70\PythonProject\documentAI\wordformat\word範本_patched.docx"

def patch_docx(src, dst):
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    data = data.replace(
                        b'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
                        b'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
                    )
                zout.writestr(item, data)

if os.path.exists(path):
    patch_docx(path, temp_path)
    print(f"Patched file created at {temp_path}")
else:
    print(f"File not found: {path}")

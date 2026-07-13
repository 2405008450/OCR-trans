from docx import Document
from revision import RevisionManager

doc = Document(r"C:\Users\Administrator\Desktop\多语种标点检查\多语种测试文件\译文修订——马来语.docx")
rm = RevisionManager(doc, author="标点检查")

# 在段落中替换标点
for para in doc.paragraphs:
    rm.replace_in_paragraph(para, ".", "{", reason="中文逗号误用为英文")

doc.save(r"译文修订完成——马来语.docx")

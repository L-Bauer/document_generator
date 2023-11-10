from docx import Document

document = Document("/Users/umityalcin/Desktop/Test.docx")
Dictionary = {"sea": "ocean"}

for paragraph in document.paragraphs:
    if "sea" in paragraph.text:
        paragraph.text = paragraph.text.replace("sea", Dictionary["sea"])

document.save("/Users/umityalcin/Desktop/Test.docx")

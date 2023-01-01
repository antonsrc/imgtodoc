import os
import docx
from docx.shared import Cm

doc_file = r'D:\d\misc\out.docx'
img_dir = r'D:\d\misc\data'

doc = docx.Document(doc_file)

tree = list(os.walk(img_dir))
for address, dirs, files in tree:
    for name in files:
        doc.add_picture(img_dir + '\\' + name, width = Cm(10))

doc.save(doc_file)

# Если tiff многостраничный, то чтобы разбить файл на несколько файлов:
# 1 Сохраняем в PDFCreator в формате pdf
# 2 В PDF-XChange выделить страницы, ПКМ и Export to Image...
# 3 Экспортировать в JPEG
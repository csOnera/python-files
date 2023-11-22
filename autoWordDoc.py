from docx.api import Document

document = Document(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\雜\統一托盤，紙箱標識.docx")

for table in document.tables:
    print('new table')
    for row in table.rows:
        for cell in row.cells:
            print(cell.text)
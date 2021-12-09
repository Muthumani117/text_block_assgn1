import fitz   # PyMUPdf library
import xlsxwriter

# pdf_filename = "http://www.kepcorp.com/annualreport2018/pdf/keppel-corporation-limited-annual-report-2018.pdf"
pdf_filename = input("Enter the pdf file path:")
excel_sheet_name = input("Enter excel file name to export the text [Example test1.xlsx]:")

workbook = xlsxwriter.Workbook(excel_sheet_name)        #create file

doc = fitz.open(pdf_filename)
for page in doc:
        blocks = page.get_text("blocks")
        blocks.sort(key = lambda block: block[0])  # sort vertically ascending for natural reading order
        row_iter = 0
        col = 0
        worksheet = workbook.add_worksheet()
        for b in blocks:
                worksheet.write(row_iter, col, b[4].replace('\n', ''))
                row_iter+=1
workbook.close()

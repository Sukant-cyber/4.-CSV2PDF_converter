import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors


def create_pdf_report(input_csv, output_pdf):
    # Read CSV data using pandas
    df = pd.read_csv(input_csv)

    # Convert the DataFrame to a list of lists (2D array) for creating the table
    table_data = [df.columns.tolist()] + df.values.tolist()

    # Create the PDF
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    table = Table(table_data)

    # Apply styles to the table
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.gray),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ])

    table.setStyle(style)

    # Add the table to the PDF
    elements = [table]
    doc.build(elements)

if __name__ == "__main__":
    input_csv_file = "input.xlsx"
    output_pdf_file = "report.pdf"
    create_pdf_report(input_csv_file, output_pdf_file)


# _______________________________________________

# CODE TO CONVERT XLSX FILE TO PDF FILE

# from win32com import client
#
# app = client.DispatchEx("Excel.Application")
# app.Interactive = False
# app.Visible = False
#
# path = "C:\pythonProject\CSV2PDF\input.xlsx"
# print("Converting into PDF, Please Wait...")
# workbook = app.Workbooks.Open(path)
# # worksheet = workbook.ActiveSheet
# # cell_range = worksheet.Range("A1")
# # cell_range.Font.Size = 14
# worksheet = workbook.ActiveSheet
#
# # Adjust the column widths based on your requirement
# worksheet.Columns.AutoFit()
# workbook.ActiveSheet.ExportAsFixedFormat(0, path)
# workbook.Close()
#
# print("Completed!!")



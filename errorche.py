import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('error.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.

def create(error_all):
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0
    worksheet.write(row, col,     "num")
    worksheet.write(row, col + 1, "name")
    worksheet.write(row, col + 2, "email")
    worksheet.write(row, col + 3, "cert_num")
    row += 1
    # Iterate over the data and write it out row by row.
    for num, name, email, cert_num in (error_all):
        worksheet.write(row, col,     num)
        worksheet.write(row, col + 1, name)
        worksheet.write(row, col + 2, email)
        worksheet.write(row, col + 3, cert_num)
        row += 1

    # Write a total using a formula.
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')

    workbook.close()
error_all=[]
def errorcheck(num,name,email,cert_num):
    error_list =[num,name,email,cert_num]
    error_all.append(error_list)

from openpyxl import Workbook

def create_excel():
    # create a new Excel file
    wb = Workbook()

    # get active worksheet
    sheet = wb.active

    # create header
    sheet['A1'] = 'ID'
    sheet['B1'] = 'First Name'
    sheet['C1'] = 'Last Name'
    sheet['D1'] = 'Group ID'

    # add data
    sheet.append([1,'Sophie','Brown',3])
    sheet.append([2,'Zoe','Taylor',4])
    sheet.append([3,'Caren','Tomas',6])

    # save the Excel file
    wb.save('file/student.xlsx')


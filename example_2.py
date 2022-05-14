import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Grades.xlsx')
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

class1_grades = [
    ['Name', 'Score', 'Letter Grade'],
    ['John', 31, 'F'],
    ['Jason', 97, 'A'],
    ['Amith', 59, 'B']
]

class2_grades = [
    ['Name', 'Score', 'Letter Grade'],
    ['Alex', 72, 'C'],
    ['Sebastian', 61, 'D'],
    ['Amith', 59, 'B']
]

def convert_array_to_excel_worksheet(grades_array, worksheet):
    row_index = 0
    column_index = 0

    for row in grades_array:
        for column in row:
            worksheet.write(row_index, column_index, column)
            column_index +=1
        column_index = 0
        row_index +=1
    
    return row_index
    
final_row_index_1 = convert_array_to_excel_worksheet(class1_grades, worksheet1)
final_row_index_2 = convert_array_to_excel_worksheet(class2_grades, worksheet2)

worksheet1.write(final_row_index_1, 0, 'Average Grade')
worksheet1.write(final_row_index_1, 1, '=AVERAGE(B2:B4)')

worksheet1.write(final_row_index_2, 0, 'Average Grade')
worksheet1.write(final_row_index_2, 1, '=AVERAGE(B2:B4)')

workbook.close()






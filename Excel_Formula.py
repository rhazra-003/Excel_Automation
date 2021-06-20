wb = load_workbook('report_2021.xlsx')
sheet = wb['Report']

# sum in columns B-G
for i in excel_alphabet:
    if i!='A':
        sheet[f'{i}{max_row+1}'] = f'=SUM({i}{min_row+1}:{i}{max_row})'
        sheet[f'{i}{max_row+1}'].style = 'Currency'

# adding total label
sheet[f'{excel_alphabet[0]}{max_row+1}'] = 'Total'

wb.save('report_2021.xlsx')

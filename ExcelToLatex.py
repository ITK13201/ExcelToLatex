import openpyxl, pyperclip
from openpyxl.utils import get_column_letter

# make
'''
\begin{table}[h]
			\caption{ageの統計量と代表値}
			\centering
			\begin{tabular}{cr}
				\hline
				統計量または代表値	&	値\\
				\hline \hline
				中央値	&	43.00\\
				最頻値	&	50.00\\
				最大値	&	60.00\\
				最小値	&	20.00\\
				範囲		&	40.00\\
				分散		&	122.35\\
				標準偏差	&	11.06\\
				\hline
			\end{tabular}
		\end{table}
'''

# ExcelWorksheetを読み込む
#wbname = input('Excel-Workbook-Name: ')
wbname = 'data.xlsx'
wb = openpyxl.load_workbook(wbname, data_only=True)
#wbsheet = input('Excel-Workbook-Sheet-Name: ')
wbsheet = 'toLatex'
sheet = wb[wbsheet]
sheet = wb.active

linecode = 'ccccc'
OutputText_begin =  '\\begin{tabular}{' + linecode + '} \n' \
                    '\t\\hline'
    

OutputText_end =    '\tend{tabular}\n' \
                    '\\end\{table}'


# debug #####################
max_cell = get_column_letter(sheet.max_column) + str(sheet.max_row)

for row_of_cell_obj in sheet['B2':max_cell]:
    for cell_obj in row_of_cell_obj:
        cell_obj.value = round(cell_obj.value, 1)

for row_of_cell_obj in sheet['A1':max_cell]:
    for cell_obj in row_of_cell_obj:
        print(cell_obj.value, end='\t')
    print()

print(OutputText_begin)
print(OutputText_end)




##############################

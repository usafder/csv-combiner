from os import listdir, mkdir, path
from time import time
from pandas import ExcelWriter, read_csv

 # get names of all the files and folders within the current directory
files_list = listdir('Input')
# create Output directory if it doesn't exist
if not path.exists('Output'):
    try:
        mkdir('Output')
    except Exception as e:
        print('Failed to create Output directory due to ' + str(e))

output_file = 'Output/Result.xlsx'
xlsxwriter = ExcelWriter(output_file, engine = 'xlsxwriter')
workbook = xlsxwriter.book
sheet_count = 0
start_time = time()

for source_file_name in files_list:
    source_file = f'Input/{source_file_name}'
    if path.isfile(source_file):
        entity_name = source_file_name

        # need this condition since Excel allows max 31 characters for a sheet name
        if len(entity_name) > 31:
            entity_name = f'{entity_name[0:29]}-D'
    
        try:
            df = read_csv(source_file)
            # df = df.fillna('')
            row_count, col_count = df.shape
            row_count = row_count+1
            df.to_excel(xlsxwriter, sheet_name = entity_name, index = False)
            cell_format = workbook.add_format()
            cell_format.set_text_wrap()
            cell_format.set_align('center')
            cell_format.set_align('vcenter')
            worksheet = xlsxwriter.sheets[entity_name]
            # replace A and D with the columns used in the CSV file
            worksheet.set_column('A:D', 35, cell_format)
            cell_range = f'A1:D{row_count}'
            worksheet.autofilter(cell_range)
            sheet_count += 1
        except Exception as e:
            print(f'{source_file} failed due to {str(e)}')
    else:
        print(f'{source_file} is not a file')

print(f'\nNo. of Sheets added: {sheet_count}')

try:
    print('\nSaving changes...')
    xlsxwriter.save()
    print('Saved successfully')
except Exception as e:
    print(f'Failed to save changes due to {str(e)}')

end_time = time()
total_elapsed_time = end_time - start_time
print(f'Total Elapsed Time: {total_elapsed_time} seconds')

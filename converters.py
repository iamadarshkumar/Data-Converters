import xlrd
import unicodecsv
import os
import csv
import shutil
import math

def xlsx_to_custom_delimited_file(input_dir, output_dir, extension_to_save, delimiter, row_start_index): 

	try:
		wb = xlrd.open_workbook(input_dir)
	except xlrd.biffh.XLRDError as e:
		return e
	except FileNotFoundError as e:
		return e
	except:
		return "Failed to load input file."

	sheetnames = wb.sheet_names()
	print('Sheet Names', sheetnames)
	
	for sheetname in sheetnames:
		empty_count_flag = False

		try:
			fp = open(os.path.join(output_dir+'\\'+sheetname+'.'+extension_to_save),'w', encoding="utf8", newline='')

			# Open the sheet by name
			xl_sheet = wb.sheet_by_name(sheetname)

			# Iteration through cells
			for row_idx in range(row_start_index, xl_sheet.nrows):

				empty_count = 0  # Count for empty cells
				for col_idx in range(0, xl_sheet.ncols):
					cell_obj = xl_sheet.cell(row_idx, col_idx)
					cell_type = xl_sheet.cell_type(row_idx, col_idx)
					valuestr = xl_sheet.cell_value(row_idx, col_idx)

					backChar = delimiter
					#Reach the end of columns
					if col_idx == xl_sheet.ncols - 1:
						backChar = ''
					
					# Check if the current cell is blank
					if cell_type == xlrd.XL_CELL_EMPTY:
						empty_count = empty_count + 1
					
					# Check if the row is blank
					if empty_count == xl_sheet.ncols:
						empty_count_flag = True

					if empty_count_flag:
						break

					#Check if a number is same as it's decimal
					if cell_type==xlrd.XL_CELL_NUMBER:
						if(valuestr==int(valuestr)):
							fp.write(str(int(valuestr))+backChar)
							continue

					#Date handling
					if cell_type == xlrd.XL_CELL_DATE:
						try:
							y, m, d, h, i, s = xlrd.xldate_as_tuple(valuestr, wb.datemode)
							fp.write(str(d) + '/' + str(m) + '/' + str(y) + backChar)
						except:
							fp.write(str(valuestr) + backChar)
					else:
						if row_idx == 0 and cell_type == xlrd.XL_CELL_NUMBER:
							fp.write(str(int(valuestr)) + backChar)
						else:
							fp.write(str(valuestr) + backChar)
				if empty_count_flag:
					break
				fp.write('\n')
			fp.close()
		except Exception as processing_exception:
			print(processing_exception)
			return 'Error occured while processing the file'
	return 'Success'
import sys
import os
import shutil
from converters import xlsx_to_custom_delimited_file

#Command line arguments
try:
	input_dir=sys.argv[1]
	output_dir=sys.argv[2]
	extension_to_save=sys.argv[3]
	delimiter=bytes(sys.argv[4], "utf-8").decode("unicode_escape")
except:
	print('Usage: python3 converter_main.py <argument1> <argument2> <argument3> <argument4> <argument5>\n\
			<argument1>: Input directory\n\
			<argument2>: Output Directory\n\
			<argument3>: File extension to save with\n\
			<argument4>: Delimiter\n\
			<argument5> [Optional]: Row start index [Zero by default]\n\
	')
	sys.exit()

#Optional arguments
try:
	row_start_index=int(sys.argv[5])
except IndexError:
	row_start_index=0

if os.path.exists(output_dir):
	shutil.rmtree(output_dir)
os.makedirs(output_dir)

status=xlsx_to_custom_delimited_file(input_dir,output_dir,extension_to_save,delimiter,row_start_index)
print(status)
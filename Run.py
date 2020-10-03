import excel_udf
import os
import re
#from progressbar import ProgressBar

	#pbar = ProgressBar()

folder_path = r'C:\\Users\\dalrympm\\Documents\\Coding\\Files\\TMO_Sheets'
filenames = os.listdir(folder_path)
#print(filenames)
num_files = list(range(len(filenames)))
count = 0 

for i in num_files:
	if re.search('.docx', filenames[i]):
		count += 1
		print(count,filenames[i])
		excel_udf.excel(filenames[i])

print("DONE!!")
print("ALL TMO DOCX HAVE BEEN PARSED INTO EXCEL SPREAD SHEETS")




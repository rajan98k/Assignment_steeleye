#!/usr/bin/env python

import json
import xlrd
import sys


class Assignment:
	
	def __init__(self,excel_file,json_file):
		self.file_name = excel_file
		self.dump_file = json_file
		self.list_of_dict = []
		self.reading_rows = dict()		

	def file_read(self):
		workbook = xlrd.open_workbook(self.file_name)
		sheet = workbook.sheet_by_index(1)
		#initializing key values
		reload(sys)
		sys.setdefaultencoding('utf8')
		keys_list = sheet.row_values(0)		
		for i in range(1,sheet.nrows):
			temp_dict = {}			
			data_field_list = sheet.row_values(i)
			for j in range(len(keys_list)):
				temp_dict[keys_list[j]] = str(data_field_list[j])
			self.list_of_dict.append(temp_dict)




	def file_dump(self):
		#now dumping the dictionaries containing the excel rows line by line in json file
		fp = open(self.dump_file,'w+')
		for json_obj in self.list_of_dict:
			fp.write('%s\n'%(json.dumps(json_obj)))
		fp.close()






	def run_main(self):
		self.file_read()
		self.file_dump()


#please put this line inside the lambda function to make this along with the above code.
Assignment('ISO10383_MIC.xls','result.json').run_main()

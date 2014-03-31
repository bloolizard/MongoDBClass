# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the delimiter.
# An example output can be seen in the "example.csv" file.
import xlrd
import os
import csv
from zipfile import ZipFile
datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"
outfile2 = "test.csv"

def open_zip(datafile):
    with ZipFile('2013_ERCOT_Hourly_Load_Data.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)

    vals = {}
    for i in range(1,9):
    	title = sheet.cell_value(0, i)
    	column_data = sheet.col_values(i, start_rowx=1, end_rowx=8761)
    	max_value = max(column_data)
    	max_value_index = column_data.index(max_value)+1
    	date = xlrd.xldate_as_tuple(sheet.cell_value(max_value_index, 0),0)
    	vals[title] = [date, max_value]
    data = vals
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    return data

def save_file(data, filename):
	values = {}
	for keys in data:
		title = keys
		date = data[keys][0]
		max_load = data[keys][1]
		values[keys] = {
				'Year': date[0],
				'Month':date[1],
				'Day': date[2],
				'Hour': date[3],
				'Max Load': max_load
			}
	with open(filename,'wb') as csvfile:
		datawriter = csv.writer(csvfile, delimiter='|')
		datawriter.writerow(['Station','Year','Month','Day','Hour', 'Max Load'])
		for keys in values:
			datawriter.writerow([keys, 
				values[keys]['Year'],
				values[keys]['Month'],
				values[keys]['Day'],
				values[keys]['Hour'],
				values[keys]['Max Load']]) 
	return data

def test():
	open_zip(datafile)
   	data = parse_file(datafile)
   	save_file(data, outfile)
   	ans = {'FAR_WEST': {'Max Load': "2281.2722140000024", 'Year': "2013", "Month": "6", "Day": "26", "Hour": "17"}}
    
   	fields = ["Year", "Month", "Day", "Hour", "Max Load"]
   	with open(outfile) as of:
   		csvfile = csv.DictReader(of, delimiter="|")
   		for line in csvfile:
   			s = line["Station"]
   			if s == 'FAR_WEST':
   				for field in fields:
   					assert ans[s][field] == line[field]

        
test()
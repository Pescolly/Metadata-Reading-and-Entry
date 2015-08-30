#!/usr/bin/python

#########
#########					messing around with xlwt insert abilities
#########					
#########					March 2014 armen karamian
#########					!!!!!!!!!!!!!!! MUST WRITE TO A .XLS FILE !!!!!!!!!!!!!!!!!!!



import os.path, sys, string, time, re
import xlrd.sheet as sheet, xlrd
import NbcFileClass as nbc
import datetime
from xlwt import *
from copy import deepcopy
from operator import itemgetter


XLSFILE = '/Volumes/fs3/encoding/AssetManagement/NBCU_Features_Series_MASTER.xls'
NBCPATH = '/Volumes/fs3encoding/AssetManagement/nbc_incoming/'
OUTPUT = '/Volumes/fs3/encoding/AssetManagement/NBCU_Features_Series_MASTER.xls'

TV_FLAG = "television"
FEATURE_FLAG = "features"
TRAILER_FLAG = "trailers"
MISC_FLAG = "artwork and scc"


def getFilenames(nbcpath):																								#return list of files in incoming folder
	returnList = []
	
	for file in os.listdir(nbcpath):
		fullpath = NBCPATH+file
		if (file.startswith('.') or file.startswith('#')):
			continue
		if (os.path.isfile(fullpath)):
			fileStat = os.stat(fullpath)
			fileobj = nbc.NBCFile(file)
			returnList.append(fileobj)
				
	return returnList

def getFiletype(file):																					#place file into incoming files dictionary with file type (trailer, tv feature) as value
	DELIMITER = '_'																										#delimiter for splitting file name
	TV_PATTERN = '.*(television).*'																						#re pattern for television
	TV_PATTERN2 = '.*(5994).*'
	FEATURE_PATTERN = '.*(ftr).*'																							#re pattern for features
	TRAILER_PATTERN = '.*(trl).*'																							#re patterns for trailers
	TRAILER_PATTERN2 = '.*(trailer).*'
	JPEG_PATTERN = '.*(\.jpeg|\.jpg)'
	CAP_PATTERN =  '.*(\.scc|\.cap)'
	
	filename = file.filename
	splitfile = string.split(filename, DELIMITER)																			#Split the file name into a list

#	print '\n'
	for i in splitfile:																									#iterate thru list until feature or television file is confirmed
		tv_match = re.match(TV_PATTERN, string.lower(i))
		tv_match2 = re.match(TV_PATTERN2, string.lower(i))
		feature_match = re.match(FEATURE_PATTERN, string.lower(i))
		trailer_match = re.match(TRAILER_PATTERN, string.lower(i))
		trailer_match2 = re.match(TRAILER_PATTERN2, string.lower(i))
		misc_match = re.match(JPEG_PATTERN, string.lower(i))
		misc_match2 = re.match(CAP_PATTERN, string.lower(i))
#		print i
		if (tv_match or tv_match2):
#			print "tv match"
			return TV_FLAG																				#insert item into global dictionary
			
		
		if (trailer_match or trailer_match2):
#			print "Trailer"
			return FEATURE_FLAG																		#insert item into global dictionary
			
		
		if (feature_match):
#			print "feature"
			return FEATURE_FLAG																		#insert item into global dictionary
			
			
		if (misc_match or misc_match2):
#					print "misc"
			return MISC_FLAG
		
	else:
		print "\nCannot determine file type:", file.filename
		
		
def breakdownXLS(excelWorkbook):																							###open workbook and create dictionary of the spreadsheets within
	STARTCOL = 0
	ENDCOL = 10
	
	
	xlsbook = xlrd.open_workbook(excelWorkbook)

#	print "\n"+"-"*50
#	print "Checking workbook:",os.path.split(excelWorkbook)[1]	
#	print "-"*50


	if (xlsbook.nsheets != 0):																							#check if workbook has spreadsheets	
		worksheetDic = dict()																							#create list of sheets in workbook		
		sheetList = xlsbook.sheets()

		for sheet in sheetList:		
			#print sheet
			#if (sheet.nrows == 0): 																						#check if sheet has rows in it. skip if 0
			#	continue
#			print "Sheet name:",sheet.name, "Rows:",sheet.nrows
			
			rowList = []
			numberOfRows = sheet.nrows																					#get number of rows in each spread sheet
			counter = 0
			while counter < numberOfRows:
				currentRow = sheet.row_slice(counter, start_colx=STARTCOL, end_colx=ENDCOL)								#get row with only 10 columns
				if (currentRow[0].value != ''):
#					print currentRow[0].value
					rowList.append(currentRow)																				#add current row to rowlist
					counter += 1
				else:
					counter += 1
			worksheetDic[sheet.name] = rowList																			#put list of rows in dictionary with worksheet name as the key

		return worksheetDic
	
	else:
		print "Workbook has no spreadsheets in it."
		sys.exit()		

def insertNewRows(incomingFiles, workbook):																					#add new files (in rows) into workbook
	TITLE = 0																											#column name constants
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 2
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6
	DELIMITER = '_'
	
	FILENAME = 0																										#dictionary pair values
	
	previousAsset = ''																									#previous asset to compare incoming file to
	for file in incomingFiles:				
		try:
			filename = file.filename
			filetype = getFiletype(file)
	#		print filetype	
			worksheet = workbook[filetype] 
			newrow = deepcopy(worksheet[0])																					#create new row using deepcopy
			print "inserting",filename,"into",filetype
			for cell in newrow:
				cell.value = ""
			date = datetime.datetime.today()
			newrow[TITLE].value = filename
			newrow[DATE].value = date.isoformat()[:10]
			worksheet.append(newrow)
			workbook[filetype] = worksheet
			
		except KeyError, e:
			print file.filename, e
					
	return workbook
		
				
def updateWorkbook(excelfile, workbook):
	TITLE = 0																											#column name constants
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 2

	global OUTPUT
	wb = Workbook()
	for sheet in workbook.keys():		#get sheet from open workbook
		ws = wb.add_sheet(sheet)		#create a new worksheet with the same name
		ws.col(TITLE).width = 0x5A00
		format = easyxf()
		fnt = Font()
		fnt.name = 'Arial'
		format.font = fnt
		rows = workbook[sheet]			#get all rows from the original worksheet
		if (rows == None):			
			continue
		for row in rows:				#copy all the rows to the new document
			rowindex = rows.index(row)
			for cell in row:				#copy each column from each row to new worksheet
				col = row.index(cell)
				if (col == CTRLNUMBER):
					format.num_format_str = "########"
				if (col == DATE):
					format.num_format_str = "yyyy/mm/dd"
				ws.write(rowindex, col, cell.value, format)
	
	wb.save(OUTPUT)						#save new .xls file
	
	
#def insertControlNumber():
	

if __name__ == "__main__":	


	incomingFiles = getFilenames(NBCPATH)																					#get list of new files
	workbookDic = breakdownXLS(XLSFILE)																						#get dictionary of items in the excelWorkbook
#	for file in incomingFiles:
#		workbookDic[file] = getFiletype(file)
	
	workbookDic = insertNewRows(incomingFiles, workbookDic) 
	updateWorkbook(XLSFILE, workbookDic)	



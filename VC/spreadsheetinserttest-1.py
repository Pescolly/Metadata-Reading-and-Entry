#!/usr/bin/python

#########
#########					messing around with xlwt insert abilities
#########					
#########					March 2014 armen karamian
#########					!!!!!!!!!!!!!!! MUST WRITE TO A .XLS FILE !!!!!!!!!!!!!!!!!!!



import os.path, sys, string, time, re
import xlrd.sheet as sheet, xlrd
import datetime
from xlwt import *
from copy import deepcopy
from operator import itemgetter


XLSFILE = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features_Series_MASTER.xlsx'
NBCPATH = '/Volumes/fs3/encoding/AssetManagement/nbc_incoming/'
OUTPUT = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/TestSpreadSheetOUTPUT.xls'

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
			returnList.append(file)
				
	return returnList

def getFiletype(file):																					#place file into incoming files dictionary with file type (trailer, tv feature) as value
	DELIMITER = '_'																										#delimiter for splitting file name
	TV_PATTERN = '.*(television).*'																						#re pattern for television
	FEATURE_PATTERN = '.*(ftr).*'																							#re pattern for features
	TRAILER_PATTERN = '.*(trl).*'																							#re patterns for trailers
	TRAILER_PATTERN2 = '.*(trailer).*'
	JPEG_PATTERN = '.*(\.jp[e,g]|g)'
	CAP_PATTERN =  '.*(\.[scc|cap])'
	
	splitfile = string.split(file, DELIMITER)																			#Split the file name into a list

	
	for i in splitfile:																									#iterate thru list until feature or television file is confirmed
		tv_match = re.match(TV_PATTERN, string.lower(i))
		feature_match = re.match(FEATURE_PATTERN, string.lower(i))
		trailer_match = re.match(TRAILER_PATTERN, string.lower(i))
		trailer_match2 = re.match(TRAILER_PATTERN2, string.lower(i))
		misc_match = re.match(JPEG_PATTERN, string.lower(i))
		misc_match2 = re.match(CAP_PATTERN, string.lower(i))
		if (tv_match):
			return TV_FLAG																				#insert item into global dictionary
			
		
		if (trailer_match or trailer_match2):
			return FEATURE_FLAG																		#insert item into global dictionary
			
		
		if (feature_match):
			return FEATURE_FLAG																		#insert item into global dictionary
			
			
		if (misc_match or misc_match2):
			return MISC_FLAG
		
	else:
		print "\nCannot determine file type:", file
		
		
def breakdownXLS(excelWorkbook):																							###open workbook and create dictionary of the spreadsheets within
	STARTCOL = 0
	ENDCOL = 10
	
	
	xlsbook = xlrd.open_workbook(excelWorkbook)

	print "\n"+"-"*50
	print "Checking workbook:",os.path.split(excelWorkbook)[1]	
	print "-"*50


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
	DATE = 3
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6
	DELIMITER = '_'
	
	FILENAME = 0																										#dictionary pair values
	
#	dummyInsert = 'War3The_DA000460957_05456_FTR_Theatrical_2398_16x9LB_185_CSP_HD_ProResHQ_220M.mov'					#test filename
	previousAsset = ''																									#previous asset to compare incoming file to
#	print "Inserting:", incomingFiles
	for file in incomingFiles:																				
		filename = file
		filetype = getFiletype(file)
		print filetype	
		worksheet = workbook[filetype] 
		newrow = deepcopy(worksheet[0])																					#create new row using deepcopy
		print "inserting",filename,"into",filetype
		for cell in newrow:
			cell.value = ""
		date = datetime.datetime.today()
		newrow[TITLE].value = filename
		newrow[DATE].value = str(date.year)+"-"+str(date.month)+"-"+str(date.day)
		worksheet.append(newrow)
		workbook[filetype] = worksheet
					
#	for sheet in workbook.keys():																						#incomplete sorting routine
#		print sheet
#		print workbook[sheet]
#		list_to_sort = workbook[sheet]
#		list_to_sort.sort(key=itemgetter('text'))
#		for i in list_to_sort:
#			print i
	
	return workbook
		
				
def updateWorkbook(excelfile, workbook):
	global OUTPUT
	wb = Workbook()
	for sheet in workbook.keys():
		ws = wb.add_sheet(sheet)
		rows = workbook[sheet]
		if (rows == None):
			continue
		for row in rows:
			
			rowindex = rows.index(row)
			
			for col in row:
				cell = row.index(col)
				ws.write(rowindex, cell, col.value)
				
	
	wb.save(OUTPUT)

if __name__ == "__main__":	


	incomingFiles = getFilenames(NBCPATH)																					#get list of new files
	workbookDic = breakdownXLS(XLSFILE)																						#get dictionary of items in the excelWorkbook
#	for file in incomingFiles:
#		workbookDic[file] = getFiletype(file)
	
	workbookDic = insertNewRows(incomingFiles, workbookDic) 
	updateWorkbook(XLSFILE, workbookDic)	



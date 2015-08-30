#!/usr/bin/python

#########
#########					messing around with xlwt insert abilities
#########					
#########					March 2014 armen karamian
#########



import os.path, sys, string, time, re
import xlrd.sheet as sheet, xlrd
from xlwt import *
from copy import deepcopy

XLSFILE = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/TestSpreadSheet.xlsx'
NBCPATH = "/Volumes/fs3/encoding/AssetManagement/nbc_incoming/"

television = []
features = []
trailers = []

def getFilenames(nbcpath):																								#return list of files in incoming folder
#	SECONDSINDAY = 60*60*24																								#decommissioned age checking 
	returnList = []

	for file in os.listdir(nbcpath):
		fullpath = NBCPATH+file
		if (file.startswith('.') or file.startswith('#')):
			continue
		if (os.path.isfile(fullpath)):
			fileStat = os.stat(fullpath)
			#print file
#			fileAge = (time.time() - fileStat.st_mtime)
#			if (fileAge < (SECONDSINDAY)):
#				print fileAge
#				print file
			returnList.append(file)
				
	return returnList

def getFiletype(file):
	DELIMITER = '_'																										#delimiter for splitting file name
	TV_FLAG = '.*(television).*'																						#re pattern for television
	FEATURE_FLAG = '.*(ftr).*'																							#re pattern for features
	TRAILER_FLAG = '.*(trl).*'																							#re patterns for trailers
	TRAILER_FLAG2 = '.*(trailer).*'
	
	splitfile = string.split(file, DELIMITER)																			#Split the file name into a list
	
	for i in splitfile:																									#iterate thru list until feature or television file is confirmed
		tv_match = re.match(TV_FLAG, string.lower(i))
		feature_match = re.match(FEATURE_FLAG, string.lower(i))
		trailer_match = re.match(TRAILER_FLAG, string.lower(i))
		trailer_match2 = re.match(TRAILER_FLAG2, string.lower(i))
		if (tv_match):
			television.append(file)
			break
		
		if (trailer_match or trailer_match2):
			trailers.append(file)
			break
		
		if (feature_match):
			features.append(file)
			break
		
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
			
			if (sheet.nrows == 0): 																						#check if sheet has rows in it. skip if 0
				continue
#			print "Sheet name:",sheet.name, "Rows:",sheet.nrows
			
			rowList = []
			numberOfRows = sheet.nrows																					#get number of rows in each spread sheet
			counter = 0
			while counter < numberOfRows:
				currentRow = sheet.row_slice(counter, start_colx=STARTCOL, end_colx=ENDCOL)								#get row with only 10 columns
				rowList.append(currentRow)																				#add current row to rowlist
				counter += 1
			
			worksheetDic[sheet.name] = rowList																			#put list of rows in dictionary with worksheet name as the key
		return worksheetDic
	else:
		print "Workbook has no spreadsheets in it."
		sys.exit()		

def insertRow(filename, workbook):
	TITLE = 0																											#column name constants
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 3
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6
	DELIMITER = '_'
	
	dummyInsert = 'War3The_DA000460957_05456_FTR_Theatrical_2398_16x9LB_185_CSP_HD_ProResHQ_220M.mov'					#test filename
	previousAsset = ''																									#previous asset to compare incoming file to
	
	#filename = dummyInsert ############# REMOVE ##################	
	for worksheetKey in workbook.keys():																				#iterate thru worksheets in workbook
#		print worksheetKey
		worksheet = workbook[worksheetKey] 
		for row in worksheet:																							#iterate thru rows in worksheet
			asset = row[TITLE].value
			if (asset == ""):																							#skip empty cells
				continue

			assetStrings = string.split(asset, DELIMITER)																#split title into chunks and pull title and da number
			filenameStrings = string.split(filename, DELIMITER)															
#			originalAssetTitle = assetStrings[TITLE]																	#title of currently selected cell	
			assetTitle = string.lower(assetStrings[TITLE])																#title of currently selected cell in lowercase for comparison
#			originalFileTitle = filenameStrings[TITLE]
			fileTitle = string.lower(filenameStrings[TITLE])															#title of file to insert in lowercase
			


			
#			print assetTitle, fileTitle
#				print previousAsset
			if (fileTitle < assetTitle):																				#if title names are the same look for insertion spot
				if (fileTitle > previousAsset):																			#print  
					print "inserting asset", filename,"into",worksheetKey
					rowIndex = (worksheet.index(row) + 1)
					newRow = deepcopy(row)																				#all other copy methods failed, using deepcopy to create a new row

					for i in newRow:																					#blank out values in new row
						i.value = ''
					newRow[TITLE].value = unicode(filename)																#insert new row with filename in TITLE field
					worksheet.insert(rowIndex, newRow)
					workbook[worksheetKey] = worksheet																	#update workbook
					return 																								#move onto next incoming file
					
			previousAsset = assetTitle

		
				
def updateWorkbook(excelFile, ):
	pass
			


incomingFiles = getFilenames(NBCPATH)																				#get list of new files
workbookDic = breakdownXLS(XLSFILE)																			#get dictionary of items in the excelWorkbook
for file in incomingFiles:
	getFiletype(file)
	
print "Telly"	
for i in television:
	print i
	
print "\nbig screen shti"
for j in features:
	print j
	
print "\ntrailers"
for k in trailers:
	print k		
for file in incomingFiles:
	insertRow(file, workbookDic) 

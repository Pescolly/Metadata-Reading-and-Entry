#!/usr/bin/python

#########
#########					nbc-universal directory parser and spreadsheet checker
#########					parses designated spreadsheet and nbcincoming path
#########					March 2014 armen karamian
#########



import os.path, sys, string, time, re
import xlrd.sheet as sheet, xlrd

SPREADSHEET = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features-Series.xlsx'
NBCPATH = "/Volumes/fs3/encoding/AssetManagement/nbc_incoming/"



def getFilenames(nbcpath):											#return list of files in incoming folder
	SECONDSINDAY = 60*60*24
	returnList = []

	for file in os.listdir(nbcpath):
		fullpath = NBCPATH+file
		if (file.startswith('.') or file.startswith('#')):
			continue
		if (os.path.isfile(fullpath)):
			fileStat = os.stat(fullpath)
			#print file
			fileAge = (time.time() - fileStat.st_mtime)
			if (fileAge < (SECONDSINDAY)):
#				print fileAge
#				print file
				returnList.append(file)
				
	return returnList

def breakdownXLS(spreadsheet):									###open workboot and create dictionary of the spreadsheets
	
	xlsbook = xlrd.open_workbook(spreadsheet)

	print "\n"+"-"*50
	print "Checking workbook:",os.path.split(spreadsheet)[1]	
	print "-"*50


	if (xlsbook.nsheets != 0):					#check if workbook has spreadsheets	
		worksheetDic = dict()							#create list of sheets in workbook		
		sheetList = xlsbook.sheets()

		for sheet in sheetList:		
			
			if (sheet.nrows == 0): 				#check if sheet has rows in it. skip if 0
				continue
#			print "Sheet name:",sheet.name, "Rows:",sheet.nrows
			
			rowList = []
			numberOfRows = sheet.nrows					#get number of rows in sheet
			counter = 0
			while counter < numberOfRows:					#create a list of rows
				rowList.append(sheet.row(counter))			
				counter += 1
			
			worksheetDic[sheet.name] = rowList
		return worksheetDic
	else:
		print "Workbook has no spreadsheets in it."
		sys.exit()		

def findMatches(incomingFiles,worksheetDic):				###take list and dictionary and output matches
	TITLE = 0	
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 3
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6

	for file in incomingFiles:								#look for new found files
		print "+"*50
		print "Looking for file:",file
#		print os.path.splitext(file)[1]
		for sheet in worksheetDic:							#get list of keys in dictionary
#			print "Sheet:",sheet
			rowList = []
			rowList = worksheetDic[sheet]					#call each rowlist by name
			for row in rowList:				
				titleCell = row[TITLE]
				controlCell = row[CTRLNUMBER]
				controlNumber = ''
				title = titleCell.value
							
				daPattern = re.compile('.+_(DA.........)_.*')				#re pattern to get DA number
				cellDANumber = re.findall(daPattern, title)
				fileDANumber = re.findall(daPattern, file)
				
				prodPattern = re.compile('.+_(\w\w\w\d\d)_.*')				#re pattern to get production number
				cellprodNumber = re.findall(prodPattern, title)
				fileprodNumber = re.findall(prodPattern, file)


				if (len(cellDANumber) != 0 and len(fileDANumber) != 0):	 				
					if (cellDANumber[0] == fileDANumber[0]):							#attempt DA number match
						try:
							controlNumber = int(controlCell.value)
						except ValueError, v:
							controlNumber = "Not Assigned"

						print "\nDA Number match"
						print "Found", title,"at row",rowList.index(row)+1,"in worksheet:",sheet	#add one to row index because of header row
						print "Control Number:", controlNumber	#output information if file matches in database
						continue														#move onto next row
				
				if (len(cellprodNumber) != 0 and len(cellprodNumber) != 0):
					if (cellprodNumber[0] == fileprodNumber[0]):					#attempt production number match
						try:
							controlNumber = int(controlCell.value)
						except ValueError, v:
							controlNumber = "Not Assigned"

						print "\nProduction number match"
						print "Found", title,"at row",rowList.index(row)+1,"in worksheet:",sheet	#add one to row index because of header row
						print "Control Number:", controlNumber	#output information if file matches in database
						continue												#move onto next row



if __name__ == "__main__":

	incomingFiles = getFilenames(NBCPATH)							#get list of new files
	worksheetDic = breakdownXLS(SPREADSHEET)							#get dictionary of items in the spreadsheet
	findMatches(incomingFiles, worksheetDic)				#find any matches
		
		
	print "\n"+"-"*50
	print "Done..."	
	print "-"*50
	sys.exit()

		
				
		
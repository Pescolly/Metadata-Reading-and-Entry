#!/usr/bin/python

#########
#########					nbc-universal directory parser and spreadsheet checker
#########					parses designated spreadsheet and nbcincoming path
#########					March 2014 armen karamian
#########



import os.path, sys, string, time, re, emailer, datetime, smtplib
import NbcFileClass as nbc
import xlrd.sheet as sheet, xlrd


SPREADSHEET = '/Volumes/fs3/encoding/AssetManagement/NBCU_Features_Series_MASTER.xls'

#SPREADSHEET = raw_input("Drag spreadsheet here: ").rstrip()


NBCPATH = "/Volumes/fs3/encoding/AssetManagement/nbc_incoming/"

MESSAGESTRING = ''
duplicates = 0																											#duplicate counter

def createMov(filepath):
	return mov

def enterFilesFromObject(mov_object):
	pass

def enterFilename(filename, spreadsheet):
	pass
	
def sendEmail(messageString):
	time = datetime.datetime.today()
	timestamp = str(time.hour)+' '+str(time.minute)+' '+str(time.second) 

	SERVER = 'owa.modern.mvfinc.com'
	SENDLIST = ['akaramian@mvf.com']
	RECEIVELIST = ['akaramian@mvf.com']
	SENDER_NAME = "Asset Management"
	MESSAGE_TEXT = messageString
	SUBJECT = 'NBC Incoming List'
	try:
		MESSAGE = "From: %s\nTo: %s\nSubject: %s\n\n%s" % (SENDER_NAME, RECEIVELIST, SUBJECT, MESSAGE_TEXT)
		smtpObj = smtplib.SMTP(SERVER, 25)
		smtpObj.sendmail(SENDLIST, RECEIVELIST, MESSAGE)
		smtpObj.quit()	
		
	except Exception, e:
		print e

def getFilenames(nbcpath):																								#return list of files in incoming folder
	returnList = []
	print nbcpath
	for file in os.listdir(nbcpath):
		print file
		fileobj = nbc.NBCFile(file)
		fullpath = nbcpath+file
		if (file.startswith('.') or file.startswith('#')):
			continue
		if (os.path.isfile(fullpath)):
			fileStat = os.stat(fullpath)
			returnList.append(fileobj)
#			print file	
	return returnList

def breakdownXLS(spreadsheet):																							###open workbook and create dictionary of the spreadsheets within
	STARTCOL = 0
	ENDCOL = 10
	
	
	xlsbook = xlrd.open_workbook(spreadsheet)

	print "\n"+"-"*50
	print "Checking workbook:",os.path.split(spreadsheet)[1]	
	print "-"*50


	if (xlsbook.nsheets != 0):																							#check if workbook has spreadsheets	
		worksheetDic = dict()																							#create list of sheets in workbook		
		sheetList = xlsbook.sheets()

		for sheet in sheetList:		
			
			if (sheet.nrows == 0): 																						#check if sheet has rows in it. skip if 0
				continue
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

def findMatches(incomingFiles,worksheetDic):																			###take list and dictionary and output matches
	TITLE = 0	
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 2
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6
	returnList = []
	global duplicates#@@@@@@@@@@@@@@ DO NOT DELETE THIS VARIABLE OR ELSE EVERYTHING WILL GO TO SHIT

	for file in incomingFiles:																							#look for new found files
		

		for sheet in worksheetDic:																						#get list of keys in worksheet dictionary
			rowList = []
			rowList = worksheetDic[sheet]																				#call each rowlist by the name of the current sheet
			for row in rowList:				
				try:
					titleCell = row[TITLE]
					controlCell = row[CTRLNUMBER]
					dateCell = row[DATE]
					controlNumber = ''
					title = titleCell.value
					if (string.lower(file.filename) == string.lower(title)):														#attempt file name match
						try:
							controlNumber = int(controlCell.value)
						except ValueError, v:
							controlNumber = "Not Assigned"

						try:
							entryDate = int(dateCell.value)
						except ValueError, v:
							entryDate = "Not Assigned"
						
						duplicates += 1
						file.controlNumber = controlNumber
						file.entryDate = entryDate
						returnList.append(file)
						continue		
								
					daPattern = re.compile('.+_(DA.........)_.*')														#re pattern to get DA number
					cellDANumber = re.findall(daPattern, title)
					fileDANumber = re.findall(daPattern, file.filename)
					
					prodPattern = re.compile('.+_(\w\w\w\d\d)_.*')														#re pattern to get production number
					cellprodNumber = re.findall(prodPattern, title)
					fileprodNumber = re.findall(prodPattern, file.filename)
	
	
					if (len(cellDANumber) != 0 and len(fileDANumber) != 0):	 				
						if (cellDANumber[0] == fileDANumber[0]):														#attempt DA number match
							try:
								controlNumber = int(controlCell.value)
							except ValueError, v:
								controlNumber = "Not Assigned"

							try:
								entryDate = int(dateCell.value)
							except ValueError, v:
								entryDate = "Not Assigned"
							
							duplicates += 1
							file.controlNumber = controlNumber
							file.entryDate = entryDate
							returnList.append(file)

#@							print "\nDA Number match"
#							print "Found", title,"at row",rowList.index(row)+1,"in worksheet:",sheet					#add one to row index because of header row
#							print "Control Number:", controlNumber														#output information if file matches in database
							continue																					#move onto next row
					
				
				except NameError, e:
#					print e
					continue
				except IndexError, e:
#					print e
					continue
#	print returnList
	return returnList				


if __name__ == "__main__":

	incomingFiles = getFilenames(NBCPATH)																				#get list of new files
#	worksheetDic = breakdownXLS(SPREADSHEET)																			#get dictionary of items in the spreadsheet
#	list = findMatches(incomingFiles, worksheetDic)																			#find any matches and print to terminal
#	sendEmail(MESSAGESTRING)	
#	print list
	print "\n"+"-"*50
	if (duplicates == 0):
		print "No duplicates found"
	elif (duplicates > 0):
		print "Found",duplicates,"duplicates"
	print "Done..."	
	print "-"*50
	sys.exit()

		
				
		
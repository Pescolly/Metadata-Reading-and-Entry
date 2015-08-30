#!/usr/bin/python

#########
#########					nbc-universal directory parser and EXCELFILE checker
#########					parses designated EXCELFILE and nbcincoming path
#########					March 2014 armen karamian
#########



import os.path, sys, string, time, re, emailer, datetime, smtplib
import xlrd.sheet as sheet, xlrd
import xlrd.xldate as xldate
import xlrd.formatting as format
from datetime import date, datetime, timedelta


EXCELFILE = '/Volumes/fs3/encoding/AssetManagement/NBCU_Features_Series_MASTER.xls'								
#EXCELFILE = raw_input("Drag EXCELFILE here: ").rstrip() 															#uncomment to enable drag and drop EXCELFILE entry

NBCPATH = "/Volumes/fs3//encoding/AssetManagement/nbc_incoming/"														#path for nbc files

MESSAGESTRING = ''
duplicates = 0																											#duplicate counter

def sendEmail(messageString):																						#takes the string created and sends an email
	time = datetime.today()
	timestamp = str(time.hour)+' '+str(time.minute)+' '+str(time.second) 

	SERVER = 'owa.modern.mvfinc.com'																				
	SENDLIST = ['akaramian@mvf.com']
	RECEIVELIST = ['akaramian@mvf.com','iblanch@mvf.com', 'rrangel@mvf.com', 'jsheridan@mvf.com', 'cholsey@mvf.com']
	SENDER_NAME = "Asset Management"
	MESSAGE_TEXT = messageString
	SUBJECT = 'NBC Incoming List'
	try:
		MESSAGE = "From: %s\nTo: %s\nSubject: %s\n\n%s" % (SENDER_NAME, RECEIVELIST, SUBJECT, MESSAGE_TEXT)
		smtpObj = smtplib.SMTP(SERVER, 25)																					#connect to SMTP port of email server
		smtpObj.sendmail(SENDLIST, RECEIVELIST, MESSAGE)
		smtpObj.quit()	
		
	except Exception, e:
		print e

def getFilenames(nbcpath):																								#return list of files in incoming folder
	returnList = []																										#skips hidden and incomplete files

	for file in os.listdir(nbcpath):
		fullpath = NBCPATH+file
		if(os.path.isdir(fullpath)):
			continue
		if (file.startswith('.') or file.startswith('#')):
			continue
		if (os.path.isfile(fullpath)):
			fileStat = os.stat(fullpath)
			returnList.append(file)
				
	return returnList

def breakdownXLS(EXCELFILE):																							###open workbook and create dictionary of the EXCELFILEs within
	STARTCOL = 0																										#column zero is the title field
	ENDCOL = 10																											#only pull in 10 columns
	
	xlsbook = xlrd.open_workbook(EXCELFILE)

	print "\n"+"-"*50
	print "Checking workbook:",os.path.split(EXCELFILE)[1]	
	print "-"*50

	if (xlsbook.nsheets != 0):																							#check if workbook has EXCELFILEs	
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

def dateConversion(datestring):																							#converts from excel time format (######) to human readable time (yyyy/mm/dd)
	from1900 = datetime(1970,1,1) - datetime(1900,1,1) + timedelta(days=2)
	return date.fromtimestamp(int(datestring)*86400) - from1900


def findMatches(incomingFiles,worksheetDic):																			###take list and dictionary and output matches
	TITLE = 0																											#set constants for each column
	CTRLNUMBER = 1	
	ENTRY = 2 
	DATE = 3
	ARCHIVED = 4
	MD5 = 5 
	NOTES = 6
	global MESSAGESTRING																								#summon global message string for EMail
	global duplicates																									#summon global message string for duplicates		

	for file in incomingFiles:																							#look for new found files
		print "+"*50
		print "Looking for file:",file
		MESSAGESTRING += str("+"*50+'\n')
		MESSAGESTRING += "Looking for file:"+str(file)+'\n'
		for sheet in worksheetDic:																						#get list of keys in worksheet dictionary
			rowList = []
			rowList = worksheetDic[sheet]																				#call each rowlist by the name of the current sheet
			for row in rowList:				
				try:
					titleCell = row[TITLE]
					controlCell = row[CTRLNUMBER]
					entrycell = row[ENTRY]

					controlNumber = ''
					title = titleCell.value
	
					if (string.lower(file) == string.lower(title)):														#attempt file name match
						try:
							print "\n"+string.lower(file),string.lower(title)
							controlNumber = int(controlCell.value)
	
						except ValueError, v:
							controlNumber = "Not Assigned"
	
						try:
							entrydate = dateConversion(entrycell.value)
						except ValueError, de:
							entrydate = "Not Entered"
						
						duplicates += 1
						MESSAGESTRING += str("Found ")+str(title)+" at row "+str(rowList.index(row)+1)+" in worksheet: "+str(sheet)+'\n'
						MESSAGESTRING += "Control Number: "+str(controlNumber)+'\n'
						MESSAGESTRING += "Entry Date: "+str(entrydate)+'\n'
						MESSAGESTRING += str("+"*50+'\n')

						print "\nFile name match"
						print "Found ", title,"at row ",rowList.index(row)+1," in worksheet:",sheet						#add one to row index because of header row
						print "Control Number:", controlNumber		
						print "Entry Date:", entrydate																	#output information if file matches in database
						print "+"*50+'\n'
						
						continue		
								
					daPattern = re.compile('.+_(da.........)_.*')														#re pattern to get DA number
					cellDANumber = re.findall(daPattern, string.lower(title))
					fileDANumber = re.findall(daPattern, string.lower(file))
						
					if (len(cellDANumber) != 0 and len(fileDANumber) != 0):	 				
						if (cellDANumber[0] == fileDANumber[0]):														#attempt DA number match
							try:
								controlNumber = int(controlCell.value)
							except ValueError, v:
								controlNumber = "Not Assigned"

							try:
								entrydate = dateConversion(entrycell.value)							
							except ValueError, de:
								entrydate = "Not Entered"
							
							duplicates += 1																				#add one to duplicate count
							MESSAGESTRING += str("Found ")+str(title)+" at row "+str(rowList.index(row)+1)+" in worksheet: "+str(sheet)+'\n'
							MESSAGESTRING += "Control Number: "+str(controlNumber)+'\n'
							MESSAGESTRING += "Entry Date: "+str(entrydate)+'\n'
							MESSAGESTRING += str("+"*50+'\n')
		
							print "\nDA Number match"
							print "Found ", title,"at row ",rowList.index(row)+1," in worksheet:",sheet					#add one to row index because of header row
							print "Control Number:", controlNumber														#output information if file matches in database
							print "Entry Date:", entrydate													#output information if file matches in database
							print "+"*50+'\n'
						
							continue																					#move onto next row
				
				except NameError, e:
					continue
				except IndexError, e:
					continue

		MESSAGESTRING += str("+"*50+'\n')
					

if __name__ == "__main__":

	try:
		incomingFiles = getFilenames(NBCPATH)																				#get list of new files
		print incomingFiles
		worksheetDic = breakdownXLS(EXCELFILE)																			#get dictionary of items in the EXCELFILE
		findMatches(incomingFiles, worksheetDic)																			#find any matches and print to terminal
		sendEmail(MESSAGESTRING)	
		
		print "\n"+"-"*50
		if (duplicates == 0):
			print "No duplicates found"
		elif (duplicates > 0):
			print "Found",duplicates,"duplicates"
		print "Done..."	
		print "-"*50
		sys.exit()
	except KeyboardInterrupt:
		print "\nPeace."
		sys.exit()
				
		
#/usr/bin/python

#popup to enter control numbers into excel spreadsheet
#2014
#armen karamian 

import Tkinter as tk
import spreadsheetChecker20140401 as sc
import operator as op
import spreadsheetinserttest20140401 as si
import time

class EntryUnit(tk.Frame,tk.Label,tk.Entry):												#create sub-frames for C# entry
	def __init__(self,parent,file,rowid):
		tk.Frame.__init__(self,parent)
		self.parent = parent
		self.filename = file
		self.controlNumber = tk.StringVar()
		self.controlNumber.set('')
		self.entryLabel = tk.Label(text=self.filename)
		self.entryLabel.grid(row=rowid,column=0,sticky=tk.W)
		self.entryWindow = tk.Entry(width=7,textvariable=self.controlNumber)
		self.entryWindow.grid(row=rowid,column=1,sticky=tk.E)
		
	

class ControlNumberEntryWindow(tk.Frame):													#create main window
	def __init__(self,fileList,exceldictionary,excelfile,parent):												#get incoming file list and sort
		tk.Frame.__init__(self, parent)												#create entry units and place into entry list
		self.master.title("Control number entry")
		self.excelDictionary = exceldictionary
		self.excelfile = excelfile
		self.entryList = []

		fileList.sort()	
		for file in fileList:
			rowid = fileList.index(file)
			
			self.thisEntry = EntryUnit(self,file,rowid)
			self.thisEntry.grid(row=rowid,column=0)
			self.entryList.append(self.thisEntry)
		
		listLength = len(fileList)
		self.grid()
		self.submitButton = tk.Button(self,text='Submit',command=self.submitRows)
		self.submitButton.grid(row=listLength+1,column=0,sticky=tk.W)
	
		
		
	
	def submitRows(self):
		workbook = si.insertNewRows(self.entryList, self.excelDictionary)																			#append new files to spreadsheet
		si.updateWorkbook(self.excelfile, workbook)				

		self.__init__()
	
#		self.quit()
	
	
if __name__ == "__main__":
	time.sleep(10)
	fileList = []
	filenamesDoc = open('/Volumes/dam/_NBCU/EntryScriptOutput/._tempfilenameoutput.txt', 'rb')
	for i in filenamesDoc:
		fileList.append(i)
	app = ControlNumberEntryWindow(fileList, None, None, None)
	app.mainloop()
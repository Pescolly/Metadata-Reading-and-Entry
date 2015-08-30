#!/usr/bin/python


##### implement undo/redo

import multiprocessing as mp
import Tkinter as tk
import spreadsheetChecker20140401 as sc
import spreadsheetinserttest20140401 as si
import NbcFileClass as nbc
import pyqtmovie as qt
import OutputStringMaker as osm

class Application(tk.Frame):
	def __init__(self, master=None):
		tk.Frame.__init__(self,master)
		self.grid()
		self.createQuitBUtton()
		self.createListBox(0,0,20,"Incoming Files")
		self.createListBox(0,1,20,"Duplicates")
		self.createImageBox()
		self.createRefreshButton()
		self.createStatsButton()
		self.createInfoWindow()
		self.createSubmitButton()
		self.incomingFilePath = '/Volumes/fs3/encoding/AssetManagement/nbc_incoming/'
		self.fileout = open('/Volumes/Ingest/AssetManagement/Work/infobox.txt', 'w')

		self.incomingBoxCounter = 0

	def createListBox(self,row,column,height,label,incomingList=None):													#create and populate both list boxes
		self.listboxheader = tk.Label(self,text=label)
		self.listboxheader.grid(column=column, row=row)
		self.listbox = tk.Listbox(self,width=90,height=height,selectmode=tk.EXTENDED)
		self.listbox.grid(column=column,row=row+1,padx=10,sticky=tk.N)
		if (incomingList != None):
			for item in incomingList:
				self.listbox.insert(tk.END,item.filename)
	
	def getInfo(self):																									#get info about selected item in listboxes
		selectedFilename = self.listbox.selection_get()
		selectedFilePath = self.incomingFilePath+selectedFilename														#concat filepath with filename
		fileObj = qt.makeFileObj(selectedFilePath)
		
		self.infoWindow.insert(tk.INSERT, unicode(fileObj.title))	
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.episodeNumber))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.master))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.prodNumber))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.spec1))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.spec2))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.standard))	
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.codec))
		self.infoWindow.insert(tk.END, '\n'+unicode("UNIJA"))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.aspectratio))	
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.recPlace))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.notes))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.minutes))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.seconds))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.daNumber))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.timecode))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.language))
		self.infoWindow.insert(tk.END, '\n'+unicode("Operator Initials"))
		for track in fileObj.audiotracks:
			self.infoWindow.insert(tk.END, '\n'+unicode(track))
		self.infoWindow.insert(tk.END, '\n'+unicode(fileObj.filename))

	def submitInfo(self):
		infoBoxString = self.infoWindow.get('@0,0', tk.END)																#get all lines from info box
		infoBoxString = infoBoxString.replace('\n', '\r\n')																#add windows line breks
#		outputStringList = osm.formatForMVIS(infoBoxString)
		self.fileout.write(infoBoxString)
		self.infoWindow.delete('@0,0',tk.END)																			#clear infobox	


	def createSubmitButton(self):
		self.submitButton = tk.Button(self, text='Submit File Info', command=self.submitInfo)
		self.submitButton.grid(row=4, column = 0,padx=10, sticky=tk.E)
				 	 
	def createInfoWindow(self):
		self.infoWindow = tk.Text(self,borderwidth=10, width=90)
		self.infoWindow.grid(row=2, column=0,padx=10,pady=10, sticky=tk.E+tk.W)
	
	def createStatsButton(self):
		self.statButton = tk.Button(self, text='Info', command=self.getInfo)
		self.statButton.grid(row=3, column = 0,padx=10, sticky=tk.E)
	
	def quitApp(self):
		self.fileout.close()
		self.quit()
		
	def createQuitBUtton(self):
		self.quitButton = tk.Button(self, text='Quit', command=self.quitApp)
		self.quitButton.grid(row=3,column=0,padx=10, sticky=tk.W)

	def refreshList(self):
		workbook = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features_Series_MASTER.xlsx'
		TESTOUTPUT = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features_Series_MASTER_output.xls'

		incomingList = sc.getFilenames(self.incomingFilePath)																#get new filenames
#		print incomingList
		xlsDic = si.breakdownXLS(workbook)
		listofmatches = sc.findMatches(incomingList, xlsDic)																#get duplicates
		for duplicate in listofmatches:																					#remove duplicates from new file list
#			print duplicate.filename
			incomingList.remove(duplicate)							
		si.insertNewRows(incomingList, xlsDic)																			#append new files to spreadsheet
		si.updateWorkbook(TESTOUTPUT, xlsDic)																			#update .xls file
		self.createListBox(0,0,30,"Incoming Files", incomingList)														#put list information on gui
		self.createListBox(0,1,20,"Duplicates",listofmatches)
		
	def createRefreshButton(self):
		self.refreshButton = tk.Button(self, text="REFRESH", command=self.refreshList)
		self.refreshButton.grid(row=3, column=0, sticky=tk.N)
		
		
app = Application()
app.master.title('Project Borg')
app.mainloop()


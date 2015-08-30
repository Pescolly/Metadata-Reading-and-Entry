#!/usr/bin/python


##### implement undo/redo

import multiprocessing as mp
import Tkinter as tk
import spreadsheetChecker20140401 as sc
import spreadsheetinserttest20140401 as si
import NbcFileClass as nbc
import pyqtmovie as qt
import OutputStringMaker as osm

class Window(tk.Frame):
	def __init__(self, parent):
		tk.Frame.__init__(self,parent)
		self.image = tk.PhotoImage(file='slate.gif')
		print dir(self.image)
		tk.Label(self, image=self.image).pack()
		self.pack()

class Application(tk.Frame):
	def __init__(self, master=None):
		tk.Frame.__init__(self,master)
		self.grid()
		self.createQuitBUtton()
		self.createListBox(0,0,30,"Incoming Files")
		self.createListBox(0,1,20,"Duplicates")
		self.createRefreshButton()
		self.createStatsButton()
		self.createEntryWindows()
		self.createSubmitButton()
		self.incomingFilePath = '/Volumes/fs3/encoding/AssetManagement/nbc_incoming/'
		self.fileout = open('/Volumes/Ingest/AssetManagement/Work/infobox.txt', 'w')
		self.incomingBoxCounter = 0
		output = open('/Volumes/dam/_NBCU/EntryScriptOutput/output.txt','wb')		###create blank text file
		output.close()

		
		##### variables that get written to the out file for mvis
		self.titlevar = tk.StringVar()
		self.episodeNumbervar = tk.StringVar() 
		self.prodNumbervar = tk.StringVar()
		self.spec1var = tk.StringVar()
		self.spec2var = tk.StringVar()
		self.standardvar = tk.StringVar()
		self.codecvar = tk.StringVar()
		self.aspectRatiovar = tk.StringVar()
		self.danumbervar = tk.StringVar()
		self.minutesvar = tk.StringVar()
		self.secondsvar = tk.StringVar()
		self.recordPlacevar = tk.StringVar()
		self.OPinitialsvar = tk.StringVar()
		self.timecodevar = tk.StringVar()
		self.languagevar = tk.StringVar()
		self.ch01var = tk.StringVar()
		self.ch02var = tk.StringVar()
		self.ch03var = tk.StringVar()
		self.ch04var = tk.StringVar()
		self.ch05var = tk.StringVar()
		self.ch06var = tk.StringVar()
		self.ch07var = tk.StringVar()
		self.ch08var = tk.StringVar()
		self.ch09var = tk.StringVar()
		self.ch10var = tk.StringVar()
		self.ch11var = tk.StringVar()
		self.ch12var = tk.StringVar()
		self.filenamevar = tk.StringVar()
		
	def createListBox(self,row,column,height,label,incomingList=None):													#create and populate both list boxes
		self.listboxheader = tk.Label(self,text=label)
		self.listboxheader.grid(column=column, row=row)
		self.listbox = tk.Listbox(self,width=90,height=height,selectmode=tk.EXTENDED)
		self.listbox.grid(column=column,row=row+1,padx=10,sticky=tk.N)
		if (incomingList != None):
			for item in incomingList:
				self.listbox.insert(tk.END,item.filename)
		########move to info



	
	def getInfo(self):																									#get info about selected item in listboxes
		selectedFilename = self.listbox.selection_get()
		selectedFilePath = self.incomingFilePath+selectedFilename														#concat filepath with filename
		fileObj = qt.makeFileObj(selectedFilePath)
		self.createEntryWindows(fileObj)																				#draw entry table
		
		root = tk.Toplevel()																							#draw screenshot image
		Window(root)
	

	def submitInfo(self):
		output = open('/Volumes/dam/_NBCU/EntryScriptOutput/output.txt','ab')
		output.write(self.titlevar.get()+'\r\n')
		output.write(self.episodeNumbervar.get()+'\r\n')
		output.write('M\r\n')
		output.write(self.prodNumbervar.get()+'\r\n')
		output.write(self.spec1var.get()+'\r\n')
		output.write(self.spec2var.get()+'\r\n')
		output.write(self.standardvar.get()+'\r\n')
		output.write(self.codecvar.get()+'\r\n')
		output.write('UNIJA\r\n')	
		output.write(self.aspectRatiovar.get()+'\r\n')
		output.write(self.recordPlacevar.get()+'\r\n')
		output.write('Y\r\n')
		output.write(self.minutesvar.get()+'\r\n')
		output.write(self.secondsvar.get()+'\r\n')
		output.write(self.danumbervar.get()+'\r\n')
		output.write(self.timecodevar.get()+'\r\n')
		output.write(self.languagevar.get()+'\r\n')
		output.write(self.OPinitialsvar.get()+'\r\n')
		output.write(self.ch01var.get()+'\r\n')
		output.write(self.ch02var.get()+'\r\n')
		output.write(self.ch03var.get()+'\r\n')
		output.write(self.ch04var.get()+'\r\n')
		output.write(self.ch05var.get()+'\r\n')
		output.write(self.ch06var.get()+'\r\n')
		output.write(self.ch07var.get()+'\r\n')
		output.write(self.ch08var.get()+'\r\n')
		output.write(self.ch09var.get()+'\r\n')
		output.write(self.ch10var.get()+'\r\n')
		output.write(self.ch11var.get()+'\r\n')
		output.write(self.ch12var.get()+'\r\n')
		output.close()
		
	def createEntryWindows(self,file=None):
		if (file != None):
			LANGS = ['ENG', 'GER', 'CSP', 'FRT', 'BPO', 'LAS']
			self.filenamevar.set(file.filename)
			self.titlevar.set(file.title)
			self.episodeNumbervar.set(file.episodeNumber)
			self.prodNumbervar.set(file.prodNumber)
			self.spec1var.set(file.spec1)
			self.spec2var = tk.StringVar()
			self.spec2var.set(file.spec1)
			self.standardvar.set(file.standard)
			self.codecvar.set(file.codec)
			self.aspectRatiovar.set(file.aspectratio)
			self.danumbervar.set(file.daNumber)
			self.minutesvar.set(file.minutes)
			self.secondsvar.set(file.seconds)
			self.recordPlacevar.set(file.recPlace)
			self.timecodevar.set(file.timecode)
			for i in LANGS:
				if (i == file.language):
					if (i == 'ENG'):
						if (file.flag51 == True):
							entryLanguage = 'e'
							break
						else:
							entryLanguage = 'G'
							break
					if (i == 'GER'):
						if (file.flag51 == True):
							entryLanguage = 'g'
							break
						else:
							entryLanguage = 'A'
							break
					if (i == 'CSP'):
						if (file.flag51 == True):
							entryLanguage = 'c'
							break
						else:
							entryLanguage = 'D'
							break
					if (i == 'FRT'):
						if (file.flag51 == True):
							entryLanguage = 'f'
							break
						else:
							entryLanguage = 'I'
							break
					if (i == 'BPO'):
						if (file.flag51 == True):
							entryLanguage = 't'
							break
						else:
							entryLanguage = 'j'
							break
					if (i == 'LAS'):
						if (file.flag51 == True):
							entryLanguage = 'l'
							break
						else:
							entryLanguage = 'L'
							break

			self.languagevar.set(entryLanguage)
			self.OPinitialsvar.set(' ')
			self.ch01var.set(file.audiotracks[0])
			self.ch02var.set(file.audiotracks[1])
			self.ch03var.set(file.audiotracks[2])
			self.ch04var.set(file.audiotracks[3])
			self.ch05var.set(file.audiotracks[4])
			self.ch06var.set(file.audiotracks[5])
			self.ch07var.set(file.audiotracks[6])
			self.ch08var.set(file.audiotracks[7])
			self.ch09var.set(file.audiotracks[8])
			self.ch10var.set(file.audiotracks[9])
			self.ch11var.set(file.audiotracks[10])
			self.ch12var.set(file.audiotracks[11])
			
		
			self.title = tk.Entry(self, textvariable=self.titlevar).grid(row=2, column=0, sticky=tk.E,padx=10)
			self.titlelabel = tk.Label(self, text='title').grid(row=2, column=0, sticky=tk.W,padx=10)
		
			self.episodeNumber = tk.Entry(self, textvariable=self.episodeNumbervar).grid(row=3, column=0, sticky=tk.E,padx=10)
			self.episodelabel = tk.Label(self, text='episode number').grid(row=3, column=0, sticky=tk.W,padx=10)

			self.prodNumber = tk.Entry(self, textvariable=self.prodNumbervar).grid(row=4, column=0, sticky=tk.E,padx=10)
			self.prodlabel = tk.Label(self, text='production Number').grid(row=4, column=0, sticky=tk.W,padx=10)

			self.spec1 = tk.Entry(self, textvariable=self.spec1var).grid(row=5, column=0, sticky=tk.E,padx=10)
			self.spec1label = tk.Label(self, text='Spec 1').grid(row=5, column=0, sticky=tk.W,padx=10)

			self.spec2 = tk.Entry(self, textvariable=self.spec2var).grid(row=6, column=0, sticky=tk.E,padx=10)
			self.spec2label = tk.Label(self, text='Spec 2').grid(row=6, column=0, sticky=tk.W,padx=10)

			self.standard = tk.Entry(self, textvariable=self.standardvar).grid(row=7, column=0, sticky=tk.E,padx=10)
			self.standardlabel = tk.Label(self, text='Standard: Pal, Ntsc, 3=1080p23, 1=1080i59, 2=1080i50').grid(row=7, column=0, sticky=tk.W,padx=10)

			self.codec = tk.Entry(self, textvariable=self.codecvar).grid(row=8, column=0, sticky=tk.E,padx=10)
			self.codeclabel = tk.Label(self, text='codec').grid(row=8, column=0, sticky=tk.W,padx=10)

			self.aspectRatio = tk.Entry(self, textvariable=self.aspectRatiovar).grid(row=9, column=0, sticky=tk.E,padx=10)
			self.aspectRatioLabel = tk.Label(self, text='4x3: 3=133, L=178, C=185 16x9: H=1.33, 8=178, J=185, K=235').grid(row=9, column=0, sticky=tk.W,padx=10)

			self.recPlace = tk.Entry(self, textvariable=self.recordPlacevar).grid(row=10, column=0, sticky=tk.E,padx=10)
			self.recplacelabel = tk.Label(self, text='Rec Place').grid(row=10, column=0, sticky=tk.W,padx=10)

			self.trtminutes = tk.Entry(self, textvariable=self.minutesvar).grid(row=11, column=0, sticky=tk.N)
			self.trtlabel = tk.Label(self, text='Seconds:				  ').grid(row=11, column=0, sticky=tk.E, padx=10)
			self.trtseconds = tk.Entry(self, textvariable=self.secondsvar).grid(row=11, column=0, sticky=tk.E,padx=10)
			self.trtlabel = tk.Label(self, text='TRT			    Minutes:').grid(row=11, column=0, sticky=tk.W,padx=10)

			self.daNumber = tk.Entry(self, textvariable=self.danumbervar).grid(row=12, column=0, sticky=tk.E,padx=10)
			self.daNumber = tk.Label(self, text='DA NUmber').grid(row=12, column=0, sticky=tk.W,padx=10)

			self.timecode = tk.Entry(self, textvariable=self.timecodevar).grid(row=13, column=0, sticky=tk.E,padx=10)
			self.timecodeLabel = tk.Label(self, text='Timecode D=30DF, N=30NDF, P=24, E=25').grid(row=13, column=0, sticky=tk.W,padx=10)

			self.language = tk.Entry(self, textvariable=self.languagevar).grid(row=14, column=0, sticky=tk.E,padx=10)
			self.languageLabel = tk.Label(self, text='Language').grid(row=14, column=0, sticky=tk.W,padx=10)

			self.initials = tk.Entry(self, textvariable=self.OPinitialsvar).grid(row=2, column=1, sticky=tk.E,padx=10)
			self.initialLabel = tk.Label(self, text='Operator Initials').grid(row=2, column=1, sticky=tk.W,padx=10)

			self.initials = tk.Entry(self, textvariable=self.filenamevar).grid(row=3, column=1, sticky=tk.E,padx=10)
			self.initialLabel = tk.Label(self, text='Filename').grid(row=3, column=1, sticky=tk.W,padx=10)


			self.AudioLabel = tk.Label(self, text='Audio Channels').grid(row=15,column=0)
			self.channel01 = tk.Entry(self, textvariable=self.ch01var).grid(row=16, column=0, sticky=tk.W,padx=10)
			self.channel02 = tk.Entry(self, textvariable=self.ch02var).grid(row=16, column=0, sticky=tk.E,padx=10)
			self.channel03 = tk.Entry(self, textvariable=self.ch03var).grid(row=17, column=0, sticky=tk.W,padx=10)
			self.channel04 = tk.Entry(self, textvariable=self.ch04var).grid(row=17, column=0, sticky=tk.E,padx=10)
			self.channel05 = tk.Entry(self, textvariable=self.ch05var).grid(row=18, column=0, sticky=tk.W,padx=10)
			self.channel06 = tk.Entry(self, textvariable=self.ch06var).grid(row=18, column=0, sticky=tk.E,padx=10)
			self.channel07 = tk.Entry(self, textvariable=self.ch07var).grid(row=19, column=0, sticky=tk.W,padx=10)
			self.channel08 = tk.Entry(self, textvariable=self.ch08var).grid(row=19, column=0, sticky=tk.E,padx=10)
			self.channel09 = tk.Entry(self, textvariable=self.ch09var).grid(row=20, column=0, sticky=tk.W,padx=10)
			self.channel10 = tk.Entry(self, textvariable=self.ch10var).grid(row=20, column=0, sticky=tk.E,padx=10)
			self.channel11 = tk.Entry(self, textvariable=self.ch11var).grid(row=21, column=0, sticky=tk.W,padx=10)
			self.channel12 = tk.Entry(self, textvariable=self.ch12var).grid(row=21, column=0, sticky=tk.E,padx=10)


	
	def createStatsButton(self):
		self.statButton = tk.Button(self, text='Info', command=self.getInfo)
		self.statButton.grid(row=22, column = 0,padx=10, sticky=tk.E)
	
	def quitApp(self):
		self.fileout.close()
		self.quit()

	def createSubmitButton(self):
		self.submitButton = tk.Button(self, text='Submit File Info', command=self.submitInfo)
		self.submitButton.grid(row=23, column = 0,padx=10, sticky=tk.E)

	def createQuitBUtton(self):
		self.quitButton = tk.Button(self, text='Quit', command=self.quitApp, activeforeground='red')
		self.quitButton.grid(row=22,column=0,padx=10, sticky=tk.W)

	def createRefreshButton(self):
		self.refreshButton = tk.Button(self, text="REFRESH", command=self.refreshList)
		self.refreshButton.grid(row=22, column=0, sticky=tk.N)


	def refreshList(self):
		workbook = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features_Series_MASTER.xls'
		TESTOUTPUT = '/Volumes/dam/OPERATOR/AKARAMIAN/armenScripts/python/spreadsheetChecker/NBCU_Features_Series_MASTER.xls'

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
		
		
		
app = Application()
app.master.title('Project Borg')
app.mainloop()


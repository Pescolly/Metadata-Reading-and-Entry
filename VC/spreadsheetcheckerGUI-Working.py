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
		output = open('/Volumes/dam/_NBCU/EntryScriptOutput/output.txt','wb')
		output.close()


		##### variables that get written to the out file for mvis
		self.outtitle = None
		self.outepisodeNumber = None
		self.outprodNumber = None
		self.outspec1 = None
		self.outspec2 = None
		self.outstandard = None
		self.outcodec = None
		self.outaspectRatio = None
		self.outdanumber = None
		self.outminutes  = None
		self.outseconds = None
		self.outrecordPlace = None
		self.outtimecode = None
		self.outlanguage = None
		self.outch01 = None
		self.outch02 = None
		self.outch03 = None
		self.outch04 = None
		self.outch05 = None
		self.outch06 = None
		self.outch07 = None
		self.outch08 = None
		self.outch09 = None
		self.outch10 = None
		self.outch11 = None
		self.outch12 = None


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
		output.write(self.outtitle+'\r\n')
		output.write(self.outepisodeNumber+'\r\n')
		output.write(self.outprodNumber+'\r\n')
		output.write(self.outspec1+'\r\n')
		output.write(self.outspec2+'\r\n')
		output.write(self.outstandard+'\r\n')
		output.write(self.outcodec+'\r\n')
		output.write(self.outaspectRatio+'\r\n')
		output.write(self.outdanumber+'\r\n')
		output.write(self.outminutes+'\r\n')
		output.write(self.outseconds+'\r\n')
		output.write(self.outrecordPlace+'\r\n')
		output.write(self.outtimecode+'\r\n')
		output.write(self.outlanguage+'\r\n')
		output.write(self.outch01+'\r\n')
		output.write(self.outch02+'\r\n')
		output.write(self.outch03+'\r\n')
		output.write(self.outch04+'\r\n')
		output.write(self.outch05+'\r\n')
		output.write(self.outch06+'\r\n')
		output.write(self.outch07+'\r\n')
		output.write(self.outch08+'\r\n')
		output.write(self.outch09+'\r\n')
		output.write(self.outch10+'\r\n')
		output.write(self.outch11+'\r\n')
		output.write(self.outch12+'\r\n')
		output.close()
		
	def createEntryWindows(self,file=None):
		if (file != None):
			title = tk.StringVar()
			title.set(file.title)
			episodeNumber = tk.StringVar() 
			episodeNumber.set(file.episodeNumber)
			prodNumber = tk.StringVar()
			prodNumber.set(file.prodNumber)
			spec1 = tk.StringVar()
			spec1.set(file.spec1)
			spec2 = tk.StringVar()
			spec2.set(file.spec1)
			standard = tk.StringVar()
			standard.set(file.standard)
			codec = tk.StringVar()
			codec.set(file.codec)
			aspectRatio = tk.StringVar()
			aspectRatio.set(file.aspectratio)
			danumber = tk.StringVar()
			danumber.set(file.daNumber)
			minutes = tk.StringVar()
			minutes.set(file.minutes)
			seconds = tk.StringVar()
			seconds.set(file.seconds)
			recordPlace = tk.StringVar()
			recordPlace.set(file.recPlace)
			timecode = tk.StringVar()
			timecode.set(file.timecode)
			language = tk.StringVar()
			language.set(file.language)
			ch1 = tk.StringVar()
			ch1.set(file.audiotracks[0])
			ch2 = tk.StringVar()
			ch2.set(file.audiotracks[1])
			ch3 = tk.StringVar()
			ch3.set(file.audiotracks[2])
			ch4 = tk.StringVar()
			ch4.set(file.audiotracks[3])
			ch5 = tk.StringVar()
			ch5.set(file.audiotracks[4])
			ch6 = tk.StringVar()
			ch6.set(file.audiotracks[5])
			ch7 = tk.StringVar()
			ch7.set(file.audiotracks[6])
			ch8 = tk.StringVar()
			ch8.set(file.audiotracks[7])
			ch9 = tk.StringVar()
			ch9.set(file.audiotracks[8])
			ch10 = tk.StringVar()
			ch10.set(file.audiotracks[9])
			ch11 = tk.StringVar()
			ch11.set(file.audiotracks[10])
			ch12 = tk.StringVar()
			ch12.set(file.audiotracks[11])
			
		
			self.title = tk.Entry(self, textvariable=title).grid(row=2, column=0, sticky=tk.N)
			self.titlelabel = tk.Label(self, text='title').grid(row=2, column=0, sticky=tk.W,padx=10)
		
			self.episodeNumber = tk.Entry(self, textvariable=episodeNumber).grid(row=3, column=0, sticky=tk.N)
			self.episodelabel = tk.Label(self, text='episode number').grid(row=3, column=0, sticky=tk.W,padx=10)

			self.prodNumber = tk.Entry(self, textvariable=prodNumber).grid(row=4, column=0, sticky=tk.N)
			self.prodlabel = tk.Label(self, text='production Number').grid(row=4, column=0, sticky=tk.W,padx=10)

			self.spec1 = tk.Entry(self, textvariable=spec1).grid(row=5, column=0, sticky=tk.N)
			self.spec1label = tk.Label(self, text='Spec 1').grid(row=5, column=0, sticky=tk.W,padx=10)

			self.spec2 = tk.Entry(self, textvariable=spec2).grid(row=6, column=0, sticky=tk.N)
			self.spec2label = tk.Label(self, text='Spec 2').grid(row=6, column=0, sticky=tk.W,padx=10)

			self.standard = tk.Entry(self, textvariable=standard).grid(row=7, column=0, sticky=tk.N)
			self.standardlabel = tk.Label(self, text='standard').grid(row=7, column=0, sticky=tk.W,padx=10)

			self.codec = tk.Entry(self, textvariable=codec).grid(row=8, column=0, sticky=tk.N)
			self.codeclabel = tk.Label(self, text='codec').grid(row=8, column=0, sticky=tk.W,padx=10)

			self.aspectRatio = tk.Entry(self, textvariable=aspectRatio).grid(row=9, column=0, sticky=tk.N)
			self.aspectRatioLabel = tk.Label(self, text='Aspect Ratio').grid(row=9, column=0, sticky=tk.W,padx=10)

			self.recPlace = tk.Entry(self, textvariable=recordPlace).grid(row=10, column=0, sticky=tk.N)
			self.recplacelabel = tk.Label(self, text='Rec Place').grid(row=10, column=0, sticky=tk.W,padx=10)

			self.trtminutes = tk.Entry(self, textvariable=minutes).grid(row=11, column=0, sticky=tk.N)
			self.trtlabel = tk.Label(self, text='Seconds:				  ').grid(row=11, column=0, sticky=tk.E, padx=10)
			self.trtseconds = tk.Entry(self, textvariable=seconds).grid(row=11, column=0, sticky=tk.E)
			self.trtlabel = tk.Label(self, text='TRT			    Minutes:').grid(row=11, column=0, sticky=tk.W,padx=10)

			self.daNumber = tk.Entry(self, textvariable=danumber).grid(row=12, column=0, sticky=tk.N)
			self.daNumber = tk.Label(self, text='DA NUmber').grid(row=12, column=0, sticky=tk.W,padx=10)

			self.timecode = tk.Entry(self, textvariable=timecode).grid(row=13, column=0, sticky=tk.N)
			self.timecodeLabel = tk.Label(self, text='timecodeLabel').grid(row=13, column=0, sticky=tk.W,padx=10)

			self.language = tk.Entry(self, textvariable=language).grid(row=14, column=0, sticky=tk.N)
			self.languageLabel = tk.Label(self, text='Language').grid(row=14, column=0, sticky=tk.W,padx=10)

			self.AudioLabel = tk.Label(self, text='Audio Channels').grid(row=15,column=0)
			self.channel01 = tk.Entry(self, textvariable=ch1).grid(row=16, column=0, sticky=tk.W,padx=10)
			self.channel02 = tk.Entry(self, textvariable=ch2).grid(row=16, column=0, sticky=tk.E,padx=10)
			self.channel03 = tk.Entry(self, textvariable=ch3).grid(row=17, column=0, sticky=tk.W,padx=10)
			self.channel04 = tk.Entry(self, textvariable=ch4).grid(row=17, column=0, sticky=tk.E,padx=10)
			self.channel05 = tk.Entry(self, textvariable=ch5).grid(row=18, column=0, sticky=tk.W,padx=10)
			self.channel06 = tk.Entry(self, textvariable=ch6).grid(row=18, column=0, sticky=tk.E,padx=10)
			self.channel07 = tk.Entry(self, textvariable=ch7).grid(row=19, column=0, sticky=tk.W,padx=10)
			self.channel08 = tk.Entry(self, textvariable=ch8).grid(row=19, column=0, sticky=tk.E,padx=10)
			self.channel09 = tk.Entry(self, textvariable=ch9).grid(row=20, column=0, sticky=tk.W,padx=10)
			self.channel10 = tk.Entry(self, textvariable=ch10).grid(row=20, column=0, sticky=tk.E,padx=10)
			self.channel11 = tk.Entry(self, textvariable=ch11).grid(row=21, column=0, sticky=tk.W,padx=10)
			self.channel12 = tk.Entry(self, textvariable=ch12).grid(row=21, column=0, sticky=tk.E,padx=10)

			self.outtitle = title.get()
			self.outepisodeNumber = episodeNumber.get()
			self.outprodNumber = prodNumber.get()
			self.outspec1 = spec1.get()
			self.outspec2 = spec2.get()
			self.outstandard = standard.get()
			self.outcodec = codec.get()
			self.outaspectRatio = aspectRatio.get()
			self.outrecPlace = recordPlace.get()
			self.outdanumber = danumber.get()
			self.outminutes  = minutes.get()
			self.outseconds = seconds.get()
			self.outrecordPlace = recordPlace.get()
			self.outtimecode = timecode.get()
			self.outlanguage = language.get()
			self.outch01 = ch1.get()
			self.outch02 = ch2.get()
			self.outch03 = ch3.get()
			self.outch04 = ch4.get()
			self.outch05 = ch5.get()
			self.outch06 = ch6.get()
			self.outch07 = ch7.get()
			self.outch08 = ch8.get()
			self.outch09 = ch9.get()
			self.outch10 = ch10.get()
			self.outch11 = ch11.get()
			self.outch12 = ch12.get()
	
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
		
		
		
app = Application()
app.master.title('Project Borg')
app.mainloop()


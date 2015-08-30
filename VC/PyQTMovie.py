#!/usr/bin/python


# modified version of pyqtmovie info getter for nbc spreadsheet checker
# uses NBCFile class as primary object for input and output
#class NBCFile:
#	def __init__(self, filename):
#		self.filename = filename
#		self.rowID = None
#		self.controlNumber = None
#		self.entryDate = None
#		self.standard = None
#		self.aspectratio = None
#		self.canvas = None
#		self.videotrack = None
#		self.tctrack = None
#		self.framerate = None
#		self.audiotracks = []
#
#
#

import struct, QTKit, objc, PyObjCTools, Foundation, os, string, re, AppKit
from Foundation import NSData, NSMutableDictionary, NSAutoreleasePool, NSURL
import NbcFileClass as nbc


TIMECODE = 'tmcd'
VIDEO = 'vide'
AUDIO = 'soun'
HD = 'HD'
NTSC = 'NTSC'
PAL = 'PAL'

def createMovFile(fullfilepath):
	MEDIA_SPECS = [																										### specs in media header
		'QTMovieCurrentSizeAttribute', 'QTMovieCreationTimeAttribute', 'QTMovieHasAudioAttribute',
		'QTMovieHasVideoAttribute', 'QTMovieTimeScaleAttribute']
	
	TRACK_SPECS = [																										### specs in track header
		'QTTrackDisplayNameAttribute', 'QTTrackBoundsAttribute',
		'QTTrackFormatSummaryAttribute', 'QTTrackIDAttribute', 'QTTrackMediaTypeAttribute']

	attribs = NSMutableDictionary.dictionary()
	attribs['QTMovieFileNameAttribute'] = fullfilepath
	
	file = nbc.NBCFile(fullfilepath)
	mov, error = QTKit.QTMovie.alloc().initWithAttributes_error_(attribs,objc.nil)										#init mov object
	if error:
		print error, file
	else:
		for track in mov.tracks():																						#pull individual tracks
			try:			
				tracktype = track.trackAttributes()['QTTrackMediaTypeAttribute']
				if (tracktype == TIMECODE):
					file.tctrack = track
				elif (tracktype == VIDEO):
					file.videotrack = track
					file.codec = 'PRH' #file.videotrack.trackAttributes()['QTTrackFormatSummaryAttribute']						#get codec
				elif (tracktype == AUDIO):
					file.audiotracks.append(track.trackAttributes()['QTTrackDisplayNameAttribute'])
				
			except KeyError, e:
				continue	

		frameRate = mov.duration()[1]												#set framerate
		duration = mov.duration()[0]
#		print frameRate
		#print duration
		durMinutes = duration/frameRate/60														#get minutes of duration
		durSeconds = int((round(duration/frameRate/60.00,2)-(duration/frameRate/60))*60)		#get seconds of duration(fraction of minutes multiplied by 60 to get actual seconds)
		#print durMinutes
		##print durSeconds
		file.minutes = durMinutes
		file.seconds = durSeconds
		if ((frameRate == 2500) or (frameRate == 25)):
			file.timecode = 'E'
		if ((frameRate == 23976) or (frameRate == 24000)):
			file.timecode = 'P'
		if ((frameRate == 30000) or (frameRate == 2997)):
			file.timecode = 'D'

		if (file.videotrack.currentSize().height > 1050 and file.videotrack.currentSize().height < 1110):				#set standard with height
			if (file.timecode == 'P'):
				file.standard = 3																						#MVIS CODES: 3 - 1080-2398
			if (file.timecode == 'E'):
				file.standard = 2																						#... 2 - 1080 50i
			if (file.timecode == 'D'):
				file.standard = 1																						#... 1 - 1080 5994i
		elif (file.videotrack.currentSize().height > 470 and file.videotrack.currentSize().height < 490):
			file.standard = 'N'																							#... N - NTSC
		elif (file.videotrack.currentSize().height > 560 and file.videotrack.currentSize().height < 590):
			file.standard = 'P'																							#... P - PAL
		else:
			file.standard = None



			################ create a png of the slate and save to disk ###################			
	time = QTKit.QTTime()
	time.timeScale = mov.movieAttributes()['QTMovieTimeScaleAttribute']
	time.timeValue = frameRate * 75															### guess at location of slate (75 seconds in)

	newSize = Foundation.NSSize()									### pull slate image and resize to 720x480
	newSize.width = 720
	newSize.height = 480
	mov.setCurrentSize_(newSize)
	image = mov.frameImageAtTime_(time)
	imageRep = image.representations()[0]
	output = imageRep.representationUsingType_properties_(AppKit.NSGIFFileType, objc.nil)
	outputfile = 'slate.gif'
	output.writeToFile_atomically_(outputfile, None)
	
				
	return file

def getcanvas(file):
	CANVASES = ['4x3..','16x9..']
	filename = os.path.split(file.filename)[1]
	filesplit = string.split(filename, '_')	
	for i in filesplit:
		for j in CANVASES:
			if re.match(j, string.lower(i)):
				if (j == CANVASES[0]):
					canvas = '4x3'
					break
				if (j == CANVASES[1]):
					canvas = '16x9'	
					break
	
	return canvas
					
def getaspectratio(file):
	arList = ['133', '178', '185', '235', '240']
	arDic43 = {'133':'3', '178':'L','185':'C', '235':'D', '240':'G'}
	arDic169 = {'133':'H', '178':'8','185':'J', '235':'K', '240':'P'}
	
	filesplit = string.split(file.filename, '_')
	for fileAR in filesplit:
		#print fileAR
		for ar in arList:
			#print ar
			if (fileAR == ar):
#				print ar
#				print file.canvas
				if (file.canvas == '4x3'):
#					print arDic43[ar]
					return arDic43[ar]
				if (file.canvas == '16x9'):
#					print arDic43[ar]
					return arDic169[ar]
		
def getTitle(fullfilepath):
	filename = os.path.basename(fullfilepath)
	title = string.upper(string.split(filename, "_")[0])
	return title

def getFileName(fullfilepath):
	filename = os.path.basename(fullfilepath)
	return filename

def getProdNumber(filename):
	return (re.findall('.*[-|_](\w\w\w\d\d)[-|_].*', string.lower(filename)) or 'None')

def getDANumber(filename):
	daPattern = re.compile('.+_(DA\d\d\d\d\d\d\d\d\d)_.*')														#re pattern to get DA number
	try:
		return re.findall(daPattern, filename)[0]	
	except IndexError, e:
		return " "

def getLanguage(filename):
	LANGS = ['ENG', 'GER', 'CSP', 'FRT', 'BPO', 'LAS']
	filenamesplit = string.split(filename, '_')
	
	for lang in LANGS:
		for segment in filenamesplit:
			if (string.lower(lang) == string.lower(segment)):
				return lang
		
def assignAudioTracks(file):
	LANGS = ['ENG', 'GER', 'CSP', 'FRT', 'BPO', 'LAS']
	trackcount = len(file.audiotracks)

	if (file.language == LANGS[0]):
		entryLanguage = 'ENGLISH'
	if (file.language == LANGS[1]):
		entryLanguage = 'GERMAN'
	if (file.language == LANGS[2]):
		entryLanguage = 'CSP'
	if (file.language == LANGS[3]):
		entryLanguage = 'FRENCH'
	if (file.language == LANGS[4]):
		entryLanguage = 'BPO'
	if (file.language == LANGS[5]):
		entryLanguage = 'LAS'
	if (trackcount == 4):
		file.audiotracks[0] = "Lt. "+entryLanguage
		file.audiotracks[1] = "Rt. "+entryLanguage
		file.audiotracks[2] = "ML"
		file.audiotracks[3] = "MR"
		file.flag51 = False
		
	if (trackcount == 10):
		if (file.language == LANGS[0]):
			file.audiotracks[0] = 'LEFT '+file.language
			file.audiotracks[1] = 'RIGHT '+file.language
			file.audiotracks[2] = 'CENTER '+file.language
			file.audiotracks[3] = 'LFE'
			file.audiotracks[4] = 'LT SURR '+file.language
			file.audiotracks[5] = 'RT SURR '+file.language
			file.audiotracks[6] = 'LT '+entryLanguage		
			file.audiotracks[7] = 'RT '+entryLanguage
			file.audiotracks[8] = 'ML'
			file.audiotracks[9] = 'MR'
			file.flag51 = True
		else:
			file.audiotracks[0] = 'LEFT '+file.language
			file.audiotracks[1] = 'RIGHT '+file.language
			file.audiotracks[2] = 'CENTER '+file.language
			file.audiotracks[3] = 'LFE'
			file.audiotracks[4] = 'LT SURR '+file.language
			file.audiotracks[5] = 'RT SURR '+file.language
			file.audiotracks[6] = 'LT '+entryLanguage
			file.audiotracks[7] = 'RT '+entryLanguage
			file.audiotracks[8] = 'L'
			file.audiotracks[9] = 'R'
			file.flag51 = True


	
	while trackcount < 12:
		file.audiotracks.append(" ")
		trackcount += 1		
#	#print file.audiotracks
	return file.audiotracks	

def getspecs(filename):
	SPEC1 = ['"Title" (TEXTED VERSION)','"Title" (TEXTLESS VERSION)' ] 
	SPEC2 = ['BTS - NO TEXTLESS @ TAIL','BTS W/ TEXTLESS @ TAIL','NO BTS / TEXTLESS @ TAIL','NO BTS - W/ TEXTLESS @ TAIL']
	spec1 = 'SPEC 1'
	spec2 = 'SPEC 2'

	filesplit = string.split(filename, '_')
	for segment in filesplit:
		print segment
		if (string.lower(segment) == 'televisiontextless'):
			spec1 = SPEC1[1]
			break
		elif (string.lower(segment) == 'televisiontexted'):
			spec1 = SPEC1[0]
			break

	return spec1,spec2

def makeFileObj(fullfilepath):
	file = createMovFile(fullfilepath)
		
	file.title = getTitle(fullfilepath)
	file.filename = getFileName(fullfilepath)
	file.master = "M"
	file.prodNumber = string.upper(getProdNumber(file.filename)[0])
	file.spec1,file.spec2 = getspecs(file.filename)
	file.canvas = getcanvas(file)
	file.aspectratio = getaspectratio(file)
	file.daNumber = getDANumber(file.filename)
	file.language = getLanguage(file.filename)
	file.audiotracks = assignAudioTracks(file)
	
	return file


if __name__ == "__main__":
	for i in os.listdir('/Volumes/fs3/encoding/AssetManagement/nbc_incoming/20140418/'):
		if (i.startswith('.')):
			continue
		print i
		i = '/Volumes/fs3/encoding/AssetManagement/nbc_incoming/20140418/'+i
		file = makeFileObj(i)
		print file.filename
		print file.spec1, file.spec2
		print '\n'
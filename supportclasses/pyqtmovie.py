#!/usr/bin/python

import struct, QTKit, objc, PyObjCTools, Foundation
from Foundation import NSMutableDictionary, NSAutoreleasePool, NSURL

testFile = '/Volumes/Ingest/temp-dearchive/RealHousewivesOfBeverlyHillsThe_DA000477235_VQQ18_EPS_TelevisionTexted_5994_16x9FF_178_ENG_HD_ProResHQ_220M.mov'




class mediaObject:																										#create qt media object for .movs, wavs, etc....
	

	def __init__(self, filename):
		self.filename = filename																						
		self.movies = []
		self.MEDIA_SPECS = [																										### specs in media header
		'QTMovieCurrentSizeAttribute', 'QTMovieCreationTimeAttribute', 'QTMovieHasAudioAttribute',
		'QTMovieHasVideoAttribute', 'QTMovieTimeScaleAttribute'
		]
		
		self.TRACK_SPECS = [																										### specs in track header
		'QTTrackDisplayNameAttribute', 'QTTrackBoundsAttribute', 'QTTrackDisplayNameAttribute',
		'QTTrackFormatSummaryAttribute', 'QTTrackIDAttribute', 'QTTrackMediaTypeAttribute'
		]


		
	def printAttributes(self):																							#print out various attributes
		attribs = NSMutableDictionary.dictionary()
		attribs['QTMovieFileNameAttribute'] = self.filename
		mov, error = QTKit.QTMovie.alloc().initWithAttributes_error_(attribs,objc.nil)
		if error:
			print error
		else:
			print 'Duration:',mov.duration()[0]/mov.duration()[1]/60.0
			for i in self.MEDIA_SPECS:
				print i,":", type(mov.movieAttributes()[i])
			for i in mov.tracks():
				for j in self.TRACK_SPECS:
					try:
						print i.trackAttributes()[j]
					except KeyError, e:
						continue	
						
						
filecheck = mediaObject(testFile)

filecheck.printAttributes()


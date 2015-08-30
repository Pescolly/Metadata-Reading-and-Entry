class NBCFile:
	def __init__(self, filename):		#automated retrieval
		self.title = ''				#done
		self.episodeNumber = "Episode #"
		self.master = "M"
		self.prodNumber = ''			#done
		self.spec1 = "Spec 1"
		self.spec2 = "Spec 2"
		self.standard = ''			#done
		self.codec = ''				#done
		self.aspectratio = ''			#done
		self.recPlace = "rec place"
		self.notes = "Y"
		self.minutes = "Minutes"
		self.seconds = "Seconds"
		self.daNumber = ''			#done
		self.timecode = ''			#done
		self.language = ''			#done
		self.filename = filename		#done
		self.flag51 = ''
		self.audiotracks = []			#done


		self.rowID = None
		self.controlNumber = None
		self.entryDate = None
		self.canvas = None
		self.videotrack = None
		self.tctrack = None
		
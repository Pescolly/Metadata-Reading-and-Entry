#!/usr/bin/python


import os, string, re

def getTitle(fullfilepath):
	
	filename = os.path.basename(fullfilepath)
	title = string.split(filename,"_")[0]
	
	deCamelCasedList = re.split('([A-Z][a-z]*|[0-9]*)',title)													#use camelcasing for irony
	deCamelCasedTitle = ''
	try:
		while True:
			deCamelCasedList.remove('')
	except ValueError, e:
		pass
	for i in deCamelCasedList:
		if (i == deCamelCasedList[-2] and deCamelCasedList[-1] == 'The'):
			i += ','
		if (i == deCamelCasedList[-1]):
			deCamelCasedTitle += (str(i).upper())
		elif (i != ''):
			deCamelCasedTitle += (str(i).upper()+' ')
			
	if (deCamelCasedTitle[-4::] == ' THE'):
		print deCamelCasedTitle[-4::]
		
			
	return deCamelCasedTitle
	


if __name__ == "__main__":		
	stringa = 'RealHousewivesOfNewYorkCityThe_DA000503662_VPZ14_EPS_TelevisionTextless_50_16x9FF_178_ENG_HD_ProResHQ_220M'
	print getTitle(stringa)
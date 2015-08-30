#!/usr/bin/python

#takes the incoming string from the infobox reader and formats it for MVIS entry

########################### INCOMING FORMAT ###########################################
#CuriousGeorge
#None
#M
#AA208
#Spec 1
#Spec 2
#HD
#PRH
#178
#rec place
#Y
#41193000
#Seconds
#DA000479798
#23976
#GER
#Sound Track 1
#Sound Track 2
#Sound Track 3
#Sound Track 4
#Blank Track
#Blank Track
#Blank Track
#Blank Track
#Blank Track
#Blank Track
#Blank Track
#Blank Track
#CuriousGeorge_DA000479798_AA208_EPS_Television_2398_16x9FF_178_GER_HD_ProResHQ_220M.mov

########################### OUTGOING FORMAT ###########################################
#BATES MOTEL
#2005
#M
#CMG05
#"THE ESCAPE ARTIST" W/ TEXTLESS @ TAIL
#BTS W/ PULLED BLACKS
#3
#PRH
#UNIJA
#8
#UDS
#Y
#42
#55
#DA000482360
#P
#e
#AKaramian
#LEFT ENG
#RIGHT ENG
#CENTER ENG
#LFE
#LT SURR ENG
#RT SURR ENG
#LT ENGLISH
#RT ENGLISH
#ML
#MR
# 
# 
#BatesMotel_DA000482360_CMG05_EPS_Television_2398_16x9FF_178_ENG_HD_ProResHQ_220M.mov 

import string

instring = '''BATES MOTEL
2005
M
CMG05
"THE ESCAPE ARTIST" W/ TEXTLESS @ TAIL
BTS W/ PULLED BLACKS
3
PRH
UNIJA
8
UDS
Y
42
55
DA000482360
P
e
AKaramian
LEFT ENG
RIGHT ENG
CENTER ENG
LFE
LT SURR ENG
RT SURR ENG
LT ENGLISH
RT ENGLISH
ML
MR
Blank Track
Blank Track
'''
def formatForMVIS(incomingString):
	splitString = string.split(incomingString, '\n')
	for i in splitString:
		if (i == 'Blank Track'):
			blankIndex = splitString.index(i)
			splitString[blankIndex] = ''

	print splitString	
	return splitString
		
		
		
#formatForMVIS(instring)	
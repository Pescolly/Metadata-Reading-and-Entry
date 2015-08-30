#!/usr/bin/python

#email module


import smtplib, smtpd

class Emailer:

	def __init__(self, server, senderlist, receiverlist, sender_name, message_text, subject):
		self.SERVER = server
		self.SENDLIST = senderlist
		self.RECEIVELIST = receiverlist
		self.SENDER_NAME = sender_name
		self.MESSAGE_TEXT = message_text
		self.SUBJECT = subject

		
	
	def sendEmail(self):
		try:
			MESSAGE = "From: %s\nTo: %s\nSubject: %s\n\n%s" % (self.SENDER_NAME, self.RECEIVELIST, self.SUBJECT, self.MESSAGE_TEXT)
			smtpObj = smtplib.SMTP(self.SERVER, 25)
			smtpObj.sendmail(self.SENDLIST, self.RECEIVELIST, MESSAGE)
			smtpObj.quit()	
			
		except Exception, e:
			print e
		
#SERVER = 'owa.modern.mvfinc.com'
#SENDLIST = ['akaramian@mvf.com']
#RECEIVELIST = ['akaramian@mvf.com']
#SENDER_NAME = "Asset Management"
#MESSAGE_TEXT = 'secret russian transmission'
#SUBJECT = 'Email Test'

#e = emailer(SERVER, SENDLIST, RECEIVELIST, SENDER_NAME, MESSAGE_TEXT, SUBJECT)
#e.sendEmail()
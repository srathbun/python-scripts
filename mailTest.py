#!/usr/bin/env python

import smtplib

SERVER = "mail.ldmailmasters.com"
FROM = "ldsupport@ldmailmasters.com"
TO = ["spencerrathbun@gmail.com"]

SUBJECT = "Test of smtp"
TEXT = "This is a test of the smtp protocol to an outside email."

# Prepare actual message
message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

%s
""" % (FROM, ", ".join(TO), SUBJECT, TEXT)

# Send the mail
server = smtplib.SMTP(SERVER)
server.sendmail(FROM, TO, message)
server.quit()

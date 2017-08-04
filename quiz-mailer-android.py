#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# -------------------------------------------------------------------------------
# Name:        quiz-mailer
# Purpose:
#
# Author:      Peter Smith
#
# Created:     17/03/2014
# Copyright:   (c) Peter Smith 2015
# Licence:     All rights reserved, Peter Smith
#-------------------------------------------------------------------------------

__author__ = 'Peter Smith'

import re
import os
import csv
import smtplib

from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders
from cStringIO import StringIO
from PyPDF2 import PdfFileWriter, PdfFileReader


#
# Some constants
#
TestMode = False      # True == don't send email
InDir = "."  # Where the data files exist
StudentFile = 'students.csv'  # the ID - email file
MasterPDF = 'Quiz export-android.pdf'  # the collection of scans

MailServer = 'mailhost.auckland.ac.nz'
EmailFrom = 'f.siedlok@auckland.ac.nz'  #
CcEmail = True  # Copy the email to the sender
EmailText = """

Here is the results of the raw scan from the recent quiz.

Remember, the scans are not checked for corrections ... you should be using a pencil if you need to make changes.

Regards"""

#
# Set up the logging
#
import logging
import logging.config

# logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
logging.config.dictConfig({
    'version': 1,
    'disable_existing_loggers': False,  # this fixes the problem
    'formatters': {
        'standard': {
            'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
        },
    },
    'handlers': {
        'default': {
            'level': 'DEBUG',
            'class': 'logging.StreamHandler',
            'formatter': 'standard',
            'stream': 'ext://sys.stdout'
        },
    },
    'loggers': {
        '': {
            'handlers': ['default'],
            'level': 'DEBUG',
            'propagate': True
        }
    }
})


def getEmail(pdfPage):
    """ get text from PDF

    :param pdfPage: a PDF object from pyPDF
    :return: EmailAddress found on the page or 0
    """

    pdf_text = pdfPage.extractText().encode('ascii', errors='replace')

    #
    # Now get the ID
    #
    #searchObj = re.search( r'Student: [a-zA-Z -]*([0-9]*)',  pdf_text)
    #searchObj = re.search( r'Student: .*([0-9]*)',  pdf_text)
    searchObj = re.search(r'([a-z]*[0-9]*@aucklanduni.ac.nz)', pdf_text)

    EmailAddress = "0"
    if searchObj:
        # if ID found get email
        EmailAddress = searchObj.group(1)
        logger.info("Found email: %s", EmailAddress)
    else:
        logger.error('Failed to find email address %s in %s', EmailAddress, pdf_text)
    return (EmailAddress)


def send_mail(send_from, send_to, subject, text, files=[]):
    assert type(send_to) == list
    assert type(files) == list

    logger.info('Sending %s to %s', files, send_to)
    emailMsg = MIMEMultipart()
    emailMsg['From'] = send_from
    if CcEmail:
        emailMsg['Cc'] = COMMASPACE.join([send_from])
    emailMsg['To'] = COMMASPACE.join(send_to)
    emailMsg['Date'] = formatdate(localtime=True)
    emailMsg['Subject'] = subject

    emailMsg.attach(MIMEText(text))

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, "rb").read())
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        emailMsg.attach(part)

        smtp = smtplib.SMTP(MailServer)

        if TestMode:
            logger.warn('Test mode - no email sent [%s]', subject)
        else:
            logger.debug('SMTP: To [%s] From [%s]', send_to, emailMsg['From'])
            smtp.sendmail(emailMsg['From'], send_to, emailMsg.as_string())
        smtp.close()


def sendPDF(pdfPage, EmailAddress):
    """ Send the pdfPage to the email address found in ehte
    :param pdfPage: a PDF object from pyPDF
    :param AUID:  the ID of the student on the page
    :param emailDictionary: a dictionary of IDs and email addresses
    :return: nothing

    """

    outfile = PdfFileWriter()
    outfile.addPage(pdfPage)
    pdfFilename = 'temp.pdf'
    with open(pdfFilename, 'wb') as f:
        outfile.write(f)                    # A temporary PDF
    #
    # Email the PDF page
    #
    # Send PDF to recipient
    send_mail(send_from=EmailFrom,
              send_to=[EmailAddress],
              subject='Quiz results for ' + EmailAddress,
              text=EmailText,
              files=[pdfFilename])
    #
    # Tidy up and delete the temporary pdf
    #
    if os.path.exists(pdfFilename):
        os.remove(pdfFilename)


def main():
    logger.info('Begin: main')

    infile = PdfFileReader(open(InDir + '/' + MasterPDF, 'rb'))

    for i in xrange(infile.getNumPages()):
        pdfPage = infile.getPage(i)
        EmailAddress = getEmail(pdfPage)

        if EmailAddress <> '0':
            sendPDF(pdfPage, EmailAddress)
        else:
            logger.error('Bad AUID on page %d', i + 1)

    logger.info('Pages processed: %d', i + 1)
    logger.info('End: main')


if __name__ == '__main__':
    main()

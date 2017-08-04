#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
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

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from io import StringIO
from smtplib import SMTP
from PyPDF2 import PdfFileReader
from PyPDF2 import PdfFileWriter

COMMASPACE= ', '
USERNAME = 'psmi045'

#
# Some constants
#
test_mode = True      # True == don't send email

# Where the data files exist
data_directory = "/home/psmith/NAS/Work/Teaching/Courses/2017/2017-Business-304/QuizResults"  
data_source_filename = 'Web-week1.pdf' # the collection of scans
tmp_pdf_filename = '/tmp/tmp-quiz-page.pdf'         # temporary pdf page storage

smtp_host = 'mailhost.auckland.ac.nz'
smtp_port = 587
authinfo_filename ='/home/psmith/.authinfo'
email_cc_to = True                                  # Copy the email to the sender
email_sent_from = 'p.smith@auckland.ac.nz'          # Senders email address
text_of_email_message = """

Here is the results of the raw scan from the recent quiz.

Remember, the scans are not checked for corrections ... you should be
using a pencil if you need to make changes.

Regards, Peter
"""

def get_authinfo_password(machine, login, port):
    s = 'machine {} port {} user {} password ([^ ]*)\n'.format(machine, port, login)
    p = re.compile(s)
    with open(authinfo_filename, 'r') as authinfo_file:
        authinfo = authinfo_file.read()
    return p.search(authinfo).group(1)


def extract_email_address(pdfPage):
    """ get text from PDF
    :param pdfPage: a PDF object from pyPDF
    :return: email_address found on in the dictionary  or 0
    """

    text_of_page = str(pdfPage.extractText())
    searchObj = re.search(r'([a-z]*[0-9]*@aucklanduni.ac.nz)', text_of_page)

    email_address = ''

    if searchObj:
        email_address = searchObj.group(1)
    else:
        print('Failed to find email address {} in {}'.format(email_address, pdf_text))

    return (email_address)


def send_email_to_student(send_from, send_to, subject, text, files=[]):
    assert type(send_to) == list
    assert type(files) == list

    emailMsg = MIMEMultipart()
    emailMsg['From'] = send_from
    if email_cc_to:
        emailMsg['Cc'] = COMMASPACE.join([send_from])
    emailMsg['To'] = COMMASPACE.join(send_to)
    emailMsg['Date'] = formatdate(localtime=True)
    emailMsg['Subject'] = subject

    emailMsg.attach(MIMEText(text))

    PASSWORD = get_authinfo_password(smtp_host, USERNAME, smtp_port)
    
    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        emailMsg.attach(part)
        #
        # Open up the connection to the mail server
        #
        conn = SMTP(host=smtp_host, port=smtp_port)
        conn.ehlo()
        conn.starttls()
        conn.login(USERNAME, PASSWORD)

        if test_mode:
            print('Test mode - no email sent to {}'.format(subject))
        else:
            print('Sending mail to [{}] from [{}]'.format(send_to, emailMsg['From']))
            try:
                smtp_error = conn.sendmail(emailMsg['From'], send_to, emailMsg.as_string())
            finally:
                print('Sendmail failed: '.foramt(smtp_error))
        conn.quit()


def send_pdf_to_student(pdfPage, email_address):
    """ Send the pdfPage to the email address found in ehte
    :param pdfPage: a PDF object from pyPDF
    :param email_address:  the email address of the student on the page
    :return: nothing

    """

    outfile = PdfFileWriter()
    outfile.addPage(pdfPage)
    with open(tmp_pdf_filename, 'wb') as f:
        outfile.write(f)                    # A temporary PDF
    #
    # Email the PDF page
    #
    # Send PDF to recipient
    send_email_to_student(send_from=email_sent_from,
              send_to=[email_address],
              subject='Quiz results for ' + email_address,
              text=text_of_email_message,
              files=[tmp_pdf_filename])
    #
    # Tidy up and delete the temporary pdf
    #
    if os.path.exists(tmp_pdf_filename):
        os.remove(tmp_pdf_filename)

        
def main():
    infile = PdfFileReader(open(data_directory + '/' + data_source_filename, 'rb'))
  
    for i in range(infile.getNumPages()):
        pdf_page = infile.getPage(i)
        email_address = extract_email_address(pdf_page)

        if email_address != '':
            send_pdf_to_student(pdf_page, email_address)
        else:
            print('Could not get email address from page {}'.format(i + 1))

    print('Pages processed: {}'.format(i + 1))


if __name__ == '__main__':    
    main()

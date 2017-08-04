# quiz-mailer
Takes a ZipGrade PDF and mails individual results to students

# Set up
You need to have a .authoinfo file with details of your remote smtp server. This file is used to retrieve your password, so It's kind of important.
For details. Only a plain .authoinfo file is support, not a .authinfo.gpg. For details see: https://www.emacswiki.org/emacs/GnusAuthinfo

# Customization
There is some customisation to be done at the front. Currently, the file that is being split and processed is hard coded. I should fix that.

# ZipGrade
When setting up students, I set the External ID to be the students email address, e.g., upi@aucklanduni.ac.nz. The regular expreseion is set up to look for athat.

So, there is some development to be done.

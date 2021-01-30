This repo includes some hacks for the great method published at [Delete only Attachments in Gmail without Deleting the Emails using Google Docs](http://techawakening.org/delete-attachments-from-gmail-without-deleting-the-emails/1842/).

Instructions:
- [make a copy of the master spreadsheet](https://docs.google.com/spreadsheet/ccc?key=0Aqy7rBwoHlSvdGVwVWQ0aEg0VGtBcWJ5MjJ1cDU4eWc&newcopy=true) as instructed in the original post
- before continuing following the instructions, open _Tools > Script editor_, replace the script by copy-pasting the one found here over it, and save it
- optionally, as visual aid, you might insert **View Thread** and **Skip?** to cells `G7` and `H7`, respectively
- you can proceed with the steps
  - when loading emails, a viewer link will be generated for each thread listed in column `G`
  - when processing emails, rows with `yes` in column `H` will be exempt of processing

Further user visible changes:
- attachments of specific mails are saved to dedicated folders (named of row index)

This repo includes some hacks for the great method published at [Delete only Attachments in Gmail without Deleting the Emails using Google Docs](http://techawakening.org/delete-attachments-from-gmail-without-deleting-the-emails/1842/).

Instructions:
- [make a copy of the master spreadsheet](https://docs.google.com/spreadsheet/ccc?key=0Aqy7rBwoHlSvdGVwVWQ0aEg0VGtBcWJ5MjJ1cDU4eWc&newcopy=true) as instructed in the original post
- before continuing following the instructions, open _Tools > Script editor_, replace the script by copy-pasting the one found here over it, and save it
- you can proceed with the steps; behavioural differences from original are as follows:
  - when loading emails, a viewer link will be generated for each thread listed in column `G` (labeled _View Threads_)
  - when processing emails, rows with `yes` in column `H` (labeled _Skip?_) will be exempt of processing
  - attachments of specific mails are saved to dedicated folders (named of row index)

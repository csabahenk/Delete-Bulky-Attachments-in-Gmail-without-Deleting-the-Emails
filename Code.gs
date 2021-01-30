/*
       ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~    
       Delete Attachments from Gmail without deleting the original message without the need of Desktop Email Clients: Techawakening.org
       ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
   Huge Attachments hoging your Gmail Inbox? Want to retain the email messages but not the attachments in it?
   Use this script to find and delete only attachments from your emails. Attachments are backed up in your Google Drive account.
 
   How it works?
   
   - It finds all emails with size above what you have mentioned.
   - Creates back up of all attachments in Google Drive
   - Creates a copy of Email body and mails you back the same.
   - Trashes the original message.
   
   Will sender's information, email time stamp be lost? 
   
   No. They will be retained within the new email copy. Details like from, to, date, cc are appended into the email body.
   
   ======================================================================================================================================
   
   For instructions, go to http://techawakening.org/?p=1842
   For queries, bugs reporting comment on the above article.

   Written by Shunmugha Sundaram for Techawakening.org - December 02, 2012

Change Log:

May-06-2015: Docslist updated to DriveApp;

*/


var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = SpreadsheetApp.getActiveSheet();
var isloaded = 0;

function processemails() {
    var size_att = sheet.getRange("D4:D4").getValues();
    var firstmsg = UserProperties.getProperty("firstmsgid");
    var isloaded = UserProperties.getProperty("isloaded");
    var threads = getthreads();
    var firstmessageId = getfirstmsgid(threads);
    spreadsheet.toast("Please wait while emails are processed to delete attachments. It could take few seconds", "Status", -1);
    if (size_att != '') {

        if (firstmsg == firstmessageId) {
            if (parseInt(isloaded)) {

                var folder = getdocfolder();
                var row = getFirstRow() + 1;

                var from, date, to, cc, removedatt, archivedata;

                if (parseInt(threads.length) != 0) {
                    for (var x = 0; x < threads.length; x++)

                    {

                        try {

                            var messages = threads[x].getMessages();

                            for (var y = 0; y < messages.length; y++) {
                                var skip = sheet.getRange(row, 8).getValues()[0][0];
                                if (skip.toLowerCase() == 'yes') {
                                    spreadsheet.toast("Row " + row + ": skipping...", "Status", 5);
                                    sheet.getRange(row, 5).setValue("Skipped");
                                    row++;
                                    continue;
                                }
                                var att = messages[y].getAttachments();
                                if (att.length > 0) {
				    // for a free account, two digits will do, the script's allotted runtime
				    // runs out approximately at 20 rows
                                    var rowfolder = folder.createFolder(itoa02(row));
                                }
                                var success = 1;
                                for (var z = 0; z < att.length; z++) {
                                    try {

                                        /*Credits: Moving Gmail attachments to Google Drive. Thanks to Amit Agarwal(@labnol)*/
                                        var file = rowfolder.createFile(att[z]);
                                        spreadsheet.toast("Row " + row + ": Successfully backed up " + file.getName(), "Status", 5);
                                        Utilities.sleep(1000);
                                    } catch (e) {
                                        spreadsheet.toast("Row " + row + ": Failed to backed up " + file.getName(), "Status", 5);

                                        sheet.getRange(row, 5).setValue("Failure");
                                        success = 0;
                                    }
                                }

                                if (success == 1) {
                                    from = "From: " + messages[y].getFrom() + "<br/>";
                                    date = "Date: " + messages[y].getDate() + "<br/>";
                                    to = "To: " + messages[y].getTo() + "<br/>";
                                    cc = "cc: " + messages[y].getCc() + "<br/>";
                                    archivedata = "Original Email Information:<br/>*****************************************************************<br/>" +
                                        from + date + to + cc +
                                        "<br/>*****************************************************************<br/>";



                                    GmailApp.sendEmail(Session.getActiveUser().getUserLoginId(),
                                        messages[y].getSubject(), "", {
                                            htmlBody: archivedata + messages[y].getBody()
                                        });
                                    spreadsheet.toast("Successfully created copy of email body along with sender information", "Status", 5);

                                    //Delete Original message once copy is made to free up space

                                    messages[y].moveToTrash();

                                    spreadsheet.toast("Trashed the original email along with attachments", "Status", 5);
                                    sheet.getRange(row, 5).setValue("Success");
                                }
                                row++;

                            }

                        } catch (err) {
                            spreadsheet.toast("Error Encountered" + err, "Status", 8);
                        }

                        if (x == threads.length - 1) {
                            spreadsheet.toast("Done. Processed all emails. See your inbox for email backup and Google Drive for attachment backup", "Status", -1);
                        }
                    }


                } else {
                    spreadsheet.toast("No emails found matching your query. Try searching with smaller size.");
                }

                UserProperties.setProperty("isloaded", 0);

            } else {
                spreadsheet.toast("Please first load emails before trying to delete attachments", "Status", -1);
            }
        } else {
            spreadsheet.toast("New emails have been found or Emails have been deleted. Updating the Sheet again", "Status", 8);
            getEmails();

        }

    } else {
        Browser.msgBox("Size field cannot be empty");
        spreadsheet.toast("Size field cannot be empty", "Status", -1);
    }
}

function getdocfolder() {
    var folderlist = DriveApp.getFolders(),
        available = 0,
        folder;

    for (var i = 0; i < folderlist.length; i++) {
        if (folderlist[i].getName() == "Gmail_Attachments") {
            available = 1;
            folder = folderlist[i];
            break;
        }
    }

    if (!available) {
        folder = DriveApp.createFolder("Gmail_Attachments");
    }

    return folder;


}


function getEmails() {

    var size_att = sheet.getRange("D4:D4").getValues();
    if (size_att != '')

    {
        sheet.getRange("A8:E108").clear();

        var threads = getthreads();
        var row = getFirstRow() + 1;
        var firstmessageId = getfirstmsgid(threads);
        UserProperties.setProperty("firstmsgid", firstmessageId);
        UserProperties.setProperty("isloaded", 1);
        spreadsheet.toast("Loading emails..Please wait. It could take few seconds", "Status", -1);
        sheet.getRange(7, 7).setValue("View Thread");
        sheet.getRange(7, 8).setValue("Skip?");
        if (parseInt(threads.length) != 0) {

            for (var i = 0; i < threads.length; i++) {

                try {
                    var messages = threads[i].getMessages();
                    for (var m = 0; m < messages.length; m++) {
                        var attname = '';
                        sheet.getRange(row, 1).setValue(messages[m].getFrom());
                        sheet.getRange(row, 2).setValue(messages[m].getSubject());
                        sheet.getRange(row, 3).setValue(messages[m].getDate());

                        var att = messages[m].getAttachments();

                        for (var z = 0; z < att.length; z++) {
                            attname = attname + att[z].getName() + ";";
                            sheet.getRange(row, 4).setValue(attname);
                        }
                        if (m == 0) {
                            sheet.getRange(row, 7).setValue('=hyperlink("' + threads[i].getPermalink() + '", "View")');
                        }
                        row++;
                    }
                } catch (error) {
                    spreadsheet.toast("Error Occured. Report it @ http://techawakening.org/", "Status", -1);
                }
                if (i == threads.length - 1) {
                    spreadsheet.toast("Successfully loaded emails.", "Status", -1);
                    spreadsheet.toast("Now select Delete Attachments from menu", "Status", -1);
                }

            }

        } else {
            spreadsheet.toast("No emails found matching your query. Try searching with smaller size.", "Status", -1);
        }
    } else {
        Browser.msgBox("Size field cannot be empty");
        spreadsheet.toast("Size field cannot be empty", "Status", -1);
    }
}

function getFirstRow() {
    var start = 7;
    return start;
}

/*
 * poor man's sprintf
 */
function itoa02(n) {
    if (n < 10) {
        return '0' + n;
    } else {
        return '' + n;
    }
}

function itoa03(n) {
    if (n < 10) {
        return '00' + n;
    } else if (n < 100) {
        return '0' + n;
    } else {
        return '' + n;
    }
}

function getfirstmsgid(threads) {
    try {
        var message = threads[0].getMessages()[0];
        var firstmessageId = message.getId();
        return firstmessageId;
    } catch (err) {}

}

function getthreads() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var size_att = sheet.getRange("D4:D4").getValues();
    try {

        var size_byte = 1048576 * parseInt(size_att);
        var threads = GmailApp.search("size:" + size_byte, 0, 50);
        return threads;
    } catch (err) {}
}

function onOpen() {
    var menuEntries = [{
        name: "Load Emails",
        functionName: "getEmails"
    }, {
        name: "Delete Attachments",
        functionName: "processemails"
    }];
    spreadsheet.addMenu("Strip Email Attachments", menuEntries);
    spreadsheet.toast("Please mention email attachments size in D4 cell and select Load Emails from Strip Email Attachments menu.", "Get Started", -1);
}

/**
 *  ----------------------------------
 *   Google Forms Email Notifications
 *  ----------------------------------
 * 
 *  Get Google Forms response in an email message
 * 
 *  license: MIT
 *  language: Google Apps Script
 *  author: Amit Agarwal
 *  email: amit@labnol.org
 *  web: https://digitalinspiration.com/
 *  blog: http://labnol.org/?p=20884
 * 
 */

/**
 * @OnlyCurrentDoc
 */

function Initialize() {

    try {

        var triggers = ScriptApp.getProjectTriggers();

        for (var i in triggers)
            ScriptApp.deleteTrigger(triggers[i]);

        ScriptApp.newTrigger("EmailGoogleForms")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onFormSubmit().create();

    } catch (error) {
        throw new Error("Please add this code in the Google Spreadsheet");
    }
}

function EmailGoogleForms(e) {

    if (!e) {
        throw new Error("Please go the Run menu and choose Initialize");
    }

    if (MailApp.getRemainingDailyQuota() > 0) {
        throw new Error("Sorry, your email quota for Gmail is over");
    }

    try {

        // Replace with your email address
        var email = "amit@labnol.org";

        // Replace with a custom subject
        var subject = "New response submitted"

        var message = [],
            ss = SpreadsheetApp.getActiveSheet(),
            cols = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];

        // Iterate through the Form Fields
        for (var keys in cols) {

            var key = cols[keys];
            var entry = e.namedValues[key] ? e.namedValues[key].toString() : "";

            // Only include form fields that are not blank
            if (entry && entry.replace(/,/g, "") !== "")
                message.push(key + ' :: ' + entry);
        }

        // Use MailApp if Gmail is disabled for your GSuite Domain
        GmailApp.sendEmail(email, subject, message.join("\n\n"), {
            cc: "cc@example.com",
            bcc: "bcc@example.com"
        });

    } catch (error) {
        Logger.log(error.toString());
    }
}
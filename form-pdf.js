/**
 *  --------------------------
 *   Google Forms PDF Document
 *  --------------------------
 * 
 *  Generate PDF documents from Google Form Responses
 * 
 *  license: MIT
 *  language: Google Apps Script
 *  author: Amit Agarwal
 *  email: amit@labnol.org
 *  web: https://digitalinspiration.com/
 * 
 */

function Initialize() {

    try {

        var triggers = ScriptApp.getProjectTriggers();

        for (var i in triggers)
            ScriptApp.deleteTrigger(triggers[i]);

        ScriptApp.newTrigger("EmailGoogleFormsPDF")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onFormSubmit().create();

    } catch (error) {
        throw new Error("Please add this code in the Google Spreadsheet");
    }
}

function EmailGoogleFormsPDF(e) {

    if (!e) {
        throw new Error("Please go the Run menu and choose Initialize");
    }

    if (MailApp.getRemainingDailyQuota() < 1) {
        throw new Error("Sorry, your email quota for Gmail is over");
    }

    try {

        // Replace with your email address
        var email = "amit@labnol.org";

        // Replace with a custom subject
        var subject = "New response submitted"

        // Replace with your own document template ID
        var templateId = "15bw3meKSE_9kv3p1aarTNfUv3qQ__0xS8cO-cAfziEQ";

        var data = {},
            message = [],
            ss = SpreadsheetApp.getActiveSheet(),
            cols = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];

        // Iterate through the Form Fields
        for (var keys in cols) {

            var key = cols[keys];
            var entry = e.namedValues[key] ? e.namedValues[key].toString() : "";

            data[key] = entry;

            // Only include form fields that are not blank
            if (entry && entry.replace(/,/g, "") !== "")
                message.push(key + ' :: ' + entry);

        }

        var template = DriveApp.getFileById(templateId);

        var targetFile = template.makeCopy("GoogleForms.pdf");
        var targetDocument = DocumentApp.openById(targetFile.getId());
        var targetBody = targetDocument.getBody();

        Object.keys(data).forEach(function (key) {
            var searchPattern = "<<" + key + ">>";
            targetBody.replaceText(searchPattern, data[key]);
        });

        targetDocument.saveAndClose();

        var blob = targetDocument.getAs("application/pdf").copyBlob();

        // Use MailApp if Gmail is disabled for your GSuite Domain
        GmailApp.sendEmail(email, subject, message.join("\n\n"), {
            attachments: [blob]
        });

    } catch (error) {
        Logger.log(error.toString());
    }
}
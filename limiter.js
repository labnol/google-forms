/**
 *  ----------------------
 *   Google Forms Limiter
 *  ----------------------
 * 
 *  Set response limits for Google Forms and close form automatically.
 * 
 *  license: MIT
 *  language: Google Apps Script
 *  author: Amit Agarwal
 *  email: amit@labnol.org
 *  web: https://digitalinspiration.com/
 * 
 */


/* Set the form’s open and close dates in YYYY-MM-DD HH:MM format. 
If you would rather not have an open or close date, 
just set the corresponding value to blank. */

FORM_OPEN_DATE = "2019-01-01 08:00";
FORM_CLOSE_DATE = "2019-12-25 23:30";

/* Set the RESPONSE_COUNT equal to the total number of entries 
that you would like to receive after which the form is closed automatically. 
If you would not like to set a limit, set this value to blank. */

RESPONSE_COUNT = "100";

/* Initialize the form, setup time based triggers */
function Initialize() {

    deleteTriggers_();

    if ((FORM_OPEN_DATE !== "") &&
        ((new Date()).getTime() < parseDate_(FORM_OPEN_DATE).getTime())) {
        closeForm();
        ScriptApp.newTrigger("openForm")
            .timeBased()
            .at(parseDate_(FORM_OPEN_DATE))
            .create();
    }

    if (FORM_CLOSE_DATE !== "") {
        ScriptApp.newTrigger("closeForm")
            .timeBased()
            .at(parseDate_(FORM_CLOSE_DATE))
            .create();
    }

    if (RESPONSE_COUNT !== "") {
        ScriptApp.newTrigger("checkLimit")
            .forForm(FormApp.getActiveForm())
            .onFormSubmit()
            .create();
    }

}

/* Delete all existing Script Triggers */
function deleteTriggers_() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i in triggers) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
}

/* Send a mail to the form owner when the form status changes */
function informUser_(subject) {
    var formURL = FormApp.getActiveForm().getPublishedUrl();
    MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, formURL);
}

/* Allow Google Form to Accept Responses */
function openForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(true);
    informUser_("Your Google Form is now accepting responses");
}

/* Close the Google Form, Stop Accepting Reponses */
function closeForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(false);
    deleteTriggers_();
    informUser_("Your Google Form is no longer accepting responses");
}

/* If Total # of Form Responses >= Limit, Close Form */
function checkLimit() {
    if (FormApp.getActiveForm().getResponses().length >= RESPONSE_COUNT) {
        closeForm();
    }
}

/* Parse the Date for creating Time-Based Triggers */
function parseDate_(d) {
    return new Date(d.substr(0, 4), d.substr(5, 2) - 1,
        d.substr(8, 2), d.substr(11, 2), d.substr(14, 2));
}
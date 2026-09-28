/**
 * SURVEY MAIL MERGE — sends a personalized email invite to everyone in a
 * Google Sheet contact list, using your own Gmail account.
 * ---------------------------------------------------------------------
 * SETUP:
 * 1. Make a new Google Sheet named exactly: Survey Contact List
 *    with two columns, starting in row 1 as headers:
 *       A: Name        B: Email
 *    (add as many contact rows below as you like)
 * 2. Go to https://script.google.com -> New project -> paste this file in
 *    (as a SECOND file if you already pasted create_google_form.gs, or
 *    its own separate project - either works).
 * 3. Edit the FORM_LINK constant below with your live form URL.
 * 4. Select function sendSurveyInvites from the dropdown -> Run.
 * 5. Authorize when prompted (same one-time step as before).
 * 6. Check "Sent" column added to your sheet to confirm delivery, and
 *    check the Execution log for any errors (e.g. malformed email address).
 *
 * QUOTA NOTE: personal Gmail accounts can send ~100 emails/day through
 * Apps Script; Google Workspace (college/work) accounts get ~1500/day.
 * If your contact list is bigger than that, run it again the next day -
 * the script skips anyone already marked "Sent" so it's safe to re-run.
 */

const FORM_LINK = 'PASTE_YOUR_LIVE_FORM_URL_HERE';
const SHEET_NAME = 'Survey Contact List';

function sendSurveyInvites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  // Ensure there's a "Sent" header in column C
  if (data[0][2] !== 'Sent') {
    sheet.getRange(1, 3).setValue('Sent');
  }

  let sentCount = 0;
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    const alreadySent = data[i][2];

    if (!email || alreadySent === 'Yes') continue;

    const firstName = String(name || 'there').split(' ')[0];
    const subject = 'Quick 5-min survey - magazine reading habits (MBA research)';
    const body =
      'Hi ' + firstName + ',\n\n' +
      'I\'m working on an MBA business analytics project studying reader loyalty and ' +
      'subscription habits for Indian news & business magazines (India Today, Outlook, ' +
      'Business Today, The Week, Forbes India, and similar).\n\n' +
      'If you\'ve ever read any of these - even just occasionally - I\'d be really ' +
      'grateful if you could fill out this short, anonymous survey (4-5 minutes):\n\n' +
      FORM_LINK + '\n\n' +
      'It would also help a lot if you could forward this to a couple of people you ' +
      'know who read the news!\n\n' +
      'Thanks so much,\nAishwarya';

    try {
      GmailApp.sendEmail(email, subject, body);
      sheet.getRange(i + 1, 3).setValue('Yes');
      sentCount++;
    } catch (e) {
      sheet.getRange(i + 1, 3).setValue('ERROR: ' + e.message);
    }
  }

  Logger.log('Sent ' + sentCount + ' new invites.');
}

/**
 * Optional: run this once a week to nudge anyone who hasn't responded yet.
 * It reuses the same "Sent" flag logic but with a follow-up message - only
 * fill this in and run it if you're comfortable with one reminder email;
 * do not run it more than once per contact.
 */
function sendFollowUpReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  if (data[0][3] !== 'Reminder Sent') {
    sheet.getRange(1, 4).setValue('Reminder Sent');
  }

  let sentCount = 0;
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    const invited = data[i][2];
    const reminded = data[i][3];

    if (!email || invited !== 'Yes' || reminded === 'Yes') continue;

    const firstName = String(name || 'there').split(' ')[0];
    const subject = 'Quick reminder - 5-min magazine survey';
    const body =
      'Hi ' + firstName + ',\n\n' +
      'Just a gentle nudge in case my earlier email got buried! Still trying to ' +
      'collect a few more responses for my MBA research survey on Indian magazine ' +
      'reading habits: \n\n' + FORM_LINK + '\n\n' +
      'Takes about 4-5 minutes. Thanks so much if you get a chance!\n\nAishwarya';

    try {
      GmailApp.sendEmail(email, subject, body);
      sheet.getRange(i + 1, 4).setValue('Yes');
      sentCount++;
    } catch (e) {
      sheet.getRange(i + 1, 4).setValue('ERROR: ' + e.message);
    }
  }

  Logger.log('Sent ' + sentCount + ' reminders.');
}

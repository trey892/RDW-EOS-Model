/**
 * Auto-saves report attachments from Gmail labels into their matching Drive folders,
 * then marks each message so it's never saved twice. Handles multiple sources (RTA,
 * McLeod, etc.) in one script/one trigger.
 *
 * UPDATING FROM THE RTA-ONLY VERSION:
 * Just replace all the code in your existing "RTA Reports to Drive" Apps Script
 * project with this file and save (Ctrl+S). Your existing hourly/daily trigger keeps
 * working -- no need to recreate it.
 *
 * FIRST-TIME SETUP (if you don't already have the project):
 * 1. https://script.google.com -> New project. Paste this file in, save.
 * 2. Run `saveReportAttachments` once manually to trigger the permission prompt --
 *    approve Gmail + Drive access.
 * 3. Left sidebar -> Triggers -> Add Trigger -> function `saveReportAttachments`,
 *    Time-driven, Hour timer (or Day timer). Save.
 *
 * ADDING A NEW SOURCE LATER: just add another entry to the SOURCES array below.
 *
 * extraQuery (optional): an additional Gmail search clause ANDed onto the label
 * search. Use this when a label is shared with other mail (like the existing
 * "Reports" label with ~3,600 unrelated messages in it) so the script only ever
 * touches the specific sender/subject you care about -- never the backlog.
 */

const SOURCES = [
  {
    sourceLabel: "RTA Reports",
    processedLabel: "RTA Reports/Saved",
    driveFolderId: "1YlR0efDJstVmmpInbyG6KCIuxO-I1IDF", // RTA Reports Inbox
    extraQuery: null,
  },
  {
    // McLeod automated reports -- shares the "Reports" label with ~3,600 other
    // (human-sent) messages, so extraQuery scopes this to the automated sender only.
    // Sender-based on purpose, not an exact-subject allowlist: a subject allowlist
    // silently drops any new report type McLeod starts sending (this is exactly how
    // "Unbilled Orders" got missed for weeks) -- matching the sender means every
    // current and future automated McLeod report is picked up with no code change.
    sourceLabel: "3Operations Managment/Reports",
    processedLabel: "3Operations Managment/Reports/McLeod-Saved",
    driveFolderId: "1oGIyA_0U6XLYjmv3mVadsH3xIYbkVh5k", // McLeod Reports Inbox
    extraQuery: 'newer_than:14d from:noreply@robbiedwood.com',
  },
];

function saveReportAttachments() {
  SOURCES.forEach(function (source) {
    saveForSource_(source.sourceLabel, source.processedLabel, source.driveFolderId, source.extraQuery);
  });
}

function saveForSource_(sourceLabelName, processedLabelName, driveFolderId, extraQuery) {
  const sourceLabel = GmailApp.getUserLabelByName(sourceLabelName);
  if (!sourceLabel) {
    Logger.log('Label "%s" not found -- check the name matches exactly (case-sensitive).', sourceLabelName);
    return;
  }
  const processedLabel = getOrCreateLabel_(processedLabelName);
  const folder = DriveApp.getFolderById(driveFolderId);

  let query = 'label:"' + sourceLabelName + '" -label:"' + processedLabelName + '" has:attachment';
  if (extraQuery) {
    query += " " + extraQuery;
  }
  const threads = GmailApp.search(query);

  let savedCount = 0;
  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    messages.forEach(function (message) {
      const attachments = message.getAttachments();
      if (attachments.length === 0) return;

      const dateStr = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
      attachments.forEach(function (att) {
        const cleanName = dateStr + "_" + att.getName();
        folder.createFile(att.copyBlob()).setName(cleanName);
        savedCount++;
      });
    });
    thread.addLabel(processedLabel);
  });

  Logger.log('[%s] Saved %s attachment(s) from %s thread(s).', sourceLabelName, savedCount, threads.length);
}

function getOrCreateLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

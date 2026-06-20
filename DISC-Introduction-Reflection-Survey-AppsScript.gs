/**
 * Kristy — DISC Introduction: Reflection Survey
 * BUILD: kristy-doc-v7
 *
 * Gmail: Settings → Send mail as → surveys@discconnector.com (verified + default)
 * Services: Gmail API enabled
 * After changes: Deploy → Manage deployments → Edit → New version → Deploy
 */

var CONFIG = {
  SHEET_ID: '13512j8bdu6ON1zkYb1S-3dUlSwWDJHtTw5vlnEEPEvw',
  DOC_ID: '1rA4IHuNNZ0Lus8a7V7Raf7w799er5l0mn42MP-2COnA',
  LOGO_FILE_ID: 'PASTE_LOGO_FILE_ID',
  FROM_EMAIL: 'surveys@discconnector.com',
  FROM_NAME: 'DISC Connector',
  NOTIFY_EMAIL: 'surveys@discconnector.com',
  SURVEY_TITLE: 'DISC Introduction: Reflection Survey',
  DISCLAIMER: 'This document is used for coaching and training purposes only and is confidential.'
};

var NUM_QUESTIONS = 15;

var STYLE = {
  FONT: 'Poppins',
  FONT_LIGHT: 'Poppins Light',
  SIZE: 12,
  TITLE_SIZE: 16,
  TITLE_COLOR: '#660099',
  NAME_COLOR: '#330044',
  BODY_COLOR: '#333333',
  MUTED_COLOR: '#666666'
};

var QUESTIONS = [
  'What does DISC measure?',
  'The 4-P Model from DISC, the D refers to how you deal with...',
  'The 4-P Model from DISC, the I refers to how you deal with...',
  'The 4-P Model from DISC, the S refers to how you deal with...',
  'The 4-P Model from DISC, the C refers to how you deal with...',
  'Your Natural Style in DISC describes...',
  'Your Adapted Style in DISC describes...',
  'Which behavioral dimension is your highest score above the energy line from your natural graph from your DISC report?',
  'For your highest DISC dimension (D, I, S, or C), the highest DISC dimension closet to 100% above the energy line; what strengths or advantages do you see in having this as your natural style? How do you believe these strengths can support and strengthen your team\'s collaboration?',
  'Your "co-pilot", the second-highest DISC dimension (D, I, S, or C), the second closet to 100% above the energy line; how do you believe that impacts or influences your Highest DISC dimension?',
  'For your highest DISC dimension (D, I, S, or C), what red lights or potential downsides do you see in having this as your natural style? In what ways might these tendencies hinder or complicate your team\'s collaboration if you are not managing them well?',
  'For your lowest DISC dimension (D, I, S, or C), the one closest to 0% below the energy line; what strengths or advantages do you see in this part of your style? How could these strengths support and strengthen your team\'s collaboration?',
  'Still referring to your lowest DISC dimension, how might others perceive you in terms of team collaboration? Are there any possible misperceptions that could arise because of this lower score?',
  'When you compare your Natural style to your Adapted style, which DISC dimension shows the biggest gap? In that area, do you feel like you are burning more fuel to maintain your Adapted style? If so, how might this gap drain your energy or "empty your tank" more quickly during a typical week? What specific habits or boundaries could help you keep your tank fuller while still adapting when you need to?',
  'From your DISC coaching sessions, is there anything that you specifically would like to address or get direction from your coach that has not already been discussed?'
];

/** Multiple-choice answer options for Q1–Q8 (null = open-ended question). */
var QUESTION_OPTIONS = [
  ['Emotional Intelligence', 'Behavioral Intelligence', 'Technical Skills', 'Observable Behavior'],
  ['Pace', 'People', 'Procedure', 'Problems'],
  ['People', 'Pace', 'Procedure', 'Problems'],
  ['Procedure', 'Pace', 'People', 'Problems'],
  ['Problems', 'People', 'Pace', 'Procedure'],
  ['How I show up at work.', 'How others perceive me at work.', 'Is your default, most comfortable way of behaving.', 'How you naturally behave when stressed.'],
  ['How I show up at work', 'How others perceive me at work.', 'My default, personality traits', 'How I act at home.'],
  ['D', 'I', 'S', 'C'],
  null, null, null, null, null, null, null
];

var MC_LETTERS = ['A', 'B', 'C', 'D'];

var RESPONSES_MARKER = '=== RESPONSES (NEWEST FIRST) ===';

function styleText(textElement, opts) {
  opts = opts || {};
  textElement.setFontFamily(opts.font || STYLE.FONT);
  textElement.setFontSize(opts.size || STYLE.SIZE);
  textElement.setBold(opts.bold === true);
  textElement.setItalic(opts.italic === true);
  if (opts.color) textElement.setForegroundColor(opts.color);
}

function setDocDefaultFont(body) {
  var attrs = {};
  attrs[DocumentApp.Attribute.FONT_FAMILY] = STYLE.FONT;
  attrs[DocumentApp.Attribute.FONT_SIZE] = STYLE.SIZE;
  attrs[DocumentApp.Attribute.BOLD] = false;
  body.setAttributes(attrs);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.parameter.payload);
    var email = String(data.participant_email || '').trim();
    if (!isValidEmail(email)) {
      return ContentService.createTextOutput('Invalid email').setMimeType(ContentService.MimeType.TEXT);
    }
    var row = buildRowFromPayload(data);
    var surveyId = String(data.survey_id || 'Session 1').trim();
    writeToSheet(row);
    try {
      writeSubmissionToDoc(row);
    } catch (docErr) {
      Logger.log('Doc write error: ' + docErr.message + '\n' + docErr.stack);
    }
    sendFromBusiness(
      CONFIG.NOTIFY_EMAIL,
      'New Survey Response: ' + row[1] + (row[2] ? ' | ' + row[2] : ''),
      buildFormattedEmailBody(row, surveyId, false)
    );
    sendFromBusiness(
      email,
      'Your DISC Introduction: Reflection Survey Responses',
      buildParticipantEmailBody(row, surveyId)
    );
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function buildRowFromPayload(data) {
  var row = [
    new Date(),
    String(data.participant_name || '').trim(),
    String(data.participant_org || '').trim(),
    String(data.participant_email || '').trim()
  ];
  for (var i = 1; i <= NUM_QUESTIONS; i++) {
    row.push(String(data['q' + i] || '').trim());
  }
  return row;
}

function writeToSheet(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Responses') || ss.getSheets()[0];
  var colCount = 4 + NUM_QUESTIONS;
  if (sheet.getLastRow() === 0) setupSheetHeaders(sheet);
  var values = [[
    formatTimestamp(row[0]),
    row[1],
    row[2],
    row[3]
  ].concat(row.slice(4))];
  sheet.insertRowBefore(2);
  var dataRange = sheet.getRange(2, 1, 1, values[0].length);
  dataRange.setValues(values);
  dataRange.setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontWeight('normal');
}

/** Run once to fix purple rows on existing data rows. */
function repairSheetFormatting() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Responses') || ss.getSheets()[0];
  var colCount = 4 + NUM_QUESTIONS;
  var lastRow = sheet.getLastRow();
  setupSheetHeaders(sheet);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow, colCount)
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal');
  }
  Logger.log('Sheet repaired: row 1 purple header only, data rows white.');
}

function setupSheetHeaders(sheet) {
  var colCount = 4 + NUM_QUESTIONS;
  var headers = ['Timestamp', 'Name', 'Organization', 'Email'];
  for (var i = 1; i <= NUM_QUESTIONS; i++) headers.push('Q' + i);
  sheet.getRange(1, 1, 1, colCount).setValues([headers]);
  sheet.getRange(1, 1, 1, colCount).setBackground('#660099').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function writeSubmissionToDoc(row) {
  if (!CONFIG.DOC_ID || CONFIG.DOC_ID.indexOf('PASTE_') !== -1) {
    throw new Error('DOC_ID is not set in CONFIG. Paste your Google Doc ID and save.');
  }
  var doc = DocumentApp.openById(CONFIG.DOC_ID);
  setupDocHeaderFooter(doc);
  var body = doc.getBody();
  if (body.getText().indexOf(RESPONSES_MARKER) === -1) {
    buildDocBody(body);
  }
  insertSubmissionAtTop(body, row);
  doc.saveAndClose();
}

function insertSubmissionAtTop(body, row) {
  var markerIndex = findParagraphIndex(body, RESPONSES_MARKER);
  if (markerIndex === -1) markerIndex = body.getNumChildren() - 1;
  var insertAt = markerIndex + 1;

  var divider = body.insertParagraph(insertAt, '────────────────────────────────────────────────────────');
  divider.setSpacingBefore(16).setSpacingAfter(16);
  styleText(divider.editAsText(), { size: 10, color: '#cccccc' });

  var answers = row.slice(4);
  var name = row[1] || 'Unknown';
  var org = row[2] || '';
  var email = row[3] || '';
  var ts = formatDocTimestamp(row[0]);

  for (var i = answers.length - 1; i >= 0; i--) {
    var ans = answers[i] || '(no response)';
    var isMc = QUESTION_OPTIONS[i] != null;

    if (isMc) {
      var aPara = body.insertParagraph(insertAt, formatMcAnswer(i, ans));
      aPara.setSpacingBefore(0).setSpacingAfter(12);
      styleText(aPara.editAsText(), { color: STYLE.BODY_COLOR });
      insertMcOptionsBlock(body, insertAt, i, ans);
    } else {
      var openPara = body.insertParagraph(insertAt, ans);
      openPara.setSpacingBefore(0).setSpacingAfter(12);
      styleText(openPara.editAsText(), { color: STYLE.BODY_COLOR });
    }

    var label = 'Q' + (i + 1) + '. ' + QUESTIONS[i];
    var qPara = body.insertParagraph(insertAt, label);
    qPara.setSpacingBefore(10).setSpacingAfter(4);
    styleText(qPara.editAsText(), { color: STYLE.TITLE_COLOR });
  }

  var submitted = body.insertParagraph(insertAt, 'Submitted: ' + ts);
  submitted.setSpacingBefore(4).setSpacingAfter(12);
  styleText(submitted.editAsText(), { color: STYLE.MUTED_COLOR });

  if (org) {
    var orgP = body.insertParagraph(insertAt, 'Organization: ' + org);
    orgP.setSpacingBefore(2).setSpacingAfter(2);
    styleText(orgP.editAsText(), { color: STYLE.NAME_COLOR });
  }

  var emailP = body.insertParagraph(insertAt, 'Email: ' + email);
  emailP.setSpacingBefore(2).setSpacingAfter(2);
  styleText(emailP.editAsText(), { color: STYLE.NAME_COLOR });

  var nameP = body.insertParagraph(insertAt, 'Name: ' + name);
  nameP.setSpacingBefore(20).setSpacingAfter(2);
  styleText(nameP.editAsText(), { color: STYLE.NAME_COLOR, size: STYLE.SIZE });
}

function isMcQuestion(qIndex) {
  return QUESTION_OPTIONS[qIndex] != null;
}

function formatMcAnswer(qIndex, selected) {
  var opts = QUESTION_OPTIONS[qIndex];
  if (!opts) return selected || '(no response)';
  var sel = String(selected || '').trim().toUpperCase();
  var idx = MC_LETTERS.indexOf(sel);
  if (idx === -1) return selected || '(no response)';
  return 'Answer: ' + MC_LETTERS[idx] + '. ' + opts[idx];
}

function insertMcOptionsBlock(body, insertAt, qIndex, selected) {
  var opts = QUESTION_OPTIONS[qIndex];
  if (!opts) return;
  var sel = String(selected || '').trim().toUpperCase();

  for (var o = opts.length - 1; o >= 0; o--) {
    var line = MC_LETTERS[o] + '. ' + opts[o];
    var optPara = body.insertParagraph(insertAt, line);
    optPara.setIndentStart(18).setSpacingBefore(0).setSpacingAfter(2);
    styleText(optPara.editAsText(), {
      color: sel === MC_LETTERS[o] ? STYLE.TITLE_COLOR : STYLE.MUTED_COLOR,
      bold: sel === MC_LETTERS[o]
    });
  }

  var hdr = body.insertParagraph(insertAt, 'Options:');
  hdr.setSpacingBefore(6).setSpacingAfter(2);
  styleText(hdr.editAsText(), { color: STYLE.MUTED_COLOR, size: 10 });
}

function buildFormattedEmailBody(row, surveyId, includeGreeting) {
  var name = row[1] || 'Unknown';
  var org = row[2] || '';
  var answers = row.slice(4);
  var body = '';

  if (includeGreeting) {
    body += 'Hello ' + name + ',\n\n';
    body += 'Thank you for completing the ' + CONFIG.SURVEY_TITLE + '. ';
    body += 'Below is a copy of your responses for your records. Your coach will be in touch with you shortly.\n\n';
  }

  body += CONFIG.SURVEY_TITLE + '\n';
  body += '========================================\n\n';
  body += 'Name: ' + name + '\n';
  if (org) body += 'Organization: ' + org + '\n';
  body += 'Survey: ' + surveyId + '\n';
  body += 'Submitted: ' + formatTimestamp(row[0]) + '\n\n';
  body += '--- RESPONSES ---\n\n';

  for (var i = 0; i < answers.length; i++) {
    body += 'Q' + (i + 1) + ': ' + (answers[i] || '(no response)') + '\n';
  }

  body += '\n' + CONFIG.DISCLAIMER + '\n';
  body += '\u00A9 2026 DISC-Connector. All Rights Reserved. | discconnector.com\n';
  return body;
}

function buildParticipantEmailBody(row, surveyId) {
  return buildFormattedEmailBody(row, surveyId, true);
}

function sendFromBusiness(to, subject, body) {
  var from = CONFIG.FROM_EMAIL;
  var name = CONFIG.FROM_NAME;
  var aliases = GmailApp.getAliases();

  if (aliases.indexOf(from) === -1) {
    throw new Error(
      'Gmail alias missing: ' + from + '. Found aliases: ' + aliases.join(', ')
    );
  }

  GmailApp.sendEmail(to, subject, body, {
    from: from,
    name: name,
    replyTo: from
  });
}

function setupAll() {
  setupSheet();
  setupDoc();
  Logger.log('Setup complete.');
}

function refreshHeaderFooter() {
  applyAllDocFormatting();
}

/** Run this once after pasting kristy-doc-v4 — updates logo, title, and all existing text. */
function applyAllDocFormatting() {
  reformatDoc();
}

function reformatDoc() {
  var doc = DocumentApp.openById(CONFIG.DOC_ID);
  setupDocHeaderFooter(doc);
  var body = doc.getBody();
  setDocDefaultFont(body);
  ensureDocIntro(body);
  restyleDocumentBody(body);
  doc.saveAndClose();
  Logger.log('BUILD kristy-doc-v7: logo (top-left), title 16pt, body 12pt Poppins applied.');
}

function setupSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Responses') || ss.getSheets()[0];
  setupSheetHeaders(sheet);
}

function setupDoc() {
  var doc = DocumentApp.openById(CONFIG.DOC_ID);
  setupDocHeaderFooter(doc);
  var body = doc.getBody();
  body.clear();
  buildDocBody(body);
  doc.saveAndClose();
}

function getLogoFileId() {
  if (CONFIG.LOGO_FILE_ID && CONFIG.LOGO_FILE_ID.indexOf('PASTE_') === -1) {
    return CONFIG.LOGO_FILE_ID;
  }
  try {
    var docFile = DriveApp.getFileById(CONFIG.DOC_ID);
    var parents = docFile.getParents();
    if (!parents.hasNext()) return null;
    var folder = parents.next();
    var names = [
      'DISCconnector_Brand_Header.png',
      'DISCconnector_Brand_Header.jpg',
      'DISC-Connector_logo.png',
      'DISCconnector_logo.png'
    ];
    for (var n = 0; n < names.length; n++) {
      var files = folder.getFilesByName(names[n]);
      if (files.hasNext()) return files.next().getId();
    }
    var all = folder.getFiles();
    while (all.hasNext()) {
      var f = all.next();
      var fileName = f.getName().toLowerCase();
      if (fileName.indexOf('disc') !== -1 && /\.png$|\.jpg$|\.jpeg$/i.test(fileName)) {
        return f.getId();
      }
    }
  } catch (err) {
    Logger.log('Could not auto-find logo: ' + err.message);
  }
  return null;
}

function setupDocHeaderFooter(doc) {
  var header = doc.getHeader() || doc.addHeader();
  header.clear();
  try {
    var logoId = getLogoFileId();
    if (!logoId) throw new Error('DISCconnector_Brand_Header.png not found in same Drive folder as Google Doc');
    var logo = DriveApp.getFileById(logoId).getBlob();
    var hp = header.appendParagraph('');
    hp.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    hp.setIndentStart(0);
    hp.setIndentFirstLine(0);
    hp.setSpacingBefore(8);
    hp.setSpacingAfter(4);
    var img = hp.appendInlineImage(logo);
    var w = img.getWidth();
    var h = img.getHeight();
    var maxW = 580;
    if (w > maxW) {
      img.setWidth(maxW);
      img.setHeight(Math.round(h * (maxW / w)));
    }
    header.appendParagraph('');
  } catch (err) {
    var brand = header.appendParagraph('DISC-Connector');
    brand.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    styleText(brand.editAsText(), { size: STYLE.SIZE, color: STYLE.TITLE_COLOR });
    Logger.log('Logo not loaded: ' + err.message);
  }

  var footer = doc.getFooter() || doc.addFooter();
  footer.clear();

  var copyright = footer.appendParagraph('\u00A9 2026 DISC-Connector. All Rights Reserved.');
  copyright.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  styleText(copyright.editAsText(), { size: 9, color: '#888888' });

  var disclaimer = footer.appendParagraph(CONFIG.DISCLAIMER);
  disclaimer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  styleText(disclaimer.editAsText(), { size: 9, color: '#888888' });
}

function buildDocBody(body) {
  var titlePara = body.appendParagraph(CONFIG.SURVEY_TITLE + ' — Responses');
  titlePara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  styleText(titlePara.editAsText(), { color: STYLE.TITLE_COLOR, size: STYLE.TITLE_SIZE });

  var intro = body.appendParagraph('Newest responses appear first below.');
  styleText(intro.editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
  body.appendParagraph('');

  var rMarker = body.appendParagraph(RESPONSES_MARKER);
  styleText(rMarker.editAsText(), { size: 1, color: '#ffffff' });
  body.appendParagraph('');
}

function ensureDocIntro(body) {
  if (findParagraphIndex(body, RESPONSES_MARKER) === -1) {
    buildDocBody(body);
    return;
  }
  var titleText = CONFIG.SURVEY_TITLE + ' — Responses';
  var hasTitle = false;
  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var p = child.asParagraph();
    if (p.getText().indexOf(titleText) !== -1) {
      hasTitle = true;
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      styleText(p.editAsText(), { color: STYLE.TITLE_COLOR, size: STYLE.TITLE_SIZE });
      break;
    }
  }
  if (!hasTitle) {
    var markerIndex = findParagraphIndex(body, RESPONSES_MARKER);
    var titlePara = body.insertParagraph(markerIndex, titleText);
    titlePara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    styleText(titlePara.editAsText(), { color: STYLE.TITLE_COLOR, size: STYLE.TITLE_SIZE });
    var introPara = body.insertParagraph(markerIndex + 1, 'Newest responses appear first below.');
    styleText(introPara.editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
  }
  for (var j = 0; j < body.getNumChildren(); j++) {
    var introChild = body.getChild(j);
    if (introChild.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    if (introChild.asParagraph().getText() === 'Newest responses appear first below.') {
      styleText(introChild.asParagraph().editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
      break;
    }
  }
}

function restyleDocumentBody(body) {
  var markerIndex = findParagraphIndex(body, RESPONSES_MARKER);
  var titleText = CONFIG.SURVEY_TITLE + ' — Responses';

  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var p = child.asParagraph();
    var t = p.getText();
    if (!t || t.indexOf(RESPONSES_MARKER) !== -1) continue;

    if (t.indexOf('────────────────') !== -1) {
      styleText(p.editAsText(), { size: 10, color: '#cccccc' });
      continue;
    }
    if (t.indexOf(titleText) !== -1) {
      p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      styleText(p.editAsText(), { color: STYLE.TITLE_COLOR, size: STYLE.TITLE_SIZE });
      continue;
    }
    if (t === 'Newest responses appear first below.') {
      styleText(p.editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
      continue;
    }
    if (markerIndex === -1 || i <= markerIndex) continue;

    if (t.indexOf('Name:') === 0 || t.indexOf('Email:') === 0 || t.indexOf('Organization:') === 0) {
      styleText(p.editAsText(), { color: STYLE.NAME_COLOR, size: STYLE.SIZE });
    } else if (t.indexOf('Submitted:') === 0) {
      styleText(p.editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
    } else if (t === 'Options:') {
      styleText(p.editAsText(), { color: STYLE.MUTED_COLOR, size: 10 });
    } else if (t.indexOf('Answer:') === 0) {
      styleText(p.editAsText(), { color: STYLE.BODY_COLOR, size: STYLE.SIZE });
    } else if (/^[A-D]\.\s/.test(t)) {
      styleText(p.editAsText(), { color: STYLE.MUTED_COLOR, size: STYLE.SIZE });
    } else if (/^Q\d+\.\s/.test(t)) {
      styleText(p.editAsText(), { color: STYLE.TITLE_COLOR, size: STYLE.SIZE });
    } else {
      styleText(p.editAsText(), { color: STYLE.BODY_COLOR, size: STYLE.SIZE });
    }
  }
}

function findParagraphIndex(body, text) {
  for (var i = 0; i < body.getNumChildren(); i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      if (child.asParagraph().getText().indexOf(text) !== -1) return i;
    }
  }
  return -1;
}

/** Rebuild Google Doc from whatever is currently in the spreadsheet (newest first). */
function rebuildDocFromSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Responses') || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var colCount = 4 + NUM_QUESTIONS;

  var doc = DocumentApp.openById(CONFIG.DOC_ID);
  setupDocHeaderFooter(doc);
  var body = doc.getBody();
  body.clear();
  setDocDefaultFont(body);
  buildDocBody(body);

  var count = 0;
  if (lastRow > 1) {
    var numDataRows = lastRow - 1;
    var data = sheet.getRange(2, 1, numDataRows, colCount).getValues();
    for (var r = 0; r < data.length; r++) {
      insertSubmissionAtTop(body, rowFromSheetValues(data[r]));
      count++;
    }
  }

  ensureDocIntro(body);
  restyleDocumentBody(body);
  doc.saveAndClose();
  Logger.log('Doc rebuilt from sheet: ' + count + ' response(s).');
}

function rowFromSheetValues(values) {
  var ts = values[0];
  if (ts && !(ts instanceof Date)) {
    var parsed = new Date(ts);
    if (!isNaN(parsed.getTime())) ts = parsed;
  }
  var row = [ts, values[1], values[2], values[3]];
  for (var i = 4; i < 4 + NUM_QUESTIONS; i++) {
    row.push(values[i] != null ? String(values[i]) : '');
  }
  return row;
}

function diagnoseEmailSetup() {
  Logger.log('BUILD: kristy-doc-v7');
  Logger.log('DOC_ID: ' + CONFIG.DOC_ID);
  Logger.log('Active user: ' + Session.getActiveUser().getEmail());
  Logger.log('Aliases: ' + GmailApp.getAliases().join(', '));
}

/** Run once after pasting DOC_ID — formats the Google Doc with logo and title. */
function testDocSetup() {
  setupDoc();
  Logger.log('Doc setup complete. Open your Google Doc to confirm logo and title.');
}

/** Run once to test writing a sample response into the Google Doc. */
function testWriteToDoc() {
  var sample = [
    new Date(),
    'Doc Test User',
    'Test Org',
    'test@example.com',
    'A', 'B', 'C', 'D', 'A', 'B', 'C',
    'Sample open answer',
    'Another sample',
    'Third sample',
    'Fourth sample',
    'Fifth sample',
    'Final sample'
  ];
  writeSubmissionToDoc(sample);
  Logger.log('Sample response written to Google Doc.');
}

function testSendAsAlias() {
  var to = Session.getActiveUser().getEmail() || 'apgblue3@gmail.com';
  sendFromBusiness(
    to,
    'Kristy survey test — kristy-doc-v1',
    'If From shows surveys@discconnector.com, send-as is working.'
  );
  Logger.log('Test sent to ' + to);
}

function formatTimestamp(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d/yyyy, h:mm:ss a');
  }
  return String(value || '');
}

function formatDocTimestamp(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
  }
  return String(value || '');
}

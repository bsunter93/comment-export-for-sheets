/**
 * Comment Export for Sheets: the add-on surface.
 *
 * One button. The product is the tab it writes, not this panel, so nothing here tries to
 * render comments. Sheets already sorts, filters and exports better than a sidebar would,
 * and the user already knows how to drive it.
 *
 * Access model: drive.file only, granted per file by the user through
 * requestFileScopeForActiveDocument. Spike 1 established that neither the comments API
 * anchor nor an implicit container grant gets us there, so the grant is explicit and the
 * scope stays non-sensitive, which keeps this out of restricted-scope security review.
 */

function onHomepage(e) {
  return hasFileAccess(e) ? mainCard() : grantCard();
}

/** Fired after the user grants file access, so we can go straight to the useful card. */
function onFileScopeGranted(e) {
  return mainCard();
}

/**
 * The only honest test is the flag the host hands us. An earlier version asked whether
 * SpreadsheetApp could read the active sheet, which is always true under
 * spreadsheets.currentonly whether or not drive.file was ever granted for this file, so
 * the add-on showed its main card, never offered the grant, and failed at the export.
 */
function hasFileAccess(e) {
  return !!(e && e.sheets && e.sheets.addonHasFileScopePermission);
}

function grantCard() {
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText(
      'This pulls every comment in the spreadsheet into a new tab: which sheet and cell ' +
      'it sits on, who wrote it, when, whether it is resolved, and its replies.'))
    .addWidget(CardService.newTextParagraph().setText(
      '<b>It needs access to this file.</b> Access is granted one file at a time and ' +
      'nothing leaves your account.'))
    .addWidget(CardService.newTextButton()
      .setText('Grant access to this file')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('requestScope')));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Comment Export'))
    .addSection(section)
    .build();
}

function requestScope() {
  return CardService.newEditorFileScopeActionResponseBuilder()
    .requestFileScopeForActiveDocument()
    .build();
}

function mainCard(note) {
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText(
      'Writes every comment thread to a tab called <b>' + OUTPUT_TAB + '</b>: sheet, cell, ' +
      'author, date, open or resolved, the comment, and its replies.'))
    .addWidget(CardService.newTextButton()
      .setText('Export comments')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('onExport')));

  if (note) {
    section.addWidget(CardService.newTextParagraph().setText('<font color="#5B6068">' + note + '</font>'));
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Comment Export'))
    .addSection(section)
    .build();
}

/**
 * The button.
 *
 * Deliberately does NOT pre-check the grant. An earlier version gated on
 * e.sheets.addonHasFileScopePermission here; that flag is dependable on the homepage
 * trigger but not in an action callback, so on an already-granted file the check failed,
 * the handler returned a file-scope request, and Google renders a redundant scope request
 * as absolutely nothing. The button looked dead.
 *
 * Drive is the authority on whether we have access, so ask Drive and let it answer.
 */
function onExport(e) {
  var result;
  try {
    result = exportCommentsToTab();
  } catch (err) {
    if (err && err.message === NO_SCOPE) return requestScope();
    return notify('Could not export: ' + ((err && err.message) || 'unknown error'));
  }

  if (!result || result.threads === 0) {
    return notify('No comments found in this spreadsheet. Notes are a different Sheets ' +
                  'feature and do not appear here.');
  }

  return notify(result.threads + ' thread' + (result.threads === 1 ? '' : 's') +
                ' written to the "' + OUTPUT_TAB + '" tab.');
}

function notify(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .build();
}

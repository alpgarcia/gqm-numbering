/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
      .addItem('Use Case Numbering', '__numbering_ucs')
      .addItem('Question Numbering', '__numbering_questions')
      .addItem('Metric Numbering', '__numbering_metrics')
      .addSeparator()
      .addItem('Numbering All', '__numbering_all')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

function __numbering_ucs() {
  __numbering(DocumentApp.ParagraphHeading.HEADING2, 'UC');
}

function __numbering_questions() {
  __numbering(DocumentApp.ParagraphHeading.HEADING4, 'Q');
}

function __numbering_metrics() {
  __numbering(DocumentApp.ParagraphHeading.HEADING5, 'M');
}

function __numbering_all() {
  __numbering(DocumentApp.ParagraphHeading.HEADING2, 'UC');
  __numbering(DocumentApp.ParagraphHeading.HEADING4, 'Q');
  __numbering(DocumentApp.ParagraphHeading.HEADING5, 'M');
}

function __numbering(heading, prefix) {
  // Add a numbered 'prefix' to the specified 'heading' texts.
  // PLEASE DON'T USE TABS IN TITLES, they are used to separate and identify
  // the numbered prefixes.
  // Support using same number for the same title: if the same title is used
  // several times with the same heading, then this function will assign the
  // same numbered prefix to all its occurences.  
  
  // Base code borrowed from https://webapps.stackexchange.com/a/46905
  
  var pars = DocumentApp.getActiveDocument().getBody().getParagraphs();
  var counterh = 0;
  var titles = {}
  
  for(var i=0; i < pars.length; i++) {
    var par = pars[i];
    var hdg = par.getHeading();
    
    if (hdg == heading) {
      var content = par.getText();
      var chunks = content.split('\t');
      var title = chunks[0]
      if(chunks.length > 1) {
        title = chunks[1]        
      }
      
      // If there is another heading with the same title, re-use its numbering
      // in other case number it normally.
      var pre_title = ''
      if (title in titles) {
          pre_title = titles[title]
          
      } else {
          counterh++;
          pre_title = prefix + counterh
          titles[title] = pre_title
      }
      
      par.setText(pre_title +'.\t' + title); 
      
    }
  }
  
}

//function _old_numbering() {
//  
//  // Base code borrowed from https://webapps.stackexchange.com/a/46905
//  
//  var pars = DocumentApp.getActiveDocument().getBody().getParagraphs();
//  var counterh2 = 0;
//  var counterh4 = 0;
//  var counterh5 = 0;
//  for(var i=0; i < pars.length; i++) {
//    var par = pars[i];
//    var hdg = par.getHeading();
//    
//    if (hdg == DocumentApp.ParagraphHeading.HEADING2) {
//      counterh2++; 
//      var content = par.getText();
//      var chunks = content.split('\t');
//      if(chunks.length > 1) { 
//        par.setText('UC' + counterh2 + '.\t' + chunks[1]); 
//      } else {
//        par.setText('UC' + counterh2 +'.\t' + chunks[0]); 
//      }
//    }
//    
//    if (hdg == DocumentApp.ParagraphHeading.HEADING4) {
//      counterh4++; 
//      var content = par.getText();
//      var chunks = content.split('\t');
//      if(chunks.length > 1) { 
//        par.setText('Q' + counterh4+'.\t' + chunks[1]); 
//      } else {
//        par.setText('Q' + counterh4+'.\t' + chunks[0]); 
//      }
//    }
//    
//    if (hdg == DocumentApp.ParagraphHeading.HEADING5) {
//      counterh5++; 
//      var content = par.getText();
//      var chunks = content.split('\t');
//      if(chunks.length > 1) { 
//        par.setText('M' + counterh5+'.\t' + chunks[1]); 
//      } else {
//        par.setText('M' + counterh5+'.\t' + chunks[0]); 
//      }
//    }
//    
//  }
//}

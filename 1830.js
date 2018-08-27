/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */
// https://github.com/mikelietz/18xx-sheet-scripts/blob/master/1830.js

/* Version 0.9 */

function nextRound( s, d ) {
  var source = s; // current tab
  var ss = String( source );
  var destination = d; // new tab
  var thisRound = ss.substring( 0, 2 ); // SR or OR (or IS)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than SR99.3?)
  var thisOR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2, 3

  // if it is a SR, check Phase. destination could be wrong?
  
  var sheet = SpreadsheetApp.getActive();
  var templateSheet = sheet.getSheetByName( 'template' );
  sheet.insertSheet(d, 999, { template: templateSheet} );

  sheet.getRange( 'F14:U17' ).setNumberFormat( '@' ); // plaintext for X/X stuff
  sheet.getRange( 'V3:W9' ).setNumberFormat( '@' ); // more plaintext, privates.
  
  CopyRange( source, destination, 'A11:A11' ); // phase
  var Phase = sheet.getRange( 'A11' ).getValue();  

  CopyRange( source, destination, 'A3:B9' ); // players and priority
  if ( Phase < 5 ) {
    CopyRange( source, destination, 'F3:W9' ); // player stocks and privates
    CopyRange( source, destination, 'F9:U17' ); // company share price and other company stuff including privates
  } else {
    // Privates are closed in phase 5.
    CopyRange( source, destination, 'F3:U9' ); // player stocks
    CopyRange( source, destination, 'F9:U15' ); // company share price and other company stuff
  }
  CopyRange( source, destination, 'F10:U10' ); // company shares in market
  CopyRange( source, destination, 'Y20:AA22' ); // Trains in Market
  
  // set up the AE column
  sheet.getRange( 'AE9' ).setValue( source ); // previous round
  sheet.getRange( 'AE11' ).setValue( destination ); // this round
  CopyRange( source, destination, 'AE15' ); // carry over ORs, will get overwritten if SR

  var ORs = sheet.getRange( 'AE15' ).getValue();

  if ( thisRound == 'OR' ) {
     if ( ORs == 3 ) {
     // check for which .X
        if ( thisOR == 1 ) { // 1 of 3, next next = OR .3
        sheet.getRange( 'AE13' ).setValue( 'OR' + String( ( thisRoundNumber + .2 ).toFixed( 1 )  ) );
      } else if ( thisOR == 2 ) { // 1 of 3, next next = SR+1
        sheet.getRange( 'AE13' ).setValue( 'SR' + String( parseInt( thisRoundNumber ) + 1 ) );
      } else { // 3 of 3, next next = OR+1.1
        sheet.getRange( 'AE13' ).setValue( 'OR' + String( ( parseInt( thisRoundNumber ) + 1.1 ).toFixed( 1 ) ) ); // 1.1 + .1 = 1.19999999 otherwise
        // color the next tab for SRs
        sheet.getSheetByName( destination ).setTabColor( "888888" );
      }
    } else if ( ORs == 2 ) {
      if ( thisOR == 1 ) { // 1 of 2, next next = SR+1
        sheet.getRange( 'AE13' ).setValue( 'SR' + String( parseInt( thisRoundNumber ) + 1 ) );
      } else {
        sheet.getRange( 'AE13' ).setValue( 'OR' + String( ( parseInt( thisRoundNumber ) + 1.1 ).toFixed( 1 ) ) ); // 1.1 + .1 = 1.19999999 otherwise
        // color the next tab for SRs
        sheet.getSheetByName( destination ).setTabColor( "888888" );
      }
    } else if ( ORs == 1 ) {
      if ( Phase > 2 ) {
       // More than one OR next next time
       sheet.getRange( 'AE13' ).setValue( 'OR' + String( ( thisRoundNumber + 1.1 ).toFixed( 1 ) ) ); // 1.1 + .1 = 1.19999999 otherwise
      } else {
       // Only one OR next next time
       sheet.getRange( 'AE13' ).setValue( 'OR' + String( thisRoundNumber + 1 ) );
      }
      // color the next tab for SRs
      sheet.getSheetByName( destination ).setTabColor( "888888" );
    } else {
      // this shouldn't ever happen, can probably make ==1 the else.
      sheet.getRange( 'A16' ).setValue( 'ERROR!' );
      sheet.getRange( 'A31' ).setValue( 'Something has gone very wrong with ORs count.' );
    }
  } 
  else if (thisRound == 'SR') {
    // determine the number of ORs based on the Phase
    switch( Phase ) {
      case 2:
        // 1 OR, so next round's next round is the next SR
        sheet.getRange( 'AE15' ).setValue( 1 );
        sheet.getRange( 'AE13' ).setValue( 'SR' + String( thisRoundNumber + 1 ) );
        break;
      case 3:
      case 4:
        // 2 ORs
        sheet.getRange( 'AE15' ).setValue( 2 );
        sheet.getRange( 'AE13' ).setValue( 'OR' + String( thisRoundNumber ) + '.2' );
        break;
      // case 5-7:
      default:
        // 3 ORs, so next round of the next sheet is ORX.2
        sheet.getRange( 'AE15' ).setValue( 3 );
        sheet.getRange( 'AE13' ).setValue( 'OR' + String( thisRoundNumber ) + '.2' );
    }
  }
  else if (thisRound == 'IS' ) {
    // next round's next round is OR1
    sheet.getRange( 'AE13' ).setValue( 'OR1' ); 
    // advance the Phase.
    sheet.getRange( 'A11' ).setValue( 2 );
    // color the next tab for SRs
    sheet.getSheetByName( destination ).setTabColor( "888888" );
  } else {
    // this shouldn't ever happen, can probably make 'IS' the else.
    sheet.getRange( 'A16' ).setValue( 'ERROR!' );
    sheet.getRange( 'A31' ).setValue( 'Something has gone very wrong with round type.' );
  }    
}

function CopyRange( source, dest, copyrange ) {
  // https://productforums.google.com/forum/#!msg/docs/SwIuouNeblw/rTTJq0WyNwAJ
  var sss = SpreadsheetApp.getActive();
  var ss = sss.getSheetByName( source );
  var data = ss.getRange( copyrange ).getValues();
  var ts = sss.getSheetByName( dest );
  ts.getRange( copyrange ).setValues( data );
}

function Cleanup() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for (var i = 0; i < numSheets; i++){

    var thisSheet = sheets[i].getName();
    if (thisSheet != 'ISR' && thisSheet != 'template' && thisSheet != 'Privates Auction' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }
  return;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu( '1830 Menu' )
      .addItem( 'Next Round' , 'menuItem1' )
//      .addItem( 'Destructive Reset' , 'Cleanup' ) /* uncomment this line to have an easy cleanup in the menu */
  .addToUi();
}



function menuItem1() {
  nextRound( 
    SpreadsheetApp.getActiveSheet().getRange('AE11').getValues(), // source
    SpreadsheetApp.getActiveSheet().getRange('AE13').getValues() // destination
  );
}

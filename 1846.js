/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */
function nextRound( s ) {
  var source = s; // current tab
  var sheet = SpreadsheetApp.getActive();
  var Phase = sheet.getRange( 'A9' ).getValue();
  
  var destination = DetermineNextRound( s, Phase ); // new tab

  sheet.getRange( 'A22' ).setValue( 'old' ); // mark existing sheet 'old' under the button
  
  var templateSheet = sheet.getSheetByName( 'template' );
  sheet.insertSheet( destination, 999, { template: templateSheet } );

  sheet.getRange( 'F12:U16' ).setNumberFormat( '@' ); // plaintext for X/X stuff
  CopyRange( source, destination, 'A9:A9' ); // phase
  CopyRange( source, destination, 'B3:B7' ); // players and priority
  CopyRange( source, destination, 'F3:W7' ); // player stocks and privates
  CopyRange( source, destination, 'H8:U8' ); // company shares in market
  CopyRange( source, destination, 'F12:U16' ); // company stuff (including Independent trains)
  CopyRange( source, destination, 'Y19:AA20' ); // Trains in Market
  
  // set up the AE column
  sheet.getRange( 'AE9' ).setValue( source ); // previous round
  sheet.getRange( 'AE11' ).setValue( destination ); // this round

  // color the new tab for SRs
  if( destination.substring( 0, 2 ) == 'SR' ) {
      sheet.getSheetByName( destination ).setTabColor( "888888" );
  }  
}

function DetermineNextRound( source, phase ) {
  var ss = String( source );
  var thisRound = ss.substring( 0, 2 ); // SR or OR (or IS)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than SR99.3?)
  var thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2

  // return the next round
  switch ( thisRound + thisR ) {
    case 'SR0':
      // next round is OR.1
      return 'OR' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'OR1':
      // next round is OR.2
      return 'OR' + String( parseInt( thisRoundNumber ) ) + '.2';
    case 'OR2':
      // next round is SR+1
      return 'SR' + String( parseInt( thisRoundNumber ) + 1 );
    default:
      // ISR -> SR1
      return 'SR1';
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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu( '1846 Menu' )
      .addItem( 'Next Round' , 'menuItem1' )
      .addItem( 'Destructive Reset' , 'Cleanup' )
  .addToUi();
}

function menuItem1() {
  nextRound( 
    SpreadsheetApp.getActiveSheet().getRange( 'AE11' ).getValues() // source
  );
}

function Cleanup() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for (var i = 0; i < numSheets; i++){

    var thisSheet = sheets[i].getName();
    if (thisSheet != 'ISR' && thisSheet != 'template' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }
  return;
}

/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */
// https://github.com/mikelietz/18xx-sheet-scripts/blob/master/1830.js

/* Version 1.0 */

function nextRound( s ) {
  var source = s; // current tab
  var sheet = SpreadsheetApp.getActive();

  if( source != "Privates Auction" ) {
    var Phase = sheet.getRange( 'A11' ).getValue();
 
    var destination = DetermineNextRound( s, Phase ); // new tab
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    
    sheet.getRange( 'F15:U18' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    CopyRange( source, destination, 'A11:A11' ); // phase
    CopyRange( source, destination, 'B3:B9' ); // player priority
    if ( Phase < 5 ) {
      CopyRange( source, destination, 'F3:W9' ); // player stocks and privates
      CopyRange( source, destination, 'F14:U18' ); // company end share price and other company stuff including privates
    } else {
      // Privates are closed in phase 5
      CopyRange( source, destination, 'F3:U9' ); // player stocks
      CopyRange( source, destination, 'F14:U16' ); // company end share price and other company stuff
    } 
    
    CopyRange( source, destination, 'F10:T10' ); // company shares in market
    CopyRange( source, destination, 'F12:T12' ); // company IPO prices
    CopyRange( source, destination, 'Y20:AA21' ); // Trains in Market
  } else {  
    // ISR to SR1
    var destination = 'SR1';
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    sheet.getRange( 'F15:U18' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    CopyRange( source, destination, 'A11:A11' ); // phase
    CopyRange( source, destination, 'B3:B9' ); // players and priority
    CopyRange( source, destination, 'W3:W9' ); // player stocks and privates

    sheet.getSheetByName( 'SR1' ).getRange( 'V3:V9' ).setValues( sheet.getSheetByName( 'SR1' ).getRange( 'W3:W9' ).getValues() ); // Copy owned to income privates
     
    sheet.getRange( 'F13:T13' ).setValue( '' ); // blank out the previous market price for companies
    sheet.getRange( 'F19:T19' ).setValue( '' ); // blank out the begin treasury for companies
  }
  // set up the AE column
  sheet.getRange( 'AE11' ).setValue( source ); // previous round
  sheet.getRange( 'AE13' ).setValue( destination ); // this round
  
  // color the new tab for SRs
  if( destination.substring( 0, 2 ) == 'SR' ) {
    sheet.getSheetByName( destination ).setTabColor( "888888" );
  }  
  
  var numPlayers = sheet.getRange( 'A1' ).getValue();
  if ( numPlayers < 7 ) {
    // hide non-player rows
    sheet.getSheetByName( destination ).hideRows( 3 + numPlayers, 7 - numPlayers );
  }  
  

}

function DetermineNextRound( source, phase ) {
  var ss = String( source );
  var thisRound = ss.substring( 0, 2 ); // SR or OR (or IS)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than SR99.3?)
  var thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2
  var ORsPerPhase = [ 0,1,1,2,2,3,3,3 ];
  var numberOfORs = ORsPerPhase[ phase ]; // phases 1-6
                     
  // return the next round
  switch ( thisRound + thisR + numberOfORs ) { // read as round whatever X out of Y
    case 'SR01':
    case 'SR02':
    case 'SR03':
      // next round is always OR.1
      return 'OR' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'OR11':
    case 'OR22':
    case 'OR33':
      // next round is SR+1
      return 'SR' + String( parseInt( thisRoundNumber ) + 1 );
    case 'OR12':
    case 'OR13':
      // next round is OR.2
      return 'OR' + String( parseInt( thisRoundNumber ) ) + '.2';
    case 'OR23':
      // next round is OR.3
      return 'OR' + String( parseInt( thisRoundNumber ) ) + '.3';
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
    SpreadsheetApp.getActiveSheet().getRange( 'AE13' ).getValues()
  );
}

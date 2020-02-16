/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */
// https://github.com/mikelietz/18xx-sheet-scripts/blob/master/1889.js
// @OnlyCurrentDoc

/* Version 1.5 */
function nextRound( s ) {
  var source = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var sheet = SpreadsheetApp.getActive();
  var orp = parseInt( PropertiesService.getDocumentProperties().getProperty( 'orPhase' ) );
  if( source.substring( 0, 2 ) == 'SR' ) { //sheet properties can be used like global variables.  Initial property is ( 'orPhase', '2' ).  This updates only during SR
    PropertiesService.getDocumentProperties().setProperty( 'orPhase', sheet.getRange( 'A11' ).getValue() );
    orp = parseInt( PropertiesService.getDocumentProperties().getProperty( 'orPhase' ) );
  } 
  
  if( source != "Privates Auction" ) {
    var Phase = sheet.getRange( 'A11' ).getValue();
 
    var destination = DetermineNextRound( s, orp ); // new tab
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    
    sheet.getRange( 'F15:R18' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    sheet.getRange( 'T3:T8' ).setNumberFormat( '@' ); // plaintext for X X stuff
    
    CopyRange( source, destination, 'B3:B8' ); // player priority
    CopyRange( source, destination, 'F3:T8' ); // player stocks and privates
    if ( Phase < 5 ) {
      CopyRange( source, destination, 'F14:S17' ); // company end share price and other company stuff including privates
    } else {
      // Privates are closed in phase 5, except for Uno-Takamatsu/G if owned by a player, so ...
      CopyRange( source, destination, 'F14:S16' ); // company end share price and other company stuff
    } 
    
    CopyRange( source, destination, 'F10:R10' ); // company shares in market
    CopyRange( source, destination, 'F12:R12' ); // company IPO prices
    CopyRange( source, destination, 'W20:Y21' ); // Trains in Market
  } else {  
    // ISR to SR1
    var destination = 'SR1';
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    sheet.getRange( 'F15:R18' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    sheet.getRange( 'T3:T8' ).setNumberFormat( '@' ); // plaintext for X X stuff

    CopyRange( source, destination, 'A11:A11' ); // phase
    CopyRange( source, destination, 'B3:B8' ); // player priority
    CopyRange( source, destination, 'T3:T8' ); // player stocks and privates

    sheet.getRange( 'F13:R13' ).setValue( '' ); // blank out the previous market price for companies
    sheet.getRange( 'F19:R19' ).setValue( '' ); // blank out the begin treasury for companies

    PropertiesService.getDocumentProperties().setProperty( 'orPhase', 2 );
  }
  // set up the AC column
  sheet.getRange( 'AC11' ).setValue( source ); // previous round
  sheet.getRange( 'AC13' ).setValue( destination ); // this round
  
  // color the new tab for SRs
  if( destination.substring( 0, 2 ) == 'SR' ) {
    sheet.getSheetByName( destination ).setTabColor( "888888" );
  }  
  
  var numPlayers = sheet.getRange( 'A1' ).getValue();
  if ( numPlayers < 6 ) {
    // hide non-player rows
    sheet.getSheetByName( destination ).hideRows( 3 + numPlayers, 6 - numPlayers );
  }
  
}

function DetermineNextRound( source, orp ) {
  var ss = String( source );
  var thisRound = ss.substring( 0, 2 ); // SR or OR (or IS)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than SR99.3?)
  var thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2

  var ORsPerPhase = [ 0,0,1,2,2,2,3 ];
  if ( orp == "D" ) { orp = 6; }
  var numberOfORs = ORsPerPhase[ orp ]; // phases 1-6
  // return the next round
  switch ( thisRound + thisR + numberOfORs ) { // read as round whatever X out of Y
    case 'SR01':
      return 'OR' + String( parseInt( thisRoundNumber ) );
    case 'SR02':
    case 'SR03':
      // next round is always OR.1
      return 'OR' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'OR01':
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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu( '1889 Menu' )
      .addItem( 'Next Round' , 'menuItem1' )
//      .addItem( 'Destructive Reset' , 'Cleanup' )
      .addToUi();
}

function menuItem1() {
  nextRound( 
    SpreadsheetApp.getActiveSheet().getRange( 'AC13' ).getValues()
  );
}

function Cleanup() {
  // delete any leftover tabs
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for (var i = 0; i < numSheets; i++){

    var thisSheet = sheets[i].getName();
    if (thisSheet != 'ISR' && thisSheet != 'template' && thisSheet != 'Privates Auction' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }

  // clear out the Privates Auction
  var thisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'Privates Auction' );
  thisSheet.getRange( 'A3:A8' ).setValue( '' ); // blank out the players
  thisSheet.getRange( 'G3:M8' ).setValue( '' ); // blank out the bids
  thisSheet.getRange( 'T3:T8' ).setValue( '' ); // blank out the privates  
  
  // hide the template
  var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'template' );
  templateSheet.showSheet()
  templateSheet.hideSheet()
  
  return;
}

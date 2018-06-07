/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */

/* Version 1.0 */

function nextRound( s ) {
  var source = s; // current tab
  var sheet = SpreadsheetApp.getActive();
  var Phase = sheet.getRange( 'A10' ).getValue();

  var destination = DetermineNextRound( s, Phase) ; // new tab
  
  // sheet.getRange( 'A21' ).setValue( 'old' ); // mark existing sheet 'old' under the button
  
  var templateSheet = sheet.getSheetByName( 'template' );
  sheet.insertSheet( destination, 999, { template: templateSheet} );

  sheet.getRange( 'F13:AK15' ).setNumberFormat( '@' ); // plaintext for X/X stuff
  sheet.getRange( 'AN3:AO8' ).setNumberFormat( '@' ); // plaintext for privates

  // set up the AX column
  sheet.getRange( 'AX10' ).setValue( source ); // previous round
  sheet.getRange( 'AX12' ).setValue( destination ); // this round


  CopyRange( source, destination, 'A10:A10' ); // phase
  CopyRange( source, destination, 'B3:B8' ); // players and priority
  CopyRange( source, destination, 'F3:AO8' ); // player stocks and privates
  CopyRange( source, destination, 'V9:AJ9' ); // company pool
  CopyRange( source, destination, 'F13:AL17' ); // company trains, tokens, etc
  CopyRange( source, destination, 'AQ23:AS25' ); // Trains in Market
  sheet.getRange( 'F12:AL12' ).setValues( sheet.getRange( 'F11:AL11' ).getValues() ); // prepopulate end price with begin price. Moves down a row, so doesn't use CopyRange function

  // color the new tab  
  switch ( destination.substring( 0, 2 ) ) {
    case 'SR':
      // color the next tab for SRs
      sheet.getSheetByName( destination ).setTabColor( "FFFFFF" );
      break;
    case 'MR':
      // color the next tab for MRs
      sheet.getSheetByName( destination ).setTabColor( "afb0b3" );
      break;
    case 'OR':
      // color the next tab for ORs
      sheet.getSheetByName( destination ).setTabColor( "b8c0e5" );
      break;
    default:
      // do nothing
  }

  var numPlayers = sheet.getRange( 'A1' ).getValue();
  if ( numPlayers < 6 ) {
    // hide non-player rows
    sheet.getSheetByName( destination ).hideRows( 3 + numPlayers, 6 - numPlayers );
  }

  // hide unavailable company columns
  switch ( Phase ) {
    case 2:
      sheet.getSheetByName( destination ).hideColumns( 16, 6 ); // hide the greens
      sheet.getSheetByName( destination ).hideColumns( 22, 18 ); // hide the publics
      break;
    case 3:
    case 4:
    case 5:
    case 6:
    case 7:
      if (sheet.getSheetByName( destination ).getRange( 'V3:AK8' ).isBlank() && destination.substring( 0, 2 ) == 'OR' ) {
        sheet.getSheetByName( destination ).hideColumns( 22, 18 ); // hide the publics
      }
      break;
    case 8:
      sheet.getSheetByName( destination ).hideColumns( 6, 16 ); // hide the minors
  }

}

function DetermineNextRound( source, phase ) {
  var ss = String( source );
  var thisRound = ss.substring( 0, 2 ); // SR or OR or MR (or IS)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than SR99.3?)
  var thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2
  var RoundSwitch;
  
  if ( phase > 2 && phase < 8 ) {
    RoundSwitch = thisRound + thisR + "M";
  } else if ( thisRound == 'IS' ) {
    RoundSwitch = 'IS';
  } else { 
    RoundSwitch = thisRound + thisR;    
  }

  // return the next round
  switch ( RoundSwitch ) {
    case 'IS':
      return 'SR1';
    case 'OR1M':
      // next round is MR.1
      return 'MR' + String( thisRoundNumber );
    case 'SR0':
    case 'SR0M':
    case 'MR1':
      // next round is OR.1
      return 'OR' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'OR2M':
      // next round is MR.2
      return 'MR' + String( parseInt( thisRoundNumber ) ) + '.2';
    case 'OR1':
    case 'MR1M':
      // next round is OR.2
      return 'OR' + String( parseInt( thisRoundNumber ) ) + '.2';
    case 'OR2':
    case 'MR2': // regardless of phase, SR after MR.2
    case 'MR2M':
      // next round is SR+1
      return 'SR' + String( parseInt( thisRoundNumber ) + 1 );
    default:
      SpreadsheetApp.getActive().getRange( 'A30' ).setValue( 'Unexpected round switch! ' + RoundSwitch );
  }
  return 'SR1'; // for ISR
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
  ui.createMenu( '1867 Menu' )
      .addItem( 'Next Round' , 'menuItem1' )
//      .addItem( 'Destructive Reset' , 'Cleanup' )
  .addToUi()
}

function menuItem1() {
  nextRound( 
    SpreadsheetApp.getActiveSheet().getRange( 'AX12' ).getValues() // source
  );
}

function Cleanup() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for (var i = 0; i < numSheets; i++){

    var thisSheet = sheets[i].getName();
    if (thisSheet != 'ISR' && thisSheet != 'template' && thisSheet != 'Auctions' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }
  return;
}

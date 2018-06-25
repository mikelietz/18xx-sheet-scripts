// https://github.com/mikelietz/18xx-sheet-scripts/blob/master/1800.js

/* Version 1.0 */

function nextRound( s ) {
  var source = s; // current tab
  var sheet = SpreadsheetApp.getActive();
  var Phase = sheet.getRange( 'A9' ).getValue();
  
  var destination = DetermineNextRound( s, Phase ); // new tab

  sheet.getRange( 'A22' ).setValue( 'old' ); // mark existing sheet 'old' under the button
  
  var templateSheet = sheet.getSheetByName( 'template' );
  sheet.insertSheet( destination, 999, { template: templateSheet } );

  sheet.getRange( 'F12:U16' ).setNumberFormat( '@' ); // plaintext for X/X stuff
  sheet.getRange( 'V3:W7' ).setNumberFormat( '@' ); // plaintext for Special cards (for the 0)
 
  CopyRange( source, destination, 'A9:A9' ); // phase
  CopyRange( source, destination, 'B3:B7' ); // players and start player
  CopyRange( source, destination, 'F3:V7' ); // player merchants and special cards
  CopyRange( source, destination, 'F8:U8' ); // merchants pool
  CopyRange( source, destination, 'F12:U14' ); // nation stuff

  sheet.getRange( 'D23' ).setValue( 'King Cash (-'+ [0,5,10,15,20][Phase] +')' ); // Cash for the king changes by Phase, but happens at the beginning of the round.
  
  var incomes = [ [], [10,20,30,40,50], [5,10,15,0,0], [0,0,0,0,0], [0,0,0,0,0] ];
  for ( i = 0; i < 5; i++ ) { // fill the incomes in Z25:Z29
    sheet.getRange( 'Z' + (i + 25) ).setValue( incomes[ Phase ][ i ] ); // Special Income changes by Phase, but happens at the beginning of the round.
  }

  // set up the AF column
  sheet.getRange( 'AF9' ).setValue( source ); // previous round
  sheet.getRange( 'AF11' ).setValue( destination ); // this round

  // color the new tab for MRs
  if( destination.substring( 0, 2 ) == 'MR' ) {
      sheet.getSheetByName( destination ).setTabColor( "FFFFFF" );
  }  

  var numPlayers = sheet.getRange( 'A1' ).getValue();
  if ( numPlayers < 5 ) {
    // hide non-player rows
    sheet.getSheetByName( destination ).hideRows( 3 + numPlayers, 5 - numPlayers );
  }
}

function PRtoMR1() {
  var source = 'PR'; 
  var sheet = SpreadsheetApp.getActive();
  var Phase = 1;
  
  var destination = DetermineNextRound( 'PR', Phase ); // new tab

  sheet.getRange( 'A22' ).setValue( 'old' ); // mark existing sheet 'old' under the button
  
  var templateSheet = sheet.getSheetByName( 'template' );
  sheet.insertSheet( destination, 999, { template: templateSheet } );

  // fill in starting values for all the nations, since the previous round lookups will fail!
  
  sheet.getRange( 'F12:U16' ).setNumberFormat( '@' ); // plaintext for X/X stuff
  sheet.getRange( 'V3:W7' ).setNumberFormat( '@' ); // plaintext for Special cards (for the 0)
  
  sheet.getRange( 'F13:U13' ).setValues([[ '1/1', '', '1/1', '', '1/1', '', '1/1', '', '1/1', '', '1/1', '', '1/1', '', '1/1', '' ]]); // set New Home 1
  sheet.getRange( 'F10:U10' ).setValues([[ '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' ]]); // blank out Begin Prestige
  sheet.getRange( 'F15:U15' ).setValues([[ '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' ]]); // blank out Begin Nation Funds

  
  CopyRange( source, destination, 'A9:A9' ); // phase
  CopyRange( source, destination, 'B3:B7' ); // players and priority
  sheet.getSheetByName( 'MR1' ).getRange( 'D3:D7' ).setValues( sheet.getSheetByName( 'PR' ).getRange( 'AB3:AB7' ).getValues() ); // prepopulate worth with cash. Specials are worthless.
  sheet.getSheetByName( 'MR1' ).getRange( 'V3:V7' ).setValues( sheet.getSheetByName( 'PR' ).getRange( 'W3:W7' ).getValues() ); // Move Specials left one column.

  sheet.getRange( 'D23' ).setValue( 'King Cash (-'+ [0,5,10,15,20][Phase] +')' ); // Cash for the king changes by Phase, but happens at the beginning of the round.
  
  sheet.getRange( 'Z25' ).setValue( 10 ); // Special Income changes by Phase, but happens at the beginning of the round.
  sheet.getRange( 'Z26' ).setValue( 20 ); // Special Income changes by Phase, but happens at the beginning of the round.
  sheet.getRange( 'Z27' ).setValue( 30 ); // Special Income changes by Phase, but happens at the beginning of the round.
  sheet.getRange( 'Z28' ).setValue( 40 ); // Special Income changes by Phase, but happens at the beginning of the round.
  sheet.getRange( 'Z29' ).setValue( 50 ); // Special Income changes by Phase, but happens at the beginning of the round.


  // set up the AF column
  sheet.getRange( 'AF9' ).setValue( 'PR' ); // previous round
  sheet.getRange( 'AF11' ).setValue( destination ); // this round

  // color the new tab for MRs
  sheet.getSheetByName( destination ).setTabColor( "FFFFFF" );

  var numPlayers = sheet.getRange( 'A1' ).getValue();
  if ( numPlayers < 5 ) {
    // hide non-player rows
    sheet.getSheetByName( destination ).hideRows( 3 + numPlayers, 5 - numPlayers );
  }
}

function DetermineNextRound( source, phase ) {
  var ss = String( source );
  var thisRound = ss.substring( 0, 2 ); // MR or ER (or PR)
  var thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than ER5.3)
  var thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2

  // return the next round
  switch ( thisRound + thisR ) {
    case 'MR0':
      // next round is ER.1
      return 'ER' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'ER1':
      // next round is ER.2
      return 'ER' + String( parseInt( thisRoundNumber ) ) + '.2';
    case 'ER2':
      // next round is MR+1, unless it's ER5.2 -> ER5.3
      if ( ss == 'ER5.2' ) {
        return 'ER5.3';
      }
      return 'MR' + String( parseInt( thisRoundNumber ) + 1 );
    default:
      // PR -> MR1
      return 'MR1';
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
  ui.createMenu( 'Poseidon Menu' )
      .addItem( 'Next Round' , 'menuItem1' )
      .addItem( 'Destructive Reset' , 'Cleanup' ) /* uncomment this line to have an easy cleanup in the menu */
  .addToUi();
}

function nextRound() {
  nextRound( 
    SpreadsheetApp.getActiveSheet().getRange( 'AF11' ).getValues() // source
  );
}

function Cleanup() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for ( var i = 0; i < numSheets; i++ ) {

    var thisSheet = sheets[i].getName();
    if (thisSheet != 'PR' && thisSheet != 'template' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }
  return;
}

/* Note, this is not actually a javascript source file. This is to be used for Google Sheets in the Tools->Script Editor. */
// https://github.com/mikelietz/18xx-sheet-scripts/blob/master/18chesapeake.js
// @OnlyCurrentDoc

/* Version 0.2 */
function nextRound() {
  var source = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); //this is tab name (also in AE13)
  var sheet = SpreadsheetApp.getActive();
  var destination = DetermineNextRound( source ); // new tab

  if( source != "Privates Auction" ) {

    // If it's the end of an OR, we need to export a train. The last exported train will also represent the current Phase until Phase 5.
    var exportedTrains = GetExportedTrains( source, destination );
    var lastExportedTrain = exportedTrains.charAt( exportedTrains.length - 1);
    var Phase = Math.max( lastExportedTrain, sheet.getRange( 'A11' ).getValue() );
 
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    
    sheet.getRange( 'F15:U17' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    sheet.getRange( 'V3:V8' ).setNumberFormat( '@' ); // plaintext for X X stuff

    sheet.getRange( 'Y20' ).setValue( exportedTrains ); // actually set the exported trains now
    
    CopyRange( source, destination, 'B3:B8' ); // player priority

    if ( Phase < 5 ) {
      CopyRange( source, destination, 'F3:V8' ); // player stocks and privates
      CopyRange( source, destination, 'F14:T17' ); // company end share price and other company stuff including privates

    } else {
      // Privates are closed in phase 5 and thereafter
      CopyRange( source, destination, 'F3:U8' ); // player stocks only
      CopyRange( source, destination, 'F14:T16' ); // company end share price and other company stuff
    } 
    
    CopyRange( source, destination, 'F10:T10' ); // company shares in market
    CopyRange( source, destination, 'F12:T12' ); // company IPO prices
    
  } else {  
    // ISR to SR1
    var templateSheet = sheet.getSheetByName( 'template' );
    sheet.insertSheet( destination, 999, { template: templateSheet } );
    sheet.getRange( 'F15:U17' ).setNumberFormat( '@' ); // plaintext for X/X stuff
    sheet.getRange( 'V3:V8' ).setNumberFormat( '@' ); // plaintext for X X stuff
    
    sheet.getRange( 'A11' ).setValue( 2 );
    CopyRange( source, destination, 'V3:V8' ); // player privates

    sheet.getRange( 'F13:U13' ).setValue( '' ); // blank out the previous market price for companies
    sheet.getRange( 'F19:U19' ).setValue( '' ); // blank out the begin treasury for companies

    PropertiesService.getDocumentProperties().setProperty( 'SRPhase', 2 );
    PropertiesService.getDocumentProperties().setProperty( 'ORCount', 1 );

  }
  // set up the rightmost AE column
  sheet.getRange( 'AE11' ).setValue( source ); // previous round
  sheet.getRange( 'AE13' ).setValue( destination ); // this round
  
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

function DetermineNextRound( source ) {
  var ss = String( source );
//  Browser.msgBox( 'source string = ' + ss );
  var thisRound = ss.substring( 0, 2 ); // SR or OR
  var thisR = 0;
  var thisRoundNumber, phase, sss;
// Browser.msgBox( PropertiesService.getDocumentProperties().getProperty( 'SRPhase' ) );
  
//  Browser.msgBox( 'thisRound = ' + thisRound );

  // switch here based on whether it's an OR or an SR
  if ( thisRound == 'SR' ) {
    thisRoundNumber = parseFloat( ss.substring( 2, 4 ) ); // 1, 2, 10... (will never be bigger than SR99?)

    sss = SpreadsheetApp.getActive().getSheetByName( source );
    phase = sss.getRange( 'A11' ).getValue();
    
  } else {
    if ( thisRound != 'OR' ) { // ie Privates Auction
      if ( thisRound != 'Pr' ) {
        Browser.msgBox( 'Something is wrong in DetermineNextRound - it should be an OR, right?');
      } else {
        return 'SR1';
      }
    }    
    thisRoundNumber = parseFloat( ss.substring( 2, 6 ) ); // 1, 1.1, 2.2, etc (will never be bigger than OR99.3?)
    thisR = ( thisRoundNumber * 10 ) % 10; // 0, 1, 2
    sss = SpreadsheetApp.getActive().getSheetByName( "SR" + parseInt( thisRoundNumber ) );
//    Browser.msgBox( 'looking at SR' + parseInt(thisRoundNumber ) );
    phase = sss.getRange( 'A11' ).getValue();
//    Browser.msgBox( "SR" + parseInt(thisRoundNumber ) + 'began in phase ' + phase ); 
  }
  
  if ( phase == 'D' ) { phase = 6; }
  
  var ORsPerPhase = [ 1,2,2,3,3,3 ];
  var numberOfORs = ORsPerPhase[ phase - 2 ];
  // return the next round
  switch ( thisRound + thisR + numberOfORs ) { // read as round whatever X out of Y
    case 'SR01':
      return 'OR' + String( parseInt( thisRoundNumber ) );
    case 'SR02':
    case 'SR03':
      // next round is always OR.1
      return 'OR' + String( ( parseInt( thisRoundNumber ) + 0.1 ).toFixed( 1 ) ); // 0.1 + .1 = 1.19999999 otherwise (maybe)
    case 'OR01': // <- OR1 or any other OR in Phase 2
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
      Browser.msgBox( 'Something strange has happened in DetermineNextRound' );
      return 'SR1';
  }  
}

function GetExportedTrains( source, destination ) {
  var sss = SpreadsheetApp.getActive();
  var ss = sss.getSheetByName( source );
  var Phase = ss.getRange( 'A11' ).getValue();
  var Current = ss.getRange( 'Y20' ).getValue();
  var Last = parseInt( Current.substring( Current.length-1, Current.length ) );
  if ( Current.length == 0 ) { 
    Last = 2; // handle the first export well enough
  }
  
  // Browser.msgBox( 'most recently exported was ' + Last );
  // Browser.msgBox( String( destination ) );
  if ( Phase < 5 && String( destination ).substring( 0, 2 ) == 'SR' ) {
    switch ( Last ) {
      case 4:
      case 3:
        if ( parseFloat( String( ss.getRange( 'AA13' ).getValue() ).substring( 0,1 ) ) > 0 ) {
          // Don't even check the 2s, since they might be rusted. Are there any 3s left? If so, export one.
          return Current + " 3";
        }

        if ( parseFloat( String( ss.getRange( 'AA14' ).getValue() ).substring( 0,1 ) ) > 0 ) {
          // There weren't any 3s left. Are there any 4s left? If so, export one.
          return Current + " 4";
        } else {
          //Browser.msgBox( "ELSE Something has gone wrong with the Last Exported Train! Case 3 Last = " + Last );
 
         // There aren't any 4s left either.
          return Current;
        }
        Browser.msgBox( "Something has gone wrong with the Last Exported Train! Case 3 Last = " + Last );
        break;
      case 2:
         if ( parseFloat( String( ss.getRange( 'AA12' ).getValue() ).substring( 0,1 ) ) > 0 ) {
          // If there are any 2s left, export one.
          return Current + " 2";
        }
        
        if ( parseFloat( String( ss.getRange( 'AA13' ).getValue() ).substring( 0,1 ) ) > 0 ) {
          // There weren't any 2s left. Are there any 3s left? If so, export one.
          return Current + " 3";
        }

        if ( parseFloat( String( ss.getRange( 'AA14' ).getValue() ).substring( 0,1 ) ) > 0 ) {
          // There weren't any 3s left. Are there any 4s left? If so, export one. I find this highly unlikely (and thus don't have the check above for no 4s left either).
          return Current + " 4";
        }
        
        Browser.msgBox( "Something has gone wrong with the Last Exported Train! Case 2 Last = " + Last );
        break;
      default:
        Browser.msgBox( "Something has gone wrong with the Last Exported Train! Default Case Last = " + Last );
    }
       
    Browser.msgBox( 'Something has gone wrong in GetExportedTrains' );
    return 'UH OH!'; 
  } else {
    // There is no need to do anything new, just copy the field.
    return Current;
  }
}

function RandomCVCompany() {
  var sss = SpreadsheetApp.getActive().getSheetByName( 'Privates Auction' );
  var Companies = Array( "B&O", "C&A", "C&O", "LV", "N&W", "PRR", "PLE", "SRR" );
  sss.getRange( '$AE$16' ).setValue( Companies[ Math.floor( Math.random() * 8 ) ] );
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
  ui.createMenu( '18xx Menu' )
    .addItem( 'Next Round' , 'nextRound' )
    .addItem( 'Randomize CV', 'RandomCVCompany' )
//    .addItem( 'Destructive Reset', 'Cleanup' )
  .addToUi();
}

function Cleanup() {
  // delete any leftover tabs
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var numSheets = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  
  for ( var i = 0; i < numSheets; i++ ){

    var thisSheet = sheets[i].getName();
    if ( thisSheet != 'ISR' && thisSheet != 'template' && thisSheet != 'Privates Auction' ) {
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet( sheets[i] ); 
    }
  }

  // clear out the Privates Auction
  var thisSheet = SpreadsheetApp.getActive().getSheetByName( 'Privates Auction' );
  thisSheet.getRange( 'A3:A8' ).setValue( '' ); // blank out the players
  thisSheet.getRange( 'G3:L8' ).setValue( '' ); // blank out the bids
  thisSheet.getRange( 'V3:V8' ).setValue( '' ); // blank out the privates  
  thisSheet.getRange( 'AA3:AA8' ).setValue( '' ); // blank out the revenue
  thisSheet.getRange( 'AE16' ).setValue( '<Push the button!>' );

  // hide the template
  var templateSheet = SpreadsheetApp.getActive().getSheetByName( 'template' );
  templateSheet.showSheet()
  templateSheet.hideSheet()
  
  return;
}

var noon    = 0;
var normal  = 1;
var night   = 2;
var weekend = 3;

function locateEmptyDay() 
{
  var sheetsTab = SpreadsheetApp.getActiveSpreadsheet().getSheets()

  for(var sheetIdx = 0; sheetIdx < sheetsTab.length; sheetIdx++)
  {
    var sheetName = sheetsTab[sheetIdx].getName();          // Just for debug
    var range     = sheetsTab[sheetIdx].getRange('b2:f40'); // Range includes day and referent columns, from row 1 to 40 (40, just to be sure because of blank lines between weeks)
    dayCol        = 1;
    normalRefCol  = 2;
    nightRefCol   = 3;
    noonRefCol    = 5;
    lastRow       = range.getLastRow();
    
    // Get current date for later
    var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    var todayDate      = new Date();
    var tomorrowDate   = new Date(todayDate.getTime() + MILLIS_PER_DAY);
    
    // Loop on the days of the month
    for (var row = 1; row < lastRow; row++) 
    {  
      // Get the current day cell value to see if there a valid date in it
      var noonRefCell   = range.getCell(row, noonRefCol);
      var normalRefCell = range.getCell(row, normalRefCol);
      var nightRefCell  = range.getCell(row, nightRefCol);

      // Get the current day cell value to see if there a valid date in it
      var dayCell         = range.getCell(row, dayCol);  
      var completeDateStr = dayCell.getValue().toString();
      var currentCellDate = new Date(completeDateStr);
      
      // See if this is a valid date
      if(currentCellDate != "Invalid Date")
      {      
        // Handle mail to be sent for tomorrow
        if((currentCellDate.getDate() == (tomorrowDate.getDate())) && (currentCellDate.getMonth() == tomorrowDate.getMonth()))
        {
          /******************************************************************************/
          /* FIRST MAIL : Let's see if we have any referent here : for noon             */
          /******************************************************************************/
          if(doFirstCallTest(noon) && (noonRefCell.getValue().length <= 2))
          {
            // Send mail
            sendFirstCallMail(completeDateStr, noon);
          }
          
          /******************************************************************************/
          /* CANCELLEATION : Let's see if we have any referent here : for noon          */
          /******************************************************************************/
          if(doCancellationCallTest(noon) && (noonRefCell.getValue().length <= 2))
          {
            // Send mail
            sendCancellationMail(tomorrowDate.toString(), noon);
          }
          
          /******************************************************************************/
          /* FIRST MAIL : Let's see if we have any referent for the "normal" time slot  */
          /******************************************************************************/
          if(doFirstCallTest(normal) && (normalRefCell.getValue().length <= 2))
          {
            // Send mail : if test isn't done at 20h, this is a weekend timeslot
            if(todayDate.getHours() != 20)
            {
              var timeSlot = weekend;
            }
            else
            {
              var timeSlot = normal;
            }
            sendFirstCallMail(completeDateStr, timeSlot);
          }
          /******************************************************************************/
          /* FIRST MAIL : Let's see if we have any referent for the "night" time slot   */
          /******************************************************************************/
          if(doFirstCallTest(night) && (nightRefCell.getValue().length <= 2))
          {
            // Send mail
            sendFirstCallMail(completeDateStr, night);
          }

        }
        // Handle day to day mail send : cancel nigth and normal 
        else if((currentCellDate.getDate() == (todayDate.getDate())) && (currentCellDate.getMonth() == todayDate.getMonth()))
        {
          /******************************************************************************/
          /* CANCELLEATION : Let's see if we have a referent for the "normal" time slot */
          /******************************************************************************/
          if(doCancellationCallTest(normal) && (normalRefCell.getValue().length <= 2))
          {
            // Send mail : if test is done at 20h, this is a weekend timeslot
            if(todayDate.getHours() == 20)
            {
              var timeSlot = weekend;
              var dateForMail = tomorrowDate.toString();
            }
            else
            {
              var timeSlot = normal;
              var dateForMail = completeDateStr;
            }
            sendCancellationMail(dateForMail, timeSlot);
          }
          /******************************************************************************/
          /* CANCELLEATION : Let's see if we have a referent for the "night" time slot  */
          /******************************************************************************/
          if(doCancellationCallTest(night) && (nightRefCell.getValue().length <= 2))
          {
            // Send mail
            sendCancellationMail(completeDateStr, night);
          }
        }
      }

    }
  }
}

// Determine if we have to do the test depending on the date / hour / session to check
function doFirstCallTest(timeSlot)
{
  var todayDate    = new Date();
  var dayOfTheWeek = todayDate.getDay();
  var hour         = todayDate.getHours();

  if((timeSlot == normal) &&
    (((hour == 20) && (dayOfTheWeek <  5)) ||    // Normal case, from sunday to thursday -> 20h
     ((hour == 12) && (dayOfTheWeek == 5)) ||    // Friday for saturday -> 12h
     ((hour == 10) && (dayOfTheWeek == 6))))     // Saturday for sunday -> 10h
  {
    return true;
  }
  /* Noon : all the day from monday to friday */
  else if ((timeSlot == noon) && ((hour == 12) && (dayOfTheWeek < 5)))
  {
    return true;
  }
  /* Night time slot : monday, wednesday & friday */
  else if((timeSlot == night) &&
         (((hour == 20) && (dayOfTheWeek == 0)) ||    // Sunday for monday
          ((hour == 20) && (dayOfTheWeek == 2)) ||    // Thuesday for wednesday
          ((hour == 20) && (dayOfTheWeek == 4))))     // Thursday for friday
  {
    return true;
  }

  return false;
}

// Determine if we have to do the test depending on the date / hour / session to check
function doCancellationCallTest(timeSlot)
{
  var todayDate    = new Date();
  var dayOfTheWeek = todayDate.getDay();
  var hour         = todayDate.getHours();

  if((timeSlot == normal) &&
    (((hour == 14) && ((dayOfTheWeek > 0) && (dayOfTheWeek <  6))) ||    // Normal case, from monday to friday -> 14h
     ((hour == 20) && (dayOfTheWeek == 5))                         ||    // Friday for saturday -> 20h
     ((hour == 20) && (dayOfTheWeek == 6))))                             // Saturday for sunday -> 20h
  {
    return true;
  }
  /* Noon : all the day from Sunday to friday */
  else if ((timeSlot == noon) && ((hour == 20) && (dayOfTheWeek < 5)))
  {
    return true;
  }
  /* Night time slot : send mails at 14, on monday, wednesday & friday */
  else if((timeSlot == night) && ((hour == 14) && ((dayOfTheWeek == 1) || (dayOfTheWeek == 3) || (dayOfTheWeek == 5))))
  {
    return true;
  }

  return false;
}


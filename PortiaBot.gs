/*
function onOpen()

This function adds a menu item to the spreadsheet that allows you to call the main function without going into the Script Editor.

Input: none
Output: none
*/

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Initiate PortiaBot', functionName: 'main_'}
  ];
  spreadsheet.addMenu('Special Functions', menuItems);
}

/*
function main_()

This function is the main function that will call all other functions. 
This function generates an automated email of the event details, and selects volunteers based on the needs of the event. 

This function takes the first two spreadsheets as input. The first spreadsheet contains the Google Form Responses of all of the volunteers.
The second spreadsheet contains the email template to be filled out by this spreadsheet. 

The user must select the event to be processed by highlighting the event in the spreadsheet and then executing the command. 

Inputs: 
Event: Is selected as the active cell in the Google Spreadsheet
Volunteer Responses: The first Google Spreadsheet
Email Template: The Second Google Spreadsheet

Output:
Email: An email is sent to the email specified by email_address filling in the email template, and the chosen, and avaliable volunteers for the specific event.

*/

function main_() 
{
  //THE EMAIL ADDRESS THAT THE SCRIPT WILL SEND TO
  const email_address = "ltsppmac@gmail.com";
 
  //Initilize the current spreadsheet
  var ss = SpreadsheetApp.getActive();
  var activeCell = ss.getActiveCell();

  //Get the event's information
  var eventText = activeCell.getValue();
  
  //Get the active sheet
  var spreadSheetMaxRow = ss.getLastRow();
  var spreadSheetMaxColumn = ss.getLastColumn();
  var sheet = ss.getActiveSheet();
  
  //Get the event' information
  var eventData = GetEventData(sheet, eventText);
  
  //Get all of the avaliable volunteers
  var avaliableVolunteers = FindVolunteers(sheet, activeCell, spreadSheetMaxRow, spreadSheetMaxColumn);
  
  //Get all of the selcted volunteers
  var selectedVolunteers = PickVolunteers(eventData, avaliableVolunteers);
  
  //Go to the second sheet and get the Email Template.
  var templateSheet = ss.getSheets()[1];
  var emailTemplate = templateSheet.getRange("A1").getValue();
  
  //Generate the Email Subject Line to send out to the volunteers
  var emailSubject = GenerateEmailSubject(eventData);
  
  //Generate Email text to send out
  var emailText = GenerateEmail(avaliableVolunteers,selectedVolunteers, emailTemplate, eventData);

  //Send email
  MailApp.sendEmail(email_address, emailSubject, emailText, {htmlBody: emailText});
}

/*
function GenerateEmailSubject(eventData)

This function generates the Email Subject for the event. 

Input:
eventData: eventData Object containing the event's information

Output:
emailSubject: string containing the email subject line.
*/
function GenerateEmailSubject(eventData)
{
  var emailSubject = "Confirmation: ";
  emailSubject = emailSubject.concat(eventData.location);
  emailSubject = emailSubject.concat(" on ");
  emailSubject = emailSubject.concat(eventData.time);
  
  return(emailSubject);
}


/* THIS FUNCTION WAS TAKEN FROM YET ANOTHER MAIL MERGE*/
// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
//    var bic = normalizeHeaders(templateVars[i]);
    
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}

/* THIS FUNCTION WAS TAKEN FROM YET ANOTHER MAIL MERGE*/
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

/*
function PickVolunteers(eventData, avaliableVolunteers)

This function selects the approriate volunteers from the pool of avaliable volunteers. This selection is done by a first come first serve basis. 

A volunteer is chosen in the order of: 
1. They can drive
2. They have signed up for the least number of events
3. The order they have signed up

The function will iterate through untill it can fill up all of the volunteers needed, if there are no volunteers avaliable, it returns an empty list.

Inputs: 
evetnData: eventData Object containing the event's information. 
avaliableVolunteers: list of volunteer objects that contains all of the volunteers avaliable for the event

Output:
selectedVolunteers: list of volunteer objects that contains the volunteers that are picked for the event.
*/
function PickVolunteers(eventData, avaliableVolunteers)
{
  const defaultYes = "Yes"
  var volunteersAllocated = 0;
  var volunteersNeeded = eventData.volunteersNeeded;
  var listOfVolunteers = avaliableVolunteers;
  var selectedVolunteers = new Array();
  if (listOfVolunteers.length > 0)
  {
    //Check if anyone can drive, and allocate them into the volunteer pool.
    var carFound = false;
    for (var i = 0; i < listOfVolunteers.length; i++)
    {
      var volunteer = listOfVolunteers[i];
    
      if (volunteer.carAccess == "Yes" && carFound == false)
      {
        //Add Volunteer to selected Volunteers List
        selectedVolunteers.push(volunteer);
        //Make volunteer not avaliable any more. 
        listOfVolunteers.splice(i, 1);
        carFound = true; 
        volunteersNeeded --;
      }  
    }
  }
  
    //Allocate volunteers by who signed up for the least number of activites.
    for (var i = 0; i < volunteersNeeded; i ++)
    {
      if (listOfVolunteers.length > 0)
      {
        //Find the person who has the lowest number of sign ups
        var currentMin = listOfVolunteers[0].signedUpNumber;
        var volunteerIndex = 0;
        for (var j = 0; j < listOfVolunteers.length; j++)
        {
          if (listOfVolunteers[j].signedUpNumber < currentMin)
          {
            currentMin = listOfVolunteers[j].signedUpNumber;
            volunteerIndex = j;
          }
        }
        
        selectedVolunteers.push(listOfVolunteers[volunteerIndex]);
      
        listOfVolunteers.splice(volunteerIndex, 1);
      }  
    }
  return selectedVolunteers;
}

/*
function FindVolunteers( sheet, activeCell , number_of_entries, numberOfActivites)

This function iterates through the response column for the specific event and identifies the volunteers that have said yes. 
It then generates an array of volunteer objects, with the specifics of the volunteer such as Name Email, and driving avaliablity.

Inputs:
sheet: the SpreadSheet that the data is on
activeCell: the cell that is active that contains the event to be processed
number_of_entries: The number of rows to iterate through (the maximum of the spreadsheet's active space)
numberOfActivities: The number of activities on the SpreadSheet

Output:
avaliableVolunteers: An array of volunteer objects containing the volunteers that have said they can make it.
*/

function FindVolunteers( sheet, activeCell , number_of_entries, numberOfActivites)
{
  const notAvaliableText = "No"
  const defaultYes = "Yes";
  const firstNameColumn = 2;
  const emailColumn = 3;
  const formStart = 3;
  
  var avaliableVolunteers = new Array();
  var volunteerCount = 0;
  var eventColumn = activeCell.getColumn();
  
  for (var i = formStart; i < number_of_entries + 1; i++) 
  {
    var response = sheet.getRange(i, eventColumn).getValue();

    //If they say anything but No or empty text
    if(response !== notAvaliableText && response != '')
    {
      var volunteer = GetVolunteerInformation(sheet, i, eventColumn, numberOfActivites);
  
      avaliableVolunteers.push(volunteer);
      volunteerCount ++;
    }
    
  }
  
  return(avaliableVolunteers);
}

/*
function GenerateEmail( avaliableVolunteers, selectedVolunteers, emailTemplate, eventData )

This function generates the email text body that is to be sent out. 
It takes the email template, and the selected and avaliable volunteers and generates an HTML text body that can be sent out. 

Inputs:
avaliableVolunteers: The array of volunteer objects who have confirmed they can make it
selectedVolunteers: The array of volunteer objects chosen for the event
emailTemplate: The email template that will be filled out by the function
eventData: The eventData object containing the information for the specific event

Output:
EmailText: The HTML text that will be sent out
*/
function GenerateEmail( avaliableVolunteers, selectedVolunteers, emailTemplate, eventData )
{ 
  var EmailText = "";

  //Fill in Email with Email Template
  EmailText = EmailText.concat(fillInTemplateFromObject(emailTemplate, eventData) );

  //Print the volunteers that were selected
  EmailText = EmailText.concat("<br><br> Selected Volunteers  <br><br>");
    
  EmailText = EmailText.concat(PrintVolunteerInfo(selectedVolunteers) );
  
  //Print the number of avaliable volunteers
  EmailText = EmailText.concat("<br><br> Avaliable Volunteers <br><br>");

  //List all of the avaliable volunteers
  EmailText = EmailText.concat( PrintVolunteerInfo(avaliableVolunteers) );
  
  return(EmailText);
}

/*
function PrintVolunteerInfo( text, avaliableVolunteers)

This function generates the HTML table that lists the volunteers it's given and their properties in a human
readable manner. 

Inputs: 
avaliableVolunteers: The list of volunteer objects

Output:
text: An HTML Table containing the Volunteer's information

The output of the table is in the order:
| Name | Email | Sign Ups | Car Access | Comments |

Name: The name of the volunteer (Text/String)
Email: The email of the volunteer (Text/String)
Sign Ups: The number of times they have signed up for activites this month (Number)
Car Access: Does this person have access to a car that they can use for the activity? (True/False)
Comments: Anything that doesn't fit the above items, this is usually for all day activities. (Text/String)
*/

function PrintVolunteerInfo(avaliableVolunteers)
{
  var text = "";
  
  //Column Header
  text = text.concat("<table style=\"width:100%\">");
  text = text.concat("<tr>");
  text = text.concat("<th> Name </th> ");
  text = text.concat("<th> Email </th> ");
  text = text.concat("<th> Sign Ups </th>");
  text = text.concat("<th> Car Access </th>");
  text = text.concat("<th> Comments </th>");
  text = text.concat("</tr>")

  //Column Information
  for( var i = 0; i < avaliableVolunteers.length; i++)
  {
    text = text.concat("<tr>");
    var temp = avaliableVolunteers[i];
    
    text = text.concat("<td>");
    text = text.concat(temp.firstName);
    text = text.concat("</td>");
    
    text = text.concat("<td>");
    text = text.concat(temp.email);
    text = text.concat("</td>");
    
    text = text.concat("<td>");
    text = text.concat(temp.signedUpNumber);
    text = text.concat("</td>");

    text = text.concat("<td>");
    text = text.concat(temp.carAccess);
    text = text.concat("</td>");
    
    text = text.concat("<td>");
    text = text.concat(temp.comments);
    text = text.concat("</td>");
    
    text = text.concat("</tr>");
  }
  text = text.concat("</table>");

  return(text);
  
}

/*
function GetEventData( sheet, data )

This function takes the sheet and the active cell, and obtains the event's information. Then it returns
the event object. 

Input: 
sheet: The volunteer sign up spreadsheet
data: The cell containing the event information

Output:
eventData: An eventData object containing the different properties of the event (see: eventData Object)
*/
function GetEventData( sheet, data )
{
  //Location of where to find the number of volunteers needed
  const VolunteerNeededColumn = 2;
  
  var eventData = new event_information();
  
  //Split the information according to the the comma seperator
  var SplitData = data.split(",");
  
  //Get Event Location
  eventData.location = SplitData[0].toString();
  
  //Get Event Time
  eventData.time = SplitData[1].toString();
  
  //Put the rest of the string in eventData.kits.
  eventData.kits = SplitData[2].toString();
  
  //Get the number of volunteers needed
  var column = sheet.getActiveCell().getColumn();
  eventData.volunteersNeeded = sheet.getRange(VolunteerNeededColumn, column).getValue();
  
  //Remove any spacing
  eventData.time = eventData.time.substring(1, eventData.time.length);

  //Remove the space in front JavaScript sucks.
  eventData.kits = eventData.kits.substring(1, eventData.kits.length);

  //Get the location of the driving time from eventData.Kits
  var travelLocation = eventData.kits.indexOf("(");
  
  //Get Driving Time
  eventData.drivingTime = eventData.kits.substring(travelLocation + 1, eventData.kits.length - 1);
  
  //Cut kits down to exactly what the kits are.
  eventData.kits = eventData.kits.substring(0, travelLocation - 1);
  
//  Logger.log(eventData.drivingTime);

  
  return eventData;
}

/*
function GetVolunteerInformation( sheet, volunteerRowNumber, eventColumn, numberOfActivites)

This function is used by FindVolunteers, to determine a confirming volunteer's properties.

Inputs:

sheet: The volunteer sign up spread sheet (sheet)
volunteerRowNumber: The row number of the volunteer that said yes (integer)
eventColumn: The event that we are computing (integer)
numberOfActivities: The number of activiites this month (integer)

Output:
volunteer: A volunteer object containing the properties of this volunteer
*/

function GetVolunteerInformation( sheet, volunteerRowNumber, eventColumn, numberOfActivites)
{
  var volunteer = new Volunteer( "NULL" , "NULL", "NULL", "FALSE", 0, "NULL");
  const notAvaliableText = "No";
  const defaultYes = "Yes";
  
  const nameColumn = 2;
  const emailColumn = 3;
  const carColumn = 4;
  
  //Get Volunteer's Name and Email and Car Access
  volunteer.firstName = sheet.getRange(volunteerRowNumber, nameColumn).getValue();
  volunteer.email = sheet.getRange(volunteerRowNumber, emailColumn).getValue();
  volunteer.carAccess = sheet.getRange(volunteerRowNumber, carColumn).getValue();
  
  //If they say anything but Yes, add thier response to the comments section
  var volunteerResponse = sheet.getRange(volunteerRowNumber, eventColumn).getValue();
  if (volunteerResponse !== defaultYes && volunteerResponse != '')
  {
    volunteer.comments = volunteerResponse;
  }
  
  //Check how many events they signed up for
  var count = 0;
  for ( var i = eventColumn; i < numberOfActivites; i++)
  {
    var response = sheet.getRange(volunteerRowNumber, i).getValue();
    //If they say anything but No or empty text
    if(response !== notAvaliableText && response != '')
    {
      count ++;
    }
  }
  volunteer.signedUpNumber = count;
  
  return(volunteer);
}

/*
function Volunteer(firstName, lastName, email, carAccess, signedUpNumber, comments)

This is the volunteer object that is used for every confirming volunteer. This volunteer object will hold 
all of the charastics of the volunteer such as

firstName (Text/String)
lastName (Text/String)
email (Text/String)
carAccess (Boolean or True/False)
signedUpNumber (Integer)
comments (Text/String)
*/

function Volunteer(firstName, lastName, email, carAccess, signedUpNumber, comments) 
{
  this.firstName = firstName;
  this.lastName = lastName;
  this.email = email;
  this.carAccess = carAccess;
  this.signedUpNumber = signedUpNumber;
  this.comments = comments;
}

/*
function event_information (location, time, kits, volunteersNeeded, drivingTime)

This is the event information constructor function. it constructs the event information object. 

All of the event's information is stored in the event information object, that other functions can access.

The feilds are:

location: The location of the event (Text/String)
time: The time of the event (Text/String)
Kits: The kits that will be used for the event (Text/String)
volunteersNeeded: The number of volunteers that are needed for this event (Text/String)
drivingTime: The distance it is to drive to this event stored as "XX min drive" (Text/String)
*/

function event_information (location, time, kits, volunteersNeeded, drivingTime) 
{
  this.location = location;
  this.time = time;
  this.kits = kits;
  this.volunteersNeeded = volunteersNeeded;
  this.drivingTime = drivingTime;
}

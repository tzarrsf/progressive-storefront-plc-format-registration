//Set this to true to send emails to the replyTo address instead of the one in the sheet (for testing)
const debugMode = false;
const productName = "PWA Web Kit & Managed Runtime";
const eventName = "Training Environment Request";
const sheetName = "Form Responses 2";
const fullNameHeader = "Full Name"; // JT: adding this here because I removed from form and replaced with First / Last Name fields (to make eventual automation work)
const actionLink = "https://docs.google.com/spreadsheets/d/1rnrHNKdJcEyxs37C7tNxezMbV6cccpWQniOElh5jmSU/edit?resourcekey#gid=1646972121";
//TODO: Update the playbookLink or replace this with ~FULL AUTOMATION~!
const repositoryLink = "https://github.com/tzarrsf/account_manager_and_mrt_user_setup";
const windowInDays = 30;
const yes = "yes";

//NOTE: approverEmailTo and approverEmailCc can be comma-delimited lists if you need multiple people involved for coverage / awareness
const approverEmailTo = "jonathan.tucker@salesforce.com";
const approverEmailCc = "";
const replyTo = "jonathan.tucker@salesforce.com";

/*
  Assumed header text values in use:
  - Timestamp (this is actually used in a calculation against the windowInDays for detecting duplicates)
  - Email Address
  - I'm not LIVE
  - First Name
  - Last Name
  - Partner company name
  - Can you commit blah blah blah...
Mapping all columns by index makes debugging Apps Script easier and can be used to get prior data for duplicate checks, etc.
*/

//Column mapping for: Timestamp
const timestampHeader = "Timestamp";
const timestampColumnIndex = getHeaderIndex(sheetName, timestampHeader);
//Column mapping for: Email Address
const emailAddressHeader = "Email Address";
const emailAddressColumnIndex = getHeaderIndex(sheetName, emailAddressHeader);
//Column mapping for: Partner company name
const partnerCompanyNameHeader = "Partner company name";
const partnerCompanyNameColumnIndex = getHeaderIndex(sheetName, partnerCompanyNameHeader);
//Column mapping for: First Name
const firstNameHeader = "First Name";
const firstNameColumnIndex = getHeaderIndex(sheetName, firstNameHeader);
//Column mapping for: Last Name
const lastNameHeader = "Last Name";
const lastNameColumnIndex = getHeaderIndex(sheetName, lastNameHeader);
//Column mapping for: Country
//const countryHeader = "Country";
//const countryColumnIndex = getHeaderIndex(sheetName, countryHeader);
//Column mapping for: Can you commit to completing all of the PWA Web Kit and Managed Runtime courses within 30 days of receiving the environment access?
//const canYouCommitHeader = "Can you commit to completing all of the PWA Web Kit and Managed Runtime courses within 30 days of receiving the environment access?";
//const canYouCommitColumnIndex = getHeaderIndex(sheetName, canYouCommitHeader);
//Col mapping for Live (it's a checkbox with only 1 value, last ditch effort to stop people from doing this during live training)
const liveHeader = "I am taking the self-paced PLC Course, I am NOT registered for a Live training (Do not submit this form if you are registered for a live session and have come to this form from Prework)."
const liveColumnIndex = getHeaderIndex(sheetName, liveHeader);

// Takes a date from the sheet in a format like this and gives you whole days elapsed since that time: "3/15/2022 15:52:11"
function dateDiffDays(dateString)
{
  let lastDateEpoch = Date.parse(dateString);
  let currentDateEpoch = Date.now();
  let dateDeltaInDays = Math.floor((currentDateEpoch - lastDateEpoch)) / (1000 * 60 * 60 * 24);
  return dateDeltaInDays;
}

function onFormSubmit(e)
{
  // Get the row of submitted data using the mapped out column indexes
  let submission = {
    emailAddress: e.values[emailAddressColumnIndex]
    , partnerCompanyName: e.values[partnerCompanyNameColumnIndex]    
    , fullName: e.values[firstNameColumnIndex] + ' ' + e.values[lastNameColumnIndex]
  };
  
  //You can avoid mistakes or spamming when first assigning this code to your form using the debugMode variable at the top of the code
  if(debugMode)
  {
    submission.emailAddress = replyTo;
  }
  
  //Check for duplicates - if we have a record on file for the email address within 30 days of the first submmission it's a dup
  //let alreadyRegisteredResult = isDuplicateRegistration(submission.emailAddress, submission.partnerCompanyName, submission.canYouCommit);
  let alreadyRegisteredResult = isDuplicateRegistration(submission.emailAddress);
  //let alreadyRegisteredResult = false; // issues with this after edit
  
  Logger.log(JSON.stringify(alreadyRegisteredResult));

  if(alreadyRegisteredResult.alreadyRegistered === true)
  {
    //Duplicate request email to the person registering
    MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: "Dear " + submission.fullName + ",<br /><br />Thanks for submitting your <b>" + productName + " " + eventName + "</b>. "
      + "Please note that this looks like a duplicate request and it will be reviewed for approval as soon as possible if the request has not been fulfilled already.<br /><br />"
      + "Please also note that this form is not meant for extending the access needed to pursue the " + productName + " training. As a matter of policy, we can only provide this "
      + "access to each partner every 30 days. If you need a sandbox for a longer period, you should use your own sandbox or request an On Demand Sandbox from your customer.<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Please allow 48 hours for review, approval and provisioning to take place.<br />"
    });
    Logger.log("Sent email for duplicate request: " + productName + " " + eventName + " to <" + submission.emailAddress + ">");
    
    //Duplicate request notification to approver(s)
    MailApp.sendEmail({
      to: approverEmailTo
      , cc: approverEmailCc
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: submission.fullName + " submitted a duplicate <b>" + productName + "</b> " + eventName + ". Review may be needed.<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "<a href=\"" + actionLink + "\">Click here</a> to review and action the request or <a href=\"" + repositoryLink + "\">click here</a> for the tool repository.<br />"
    });
  }
  else
  {
    // NOTE: If your form allows editing, you will get an exception for not having a recipient here when editing - so don't allow editing on your form settings ;)
    
    // Only provide environment access to those who can commit to completing courses in the 30 days
    if(submission !== null)
    {
      // Initial request email to the person registering
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: productName + " " + eventName + " (Initial)"
        , htmlBody: "Dear " + submission.fullName + ",<br /><br /> Thanks for submitting your <b>" + productName + "</b> " + eventName + ". "
        + "We now have the following data on file:<br /><br />"
        + makeRegistrationTable(submission) + "<br />"
        + "Please allow 48 hours for review and provisioning.<br />"
      });

      // Initial request email to the approver or call to automation should happen here
      // TODO: Figure out how to action these with calls to the python scripts
      MailApp.sendEmail({
        to: approverEmailTo
        , cc: approverEmailCc
        , replyTo: replyTo
        , subject: "Action Required: " + productName + " " + eventName + " Review (New)"
        , htmlBody: submission.fullName + " submitted a new <b>" + productName + "</b> " + eventName + ". <u>Review and actioning is needed</u>:<br /><br />"
        + makeRegistrationTable(submission) + "<br /><br />"
        + "<a href=\"" + actionLink + "\">Click here</a> to review and action the request or <a href=\"" + repositoryLink + "\">click here</a> for the tool repository.<br />"
      });
      Logger.log("Sent email for initial request: " + productName + " " + eventName + " to approvers (To: " + approverEmailTo + " Cc: " + approverEmailCc + ")");
    }
    else
    {
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: productName + " " + eventName + " (Initial)"
        , htmlBody: "Dear " + submission.fullName + ",<br /><br /> Thanks for submitting your <b>" + productName + "</b> " + eventName + ".<br /><br />" + 
        "Based on your responses we are unable to provision the access for you at this time."
      });
      Logger.log("Sent denial email for initial request: " + productName + " " + eventName + " to <" + submission.emailAddress + ">");
    }
  }
}

function getHeaderIndex(sheetName, headerText)
{
  let result = -1;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if(sheet == null)
  {
    console.log('Sheet not located in getIndexByHeader using: ' + sheetName);
    return result;
  }

  let data = sheet.getDataRange().getValues();

  for (let i=0; i < data[0].length; i++)
  {
    //Logger.log("Left: '" + data[0][i].toString() + "' Right: '" + headerText + "'");
    if(data[0][i].toString() == headerText)
    {
      result = i;
      break;
    }
  }
  return result;
}

// Makes an old-school HTML table of name-value pairs from the object
function makeRegistrationTable(data)
{
  Logger.log("makeRegistrationTable(" + JSON.stringify(data) + ") invoked.");
  let registrationTable = "";
  registrationTable += "<b>Request Type</b>: " + productName + " " + eventName  + "<br />";
  registrationTable += "<b>" + partnerCompanyNameHeader + "</b>: " + data.partnerCompanyName + "<br />";
  registrationTable += "<b>" + fullNameHeader + "</b>: " + data.fullName + "<br />";
  return registrationTable;
}

// Detect duplicates based on composite of the Email Address, Partner company name, canYouCommit = yes, day delta within windowInDays 
//function isDuplicateRegistration(emailAddress, partnerCompanyName, canYouCommit)
function isDuplicateRegistration(emailAddress)
{
  Logger.log("isDuplicateRegistration(" + emailAddress);
  //The current submission does not get stopped so we need to just track it and kill the real dup looking at any matches added beyond a length of 1
  let hits = [];
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 1;
  let withinWindow = false;
   
  // Start loop at 1 to ignore the header
  for (i = 1; i < data.length; i++)
  {
    withinWindow = dateDiffDays(data[i][timestampColumnIndex].toString()) <= windowInDays;

    //If we have a matching email + matching partner name + we are still in the window it's a duplicate
    if (data[i][emailAddressColumnIndex].toString().toLowerCase() === emailAddress.toLowerCase()
    && withinWindow)
    {
      hits.push(data[i]);
    }
  }
  
  // Return some enriched results we can use with a boolean and reuse the data in responses
  if(hits.length > 1)
  {
    Logger.log("Located a duplicate based on partner email address: <" + emailAddress + "> within window of: " + windowInDays +
    " days.");
    //Remove the newest addition using the tracking array (like it never happened)
    sheet.deleteRow(i);
    //Get the data for the original (old) submission into something we can use
    let originalRow = hits[hits.length - 2];
    let priorRegistration = {
      emailAddress: originalRow[emailAddressColumnIndex]
      , partnerCompanyName: originalRow[partnerCompanyNameColumnIndex]    
      , fullName: originalRow[firstNameColumnIndex] + ' ' + originalRow[lastNameColumnIndex]
    };
    Logger.log("Located duplicate partner email: <" + emailAddress + "> within window of: " + windowInDays + ".");
    return JSON.parse("{\"alreadyRegistered\": true, \"priorData\": " + JSON.stringify(priorRegistration) +"}");
  }
  else
  {
    Logger.log("Did not find duplicate email: <" + emailAddress + "> within window of: " + windowInDays + ".");
    return JSON.parse("{\"alreadyRegistered\": false, \"priorData\": null}");
  }
}

function main() {
  var wb = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = wb.getSheetByName('Final Mark Review 3');
  var sheet2 = wb.getSheetByName('Sheet7');

  // Access the data in Sheet1
  var data = sheet1.getRange(2, 1, sheet1.getLastRow() - 1, sheet1.getLastColumn()).getDisplayValues();

  // Access all additional data from Sheet2
  var additionalDataRange = sheet2.getRange(2, 1, sheet2.getLastRow() - 1, 2).getDisplayValues(); // Assuming data starts from Row 2
  var additionalDataArray = [];

  // Assuming Column A in Sheet2 has names and Column B has marks
  for (var i = 0; i < additionalDataRange.length; i++) {
    var name = additionalDataRange[i][0];
    var marks = additionalDataRange[i][1];
    additionalDataArray.push({name: name, marks: marks});
  }

  // Loop through each row of data in Sheet1 to send individual emails
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var name = row[1];   // Column B (second column)
    var marks = row[3];  // Column D (fourth column)
    var email = row[4];  // Column E (fifth column) – Email Address

    // Create template object for dynamically constructing HTML
    var htmlTemplate = HtmlService.createTemplateFromFile('html');

    // Define HTML variables
    htmlTemplate.name = name;
    htmlTemplate.marks = marks;
    htmlTemplate.additionalData = additionalDataArray; // Include all additional data

    // Evaluate the template and get content
    var htmlForEmail = htmlTemplate.evaluate().getContent();

    if (email) {
      GmailApp.sendEmail(
        email,
        'Marks ' + name,
        'This email contains HTML',
        {htmlBody: htmlForEmail}
      );
      Logger.log('Email sent to ' + email);
    } else {
      Logger.log('No email address found for row ' + (i + 2));
    }
  }
  
  // Call the second function to send emails for additional data in table format
  sendAdditionalDataEmails();
}

// Module to send emails with only additional data from Sheet2 in table format
function sendAdditionalDataEmails() {
  var wb = SpreadsheetApp.getActiveSpreadsheet();
  var sheet2 = wb.getSheetByName('Sheet7');

  // Access all additional data from Sheet2
  var additionalDataRange = sheet2.getRange(2, 1, sheet2.getLastRow() - 1, 2).getDisplayValues();
  
  // Access emails (assuming they are in Column C)
  var emailRange = sheet2.getRange(1, 3, sheet2.getLastRow() - 1, 1).getValues(); // Column C holds email addresses
  
  for (var i = 0; i < additionalDataRange.length; i++) {
    var name = additionalDataRange[i][0];  // Column A (Name)
    var marks = additionalDataRange[i][1]; // Column B (Marks)
    var email = emailRange[i][0];          // Column C (Email Address)

    // Skip rows without email
    if (!email || email.trim() === "") {
      Logger.log('No email address found for row ' + (i + 2) + ' in Sheet2');
      continue; // Skip this iteration if no email is found
    }

    // Create the HTML table structure for the additional data
    var emailBody = '<p>Hello,</p>';
    emailBody += '<p>Review 3 Marks:</p>';
    emailBody += '<table border="1" cellpadding="5" cellspacing="0">';
    emailBody += '<tr><th>Name</th><th>Marks</th></tr>';  // Table header

    // Loop through all additional data and add to the table
    for (var j = 0; j < additionalDataRange.length; j++) {
      emailBody += '<tr>';
      emailBody += '<td>' + additionalDataRange[j][0] + '</td>'; // Name (Column A)
      emailBody += '<td>' + additionalDataRange[j][1] + '</td>'; // Marks (Column B)
      emailBody += '</tr>';
    }
    
    emailBody += '</table>';
    emailBody += '<p>Best regards,</p>';

    // Send the email with the table as the body
    try {
      GmailApp.sendEmail(
        email,
        'Class Marks - ' ,
        'This email contains HTML',  // Plain text fallback
        {htmlBody: emailBody}
      );
      Logger.log('Additional data email sent to ' + email);
    } catch (error) {
      Logger.log('Error sending email to ' + email + ': ' + error.message);
    }
  }
}

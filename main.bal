import ballerina/io;
import ballerina/lang.runtime;
import ballerinax/googleapis.gmail;
import ballerinax/googleapis.sheets;

configurable string authToken = ?;
configurable string spreadsheetId = ?;
configurable string sheetName = ?;
configurable string dataRange = ?;
configurable string templatePath = ?;
configurable int nameColumnNumber = 1;
configurable int emailColumnIndex = 2;
configurable int certificateUrlColumnIndex = 3;

// Setup Google Sheets and Gmail clients using the provided authentication token
sheets:Client sheetsClient = check new ({
    auth: {
        token: authToken
    }
});

gmail:Client gmailClient = check new ({
    auth: {
        token: authToken
    }
});

public function main() returns error? {
    do {

        // Extract recipient data from Google Sheets
        sheets:Range|error range = sheetsClient->getRange(spreadsheetId,sheetName,dataRange);
        if range is error {
            io:println("Error accessing Google Sheets: ", range.message());
            return range;
        }

        // Check if the range contains values and validate the data
        (int|string|decimal)[][]? values = range.values;
        if values is () {
            io:println("No data found in the sheet");
            return;
        }

        if values.length() == 0 {
            io:println("Sheet is empty, no recipients to process");
            return;
        }

        string templateContent = check io:fileReadString(templatePath);

        // Validate, process, and send emails
        int sent = 0;
        int skipped = 0;

        foreach var row in values {
            if row.length() <= nameColumnNumber || row.length() <= emailColumnIndex || row.length() <= certificateUrlColumnIndex {
                io:println("Skipping invalid row (insufficient columns): ", row);
                skipped += 1;
                continue;
            }

            (int|string|decimal) nameCell = row[nameColumnNumber];
            (int|string|decimal) emailCell = row[emailColumnIndex];
            (int|string|decimal) certificateUrlCell = row[certificateUrlColumnIndex];
            
            string nameValue = nameCell.toString().trim();
            string emailValue = emailCell.toString().trim();
            string certificateUrlValue = certificateUrlCell.toString().trim();

            if nameValue == "" || emailValue == "" || certificateUrlValue == "" {
                io:println("Skipping row with empty name, email, or certificate URL: ", row);
                skipped += 1;
                continue;
            }

            // Render and send email
            string emailBody = renderTemplate(templateContent, nameValue, certificateUrlValue);

            gmail:MessageRequest emailMessage = {
                to: [emailValue],
                subject: "Certificate of Participation: Innovate with Ballerina 2025",
                bodyInHtml: emailBody
            };

            gmail:Message|error sendResult = gmailClient->/users/me/messages/send.post(emailMessage);
            if sendResult is error {
                io:println("Failed to send email to ",nameValue, " <", emailValue, ">: ", sendResult.message());
                skipped += 1;
            } else {
                io:println("Email sent successfully to ", nameValue, " <", emailValue, ">");
                sent += 1;
            }

            // Rate limiting: 750ms delay between sends (80 emails/min, safely under 250/min limit)
            runtime:sleep(0.75);
        }

        // Print summary
        io:println(string `=== SUMMARY ===`);
        io:println(string `Total rows processed: ${values.length()}`);
        io:println(string `Emails sent: ${sent}`);
        io:println(string `Emails skipped: ${skipped}`);

    } on fail var e {
        return e;
    }
}


// Used to render the email template by replacing placeholders with recipient data.
function renderTemplate(string templateContent, string recipientName, string certificateUrl) returns string {
    string escapedName = escapeHtml(recipientName);
    
    string:RegExp recipientNamePattern = re `\{\{recipientName\}\}`;
    string renderedContent = recipientNamePattern.replaceAll(templateContent, escapedName);
    
    string:RegExp certificateUrlPattern = re `\{\{certificateUrl\}\}`;
    renderedContent = certificateUrlPattern.replaceAll(renderedContent, certificateUrl);
    
    return renderedContent;
}


// Reject potentially unsafe characters in the recipient's name to prevent HTML injection in the email body.
function escapeHtml(string input) returns string {
    string:RegExp ampersandPattern = re `&`;
    string escaped = ampersandPattern.replaceAll(input, "&amp;");
    
    string:RegExp lessThanPattern = re `<`;
    escaped = lessThanPattern.replaceAll(escaped, "&lt;");
    
    string:RegExp greaterThanPattern = re `>`;
    escaped = greaterThanPattern.replaceAll(escaped, "&gt;");
    
    string:RegExp doubleQuotePattern = re `"`;
    escaped = doubleQuotePattern.replaceAll(escaped, "&quot;");
    
    string:RegExp singleQuotePattern = re `'`;
    escaped = singleQuotePattern.replaceAll(escaped, "&#x27;");
    
    return escaped;
}


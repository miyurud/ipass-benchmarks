import wso2/sfdc46;
import ballerina/config;
import ballerina/io;
import ballerina/log;
import ballerina/time;
import ballerina/stringutils;
import wso2/msspreadsheets;
import wso2/msonedrive;
import wso2/twilio;


sfdc46:SalesforceConfiguration salesforceConfig = {
    baseUrl: config:getAsString("BASE_URL"),
    clientConfig: {
        accessToken: config:getAsString("ACCESS_TOKEN"),
        refreshConfig: {
            clientId: config:getAsString("CLIENT_ID"),
            clientSecret: config:getAsString("CLIENT_SECRET"),
            refreshToken: config:getAsString("REFRESH_TOKEN"),
            refreshUrl: config:getAsString("REFRESH_URL")
        }
    }
};

sfdc46:Client salesforceClient = new(salesforceConfig);

// Create Microsoft Graph Client configuration by reading from config file.
msspreadsheets:MicrosoftGraphConfiguration msGraphConfig = {
    baseUrl: config:getAsString("MS_EP_URL"),
    bearerToken: config:getAsString("BEARER_TOKEN"),

    clientConfig: {
        accessToken: config:getAsString("MS_ACCESS_TOKEN"),
        refreshConfig: {
            clientId: config:getAsString("MS_CLIENT_ID"),
            clientSecret: config:getAsString("MS_CLIENT_SECRET"),
            refreshToken: config:getAsString("MS_REFRESH_TOKEN"),
            refreshUrl: config:getAsString("MS_REFRESH_URL")
        }
    }
};

twilio:TwilioConfiguration twilioConfig = {
    accountSId: config:getAsString("TWILIO_ACCOUNT_SID"),
    authToken: config:getAsString("TWILIO_AUTH_TOKEN"),
    xAuthyKey: config:getAsString("TWILIO_AUTHY_API_KEY")
};
twilio:Client twilioClient = new(twilioConfig);

int SIX_HOURS_IN_MILISECONDS = 21600000;
string WORK_BOOK_NAME = "Book";
string WORK_SHEET_NAME = "ABC";
string TWILIO_SANDBOX_NUMBER = "+14155238886";
string DESTINATION_PHONE_NUMBER = "+94775544041";

public function main() returns @tainted error? {
    time:Time startTime = time:currentTime();
    // Create Salesforce bulk client.
    sfdc46:SalesforceBulkClient sfBulkClient = salesforceClient->createSalesforceBulkClient();
    
    //Construct Salesforce query
    time:Time time = time:currentTime();
    int currentTimeMills = time.time;
    int timeSixHoursBefore = currentTimeMills - SIX_HOURS_IN_MILISECONDS;
    io:println("Current system time in milliseconds: ", currentTimeMills);
    io:println("System time six hours before in milliseconds: ", timeSixHoursBefore);
    
    time:TimeZone zoneIdValue = time.zone;
    time:Time time1 = { time: currentTimeMills, zone: zoneIdValue };
    time:Time time2 = { time: timeSixHoursBefore, zone: zoneIdValue };
    
    string|error cString1 = time:format(time1, "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
    string customTimeString = "";
    if (cString1 is string) {
       customTimeString = cString1;
       io:println("Current system time in custom format: ", customTimeString);
    }
    
    string|error cString2= time:format(time2, "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
    string customTimeString2 = "";
    if (cString2 is string) {
       customTimeString2 = cString2;
       io:println("Six hours less time in custom format: ", customTimeString2);
    }
    
    string sampleQuery = "SELECT o.Id, o.CreatedDate, o.AccountId, o.CloseDate FROM Opportunity o WHERE o.CreatedDate > customTimeString2 and o.CreatedDate < customTimeString";
    sampleQuery= stringutils:replace(sampleQuery,"customTimeString2", customTimeString2);
    sampleQuery= stringutils:replace(sampleQuery,"customTimeString", customTimeString);
    io:println("sampleQuery: ", sampleQuery);
    sfdc46:SoqlResult|sfdc46:ConnectorError response = salesforceClient->getQueryResult(<@untainted> sampleQuery);
    
    //Create Microsoft live spreadsheet client
    msspreadsheets:MSSpreadsheetClient msSpreadsheetClient = new(msGraphConfig);
    
    if (response is sfdc46:SoqlResult) {
       io:println("TotalSize:  ", response.totalSize.toString());
       io:println("Done:  ", response.done.toString());
       int totalNumberOfrecords = response.totalSize;

        if (totalNumberOfrecords > 0) {
            //May be we will have to create a new workbook
            boolean|error result = msSpreadsheetClient->deleteWorksheet(WORK_BOOK_NAME, WORK_SHEET_NAME);
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error deleting worksheet", err = result);
            }

            result = msSpreadsheetClient->createWorksheet(WORK_BOOK_NAME, WORK_SHEET_NAME);
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error creating worksheet", err = result);
            }

            result = msSpreadsheetClient->createTable(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", <@untainted> ("A" + totalNumberOfrecords.toString() + ":D" + totalNumberOfrecords.toString()));
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error creating table", err = result);
            }

            result = msSpreadsheetClient->setTableheader(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", 1, "Id");
            result = msSpreadsheetClient->setTableheader(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", 2, "CreatedDate");
            result = msSpreadsheetClient->setTableheader(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", 3, "AccountId");
            result = msSpreadsheetClient->setTableheader(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", 4, "CloseDate");

            json[][] valuesString=[];
            int counter = 0;
            foreach var line in response.records {
                io:println("Line:  ", line.get("Id"), ", ", line.get("CreatedDate"),  ", ", line.get("AccountId"), ", ", line.get("CloseDate"));
                json[] arr = [ line.get("Id").toString(), line.get("CreatedDate").toString(), line.get("AccountId").toString(), line.get("CloseDate").toString() ];
                valuesString.push(arr);
                counter += 1;
            }

            json data = {"values": valuesString};

            //Write the processed Salesforce data to Microsoft Live spreadsheet
            result = msSpreadsheetClient->insertDataIntoTable(WORK_BOOK_NAME, WORK_SHEET_NAME, "tableOpportunities", <@untainted> data);
            io:println(result);

            msonedrive:OneDriveClient msOneDriveClient = new(msGraphConfig);

            var workBookResponse = msOneDriveClient->getItemURL(WORK_BOOK_NAME + ".xlsx");
            if (workBookResponse is string) {
                io:println(workBookResponse);
                var whatsAppResponse = twilioClient->sendWhatsAppMessage("whatsapp:" + TWILIO_SANDBOX_NUMBER, "whatsapp:" + DESTINATION_PHONE_NUMBER, "Your URL code is " + workBookResponse);

                if (whatsAppResponse is twilio:WhatsAppResponse) {
                    io:println("Message sent");
                } else {
                    log:printError("Error sending the WhatsApp message", err = whatsAppResponse);
                }

            } else {
                log:printError("Error getting the WorkBook URL", err = workBookResponse);
            }
        }
    } else {
       io:println("Error: ", response.detail()?.message.toString());
    }

    time:Time endTime = time:currentTime();
    int elapsedTime = endTime.time - startTime.time;
    io:println("Elapsed time (ms): ", elapsedTime.toString());
}
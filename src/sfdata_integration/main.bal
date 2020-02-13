import wso2/sfdc46;
import ballerina/config;
import ballerina/io;
import ballerina/log;
import ballerina/time;
import ballerina/stringutils;
import wso2/msspreadsheets;


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

int SIX_HOURS_IN_MILISECONDS = 21600000;

public function main() returns @tainted error? {
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
    msspreadsheets:MSSpreadsheetClient msGraphClient = new(msGraphConfig);
    
    if (response is sfdc46:SoqlResult) {
       io:println("TotalSize:  ", response.totalSize.toString());
       io:println("Done:  ", response.done.toString());
       int totalNumberOfrecords = response.totalSize;

        if (totalNumberOfrecords > 0) {
            //May be we will have to create a new workbook
            boolean|error result = msGraphClient->deleteWorksheet("Book", "ABC");
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error deleting worksheet", err = result);
            }

            result = msGraphClient->createWorksheet("Book", "ABC");
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error creating worksheet", err = result);
            }

            result = msGraphClient->createTable("Book", "ABC", "tableOpportunities", <@untainted> ("A" + totalNumberOfrecords.toString() + ":D" + totalNumberOfrecords.toString()));
            if (result is boolean) {
                io:println(result);
            } else {
                log:printError("Error creating table", err = result);
            }

            result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 1, "Id");
            result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 2, "CreatedDate");
            result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 3, "AccountId");
            result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 4, "CloseDate");

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
            result = msGraphClient->insertDataIntoTable("Book", "ABC", "tableOpportunities", <@untainted> data);
            io:println(result);
        }
    } else {
       io:println("Error: ", response.detail()?.message.toString());
    }
}
const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const MergeAndDeliverViaWebhookParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/merge_and_deliver_via_webhook_parameters").MergeAndDeliverViaWebhookParameters;
const MailMergeWebhookSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/mail_merge_webhook_settings").MailMergeWebhookSettings;

const MergeAndDeliverViaWebhookSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/merge_and_deliver_via_webhook_success_response").MergeAndDeliverViaWebhookSuccessResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class MergeAndDeliver {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var parameters = new MergeAndDeliverViaWebhookParameters();

            //parameters.setFileUrl("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");

            var fileName = "OfferLetter.zdoc";
            var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            parameters.setFileContent(streamWrapper);

            parameters.setOutputFormat("zdoc");
            parameters.setMergeTo("separatedoc");

            parameters.setPassword("***");

            var jsonFileName = "candidates.json";
            var jsonFilePath = __dirname + "/sample_documents/candidates.json";
            var jsonFileStream = fs.readFileSync(jsonFilePath);
            var jsonStreamWrapper = new StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            parameters.setMergeDataJsonContent(jsonStreamWrapper);

            var webhookSettings = new MailMergeWebhookSettings();

            webhookSettings.setInvokeUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");
            webhookSettings.setInvokePeriod("oncomplete");

            parameters.setWebhook(webhookSettings);
            
            //var mergeData = new Map();
            //parameters.setMergeData(mergeData);

            //var csvFileName = "csv_data_source.csv";
            //var csvFilePath = "/Users/praba-2086/Desktop/csv_data_source.csv";
            //var csvFileStream = fs.readFileSync(csvFilePath);
            //var csvStreamWrapper = new StreamWrapper(csvFileName, csvFileStream, csvFilePath);

            //parameters.setMergeDataCsvContent(csvStreamWrapper);
            
            //parameters.setMergeDataCsvUrl("https://demo.office-integrator.com/data/csv_data_source.csv");
            //parameters.setMergeDataJsonUrl("https://demo.office-integrator.com/zdocs/json_data_source.json");

            var responseObject = await sdkOperations.mergeAndDeliverViaWebhook(parameters);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof MergeAndDeliverViaWebhookSuccessResponse) {
                        console.log("Records - " + writerResponseObject.getRecords());
                        console.log("Merge report url - " + writerResponseObject.getMergeReportDataUrl());
                    } else if (writerResponseObject instanceof InvaildConfigurationException) {
                        console.log("Invalid configuration exception. Exception json - ", writerResponseObject);
                    } else {
                        console.log("Request not completed successfullly");
                    }
                }
            }
        } catch (error) {
            console.log("Exception while running sample code", error);
        }
    }
}

MergeAndDeliver.execute();
const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const MailMergeWebhookSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/mail_merge_webhook_settings").MailMergeWebhookSettings;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;
const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;
const MergeAndDeliverViaWebhookParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/merge_and_deliver_via_webhook_parameters").MergeAndDeliverViaWebhookParameters;
const MergeAndDeliverViaWebhookSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/merge_and_deliver_via_webhook_success_response").MergeAndDeliverViaWebhookSuccessResponse;

class MergeAndDeliver {

    //Include zoi-nodejs-sdk package in your package json and the execute this code.

    static async initializeSdk() {
        let user = new UserSignature("john@zylker.com");
        let environment = new Environment("https://api.office-integrator.com", null, null);
        let apikey = new APIKey("2ae438cf864488657cc9754a27daa480", Constants.PARAMS);
        let logger = new LogBuilder()
            .level(Levels.INFO)
            .filePath("./app.log")
            .build();
        let initialize = await new InitializeBuilder();

        await initialize.user(user).environment(environment).token(apikey).logger(logger).initialize();

        console.log("\nSDK initialized successfully.");
    }

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var parameters = new MergeAndDeliverViaWebhookParameters();

            parameters.setFileUrl("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");
            parameters.setMergeDataJsonUrl("https://demo.office-integrator.com/data/candidates.json");

            // var fileName = "OfferLetter.zdoc";
            // var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            // parameters.setFileContent(streamWrapper);

            parameters.setOutputFormat("zdoc");
            parameters.setMergeTo("separatedoc");

            parameters.setPassword("***");

            // var jsonFileName = "candidates.json";
            // var jsonFilePath = __dirname + "/sample_documents/candidates.json";
            // var jsonFileStream = fs.readFileSync(jsonFilePath);
            // var jsonStreamWrapper = new StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            // parameters.setMergeDataJsonContent(jsonStreamWrapper);

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
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof MergeAndDeliverViaWebhookSuccessResponse) {
                        console.log("\nRecords - " + JSON.stringify(writerResponseObject.getRecords()));
                        console.log("\nMerge report data url - " + writerResponseObject.getMergeReportDataUrl());
                    } else if (writerResponseObject instanceof InvaildConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", writerResponseObject);
                    } else {
                        console.log("\nRequest not completed successfullly");
                    }
                }
            }
        } catch (error) {
            console.log("\nException while running sample code", error);
        }
    }
}

MergeAndDeliver.execute();
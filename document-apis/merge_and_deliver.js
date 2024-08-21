import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync, writeFileSync } from 'fs';
const __dirname = import.meta.dirname;

class MergeAndDeliver {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var parameters = new SDK.V1.MergeAndDeliverViaWebhookParameters();

            //Either use url as document source or attach the document in request body use below methods
            parameters.setFileUrl("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");
            parameters.setMergeDataJsonUrl("https://demo.office-integrator.com/data/candidates.json");

            // var fileName = "OfferLetter.zdoc";
            // var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            // var fileStream = readFileSync(filePath);
            // var streamWrapper = new SDK.StreamWrapper(fileName, fileStream, filePath);
            
            // parameters.setFileContent(streamWrapper);

            // var jsonFileName = "candidates.json";
            // var jsonFilePath = __dirname + "/sample_documents/candidates.json";
            // var jsonFileStream = readFileSync(jsonFilePath);
            // var jsonStreamWrapper = new SDK.StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            // parameters.setMergeDataJsonContent(jsonStreamWrapper);

            parameters.setOutputFormat("zdoc");
            parameters.setMergeTo("separatedoc");
            parameters.setPassword("***");
            
            var webhookSettings = new SDK.V1.MailMergeWebhookSettings();

            webhookSettings.setInvokeUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");
            webhookSettings.setInvokePeriod("oncomplete");

            parameters.setWebhook(webhookSettings);

            //var mergeData = new Map();
            //parameters.setMergeData(mergeData);

            //var csvFileName = "csv_data_source.csv";
            //var csvFilePath = __dirname + "/sample_documents/csv_data_source.csv";
            //var csvFileStream = readFileSync(csvFilePath);
            //var csvStreamWrapper = new SDK.StreamWrapper(csvFileName, csvFileStream, csvFilePath);

            //parameters.setMergeDataCsvContent(csvStreamWrapper);
            //parameters.setMergeDataCsvUrl("https://demo.office-integrator.com/data/csv_data_source.csv");

            var responseObject = await sdkOperations.mergeAndDeliverViaWebhook(parameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof SDK.V1.MergeAndDeliverViaWebhookSuccessResponse) {
                        console.log("\nRecords - " + JSON.stringify(writerResponseObject.getRecords()));
                        console.log("\nMerge report data url - " + writerResponseObject.getMergeReportDataUrl());
                    } else if (writerResponseObject instanceof SDK.V1.InvalidConfigurationException) {
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

    //Include office-integrator-sdk package in your package json and the execute this code.

    static async initializeSdk() {

        // Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
        let environment = await new SDK.ApiServer.Production("https://api.office-integrator.com");

        let auth = new SDK.AuthBuilder()
                        .addParam("apikey", "2ae438cf864488657cc9754a27daa480") //Update this apikey with your own apikey signed up in office inetgrator service
                        .authenticationSchema(await new SDK.V1.Authentication().getTokenFlow())
                        .build();

        let tokens = [ auth ];

        //Sdk application log configuration
        let logger = new SDK.LogBuilder()
            .level(SDK.Levels.INFO)
            //.filePath("<file absolute path where logs would be written>") //No I18N
            .build();

        let initialize = await new SDK.InitializeBuilder();

        await initialize.environment(environment).tokens(tokens).logger(logger).initialize();

        console.log("SDK initialized successfully.");
    }
}

MergeAndDeliver.execute();
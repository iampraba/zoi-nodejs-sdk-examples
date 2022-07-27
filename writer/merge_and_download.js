const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const MergeAndDownloadDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/merge_and_download_document_parameters").MergeAndDownloadDocumentParameters;

const FileBodyWrapper = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/file_body_wrapper").FileBodyWrapper;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class MergeAndDownload {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var parameters = new MergeAndDownloadDocumentParameters();

            //parameters.setFileUrl("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");

            var fileName = "OfferLetter.zdoc";
            var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            parameters.setPassword("***");
            parameters.setOutputFormat("pdf");
            parameters.setFileContent(streamWrapper);

            var jsonFileName = "candidates.json";
            var jsonFilePath = __dirname + "/sample_documents/candidates.json";
            var jsonFileStream = fs.readFileSync(jsonFilePath);
            var jsonStreamWrapper = new StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            parameters.setMergeDataJsonContent(jsonStreamWrapper);

            /*
            var mergeData = new Map();

            parameters.setMergeData(mergeData);

            var csvFileName = "csv_data_source.csv";
            var csvFilePath = __dirname + "/sample_documents/csv_data_source.csv";
            var csvFileStream = fs.readFileSync(csvFilePath);
            var csvStreamWrapper = new StreamWrapper(csvFileName, csvFileStream, csvFilePath);

            parameters.setMergeDataCsvContent(csvStreamWrapper);

            parameters.setMergeDataCsvUrl("https://demo.office-integrator.com/data/csv_data_source.csv");
            parameters.setMergeDataJsonUrl("https://demo.office-integrator.com/zdocs/json_data_source.json");
            */

            var responseObject = await sdkOperations.mergeAndDownloadDocument(parameters);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                //TODO: Need to fix the file writing to client
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof FileBodyWrapper) {
                        var convertedDocument = writerResponseObject.getFile();

                        if (convertedDocument instanceof StreamWrapper) {
                            var outputFilePath = __dirname + "/sample_documents/merge_and_download.pdf";

                            fs.writeFileSync(outputFilePath, convertedDocument.getStream());
                            console.log("Check merged output file in file path - ", outputFilePath);
                        }
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

MergeAndDownload.execute();
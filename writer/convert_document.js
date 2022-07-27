const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;
const DocumentConversionParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_conversion_parameters").DocumentConversionParameters;
const DocumentConversionOutputOptions = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_conversion_output_options").DocumentConversionOutputOptions;
const FileBodyWrapper = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/file_body_wrapper").FileBodyWrapper;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class ConvertDocument {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var documentConversionParameters = new DocumentConversionParameters();

            //Either use url as document source or attach the document in request body use below methods
            //documentConversionParameters.setUrl("https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx");

            var fileName = "Graphic-Design-Proposal.docx";
            var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            //var streamWrapper = new StreamWrapper(null, null, filePath)

            documentConversionParameters.setDocument(streamWrapper);

            var outputOptions = new DocumentConversionOutputOptions();

            outputOptions.setFormat("pdf");
            outputOptions.setDocumentName("conversion_output.pdf");
            outputOptions.setIncludeComments("all");
            outputOptions.setIncludeChanges("all");

            documentConversionParameters.setOutputOptions(outputOptions);
            documentConversionParameters.setPassword("***");
            var responseObject = await sdkOperations.convertDocument(documentConversionParameters);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;

                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof FileBodyWrapper) {
                        var convertedDocument = writerResponseObject.getFile();

                        if (convertedDocument instanceof StreamWrapper) {
                            var outputFilePath = __dirname + "/sample_documents/conversion_output.pdf";

                            fs.writeFileSync(outputFilePath, convertedDocument.getStream());
                            console.log("Check converted output file in file path - ", outputFilePath);
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

ConvertDocument.execute();
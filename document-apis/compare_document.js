const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;
const CompareDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/compare_document_parameters").CompareDocumentParameters;
const CompareDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/compare_document_response").CompareDocumentResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class CompareDocument {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var compareDocumentParameters = new CompareDocumentParameters();

            // compareDocumentParameters.setUrl1("https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx");
            // compareDocumentParameters.setUrl2("https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx");

            var file1Name = "MS_Word_Document_v0.docx";
            var file1Path = __dirname + "/sample_documents/MS_Word_Document_v0.docx";
            var file1Stream = fs.readFileSync(file1Path);
            var stream1Wrapper = new StreamWrapper(file1Name, file1Stream, file1Path)
            //var stream1Wrapper = new StreamWrapper(null, null, file1Path)

            var file2Name = "MS_Word_Document_v1.docx";
            var file2Path = __dirname + "/sample_documents/MS_Word_Document_v1.docx";
            var file2Stream = fs.readFileSync(file2Path);
            var stream2Wrapper = new StreamWrapper(file2Name, file2Stream, file2Path)
            //var stream2Wrapper = new StreamWrapper(null, null, file2Path)
            
            compareDocumentParameters.setDocument1(stream1Wrapper);
            compareDocumentParameters.setDocument2(stream2Wrapper);

            compareDocumentParameters.setLang("en");
            compareDocumentParameters.setTitle(file1Name + " vs " + file2Name);

            var responseObject = await sdkOperations.compareDocument(compareDocumentParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    //Check if expected CompareDocumentResponse instance is received
                    if(writerResponseObject instanceof CompareDocumentResponse) {
                        console.log("Compare URL - " + writerResponseObject.getCompareUrl());
                        console.log("Document session delete URL - " + writerResponseObject.getSessionDeleteUrl());
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

CompareDocument.execute();
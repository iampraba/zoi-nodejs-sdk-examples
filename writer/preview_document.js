const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const PreviewParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/preview_parameters").PreviewParameters;

const PreviewDocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/preview_document_info").PreviewDocumentInfo;

const PreviewResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/preview_response").PreviewResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class PreviewDocument {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var previewParameters = new PreviewParameters();

            //previewParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            var fileName = "Graphic-Design-Proposal.docx";
            var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath)
            //var streamWrapper = new StreamWrapper(null, null, filePath)
            
            previewParameters.setDocument(streamWrapper);

            var previewDocumentInfo = new PreviewDocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            previewDocumentInfo.setDocumentName("Graphic-Design-Proposal.docx");

            previewParameters.setDocumentInfo(previewDocumentInfo);

            var permissions = new Map();

            permissions.set("document.print", false);

            previewParameters.setPermissions(permissions);

            var responseObject = await sdkOperations.createDocumentPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    console.log("Preview URL : " + writerResponseObject.getPreviewUrl());
                    //SDK ISSUE: Need to fix below check since instanceof check not working for proper response
                    //Check if expected PreviewResponse instance is received
                    if(writerResponseObject instanceof PreviewResponse) {
                        console.log("Document ID - " + writerResponseObject.getDocumentId());
                        console.log("Document session ID - " + writerResponseObject.getSessionId());
                        console.log("Document preview URL - " + writerResponseObject.getPreviewUrl());
                        console.log("Document delete URL - " + writerResponseObject.getDocumentDeleteUrl());
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

PreviewDocument.execute();
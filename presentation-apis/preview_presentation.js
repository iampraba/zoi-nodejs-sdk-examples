const fs = require("fs");
const { DocumentInfo } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_info");
const { PresentationPreviewParameters } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/presentation_preview_parameters");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const PreviewResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/preview_response").PreviewResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class PreviewPresentation {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var previewParameters = new PresentationPreviewParameters();
            
            //TODO: Need to fix url import
            previewParameters.setUrl("https://demo.office-integrator.com/samples/show/Zoho_Show.pptx");
            
            var fileName = "Zoho_Show.pptx";
            var filePath = __dirname + "/sample_documents/Zoho_Show.pptx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            //previewParameters.setDocument(streamWrapper);

            var documentInfo = new DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Zoho_Show.pptx");

            previewParameters.setDocumentInfo(documentInfo);

            //TODO: Language params issue
            //previewParameters.setLanguage("en");

            var responseObject = await sdkOperations.createPresentationPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){

                    console.log("Preview URL : " + writerResponseObject.getPreviewUrl());

                    if(writerResponseObject instanceof PreviewResponse){
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

PreviewPresentation.execute();
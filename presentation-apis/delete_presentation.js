const { CreatePresentationParameters } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_presentation_parameters");

const SDKInitializer = require("../SDKInitializer").SDKInitializer;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const DocumentDeleteSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_delete_success_response").DocumentDeleteSuccessResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class DeletePresentation {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createPresentationParameters = new CreatePresentationParameters();

            var newDocumentCreateResponse = await sdkOperations.createPresentation(createPresentationParameters);
            var documentId = newDocumentCreateResponse.object.getDocumentId();

            console.log("Presentation id to be deleted - ", documentId);

            var responseObject = await sdkOperations.deleteDocument(documentId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    if(writerResponseObject instanceof DocumentDeleteSuccessResponse){
                        console.log("Document delete status - " + writerResponseObject.getDocDelete());
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

DeletePresentation.execute();
const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_parameters").CreateDocumentParameters;

const SDKInitializer = require("../SDKInitializer").SDKInitializer;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const SessionDeleteSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/session_delete_success_response").SessionDeleteSuccessResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class DeleteDocumentSession {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var createDocumentParameters = new CreateDocumentParameters();
            var sdkOperations = new OfficeIntegratorSDKOperations();

            var createResponse = await sdkOperations.createDocument(createDocumentParameters);

            //TODO: Invaid document id
            //var sessionId = BigInt(1000);
            var sessionId = createResponse.object.getSessionId();

            console.log("sessionId: " + sessionId);

            var responseObject = await sdkOperations.deleteSession(sessionId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    if(writerResponseObject instanceof SessionDeleteSuccessResponse){
                        console.log("Session delete status - " + writerResponseObject.getSessionDelete());
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

DeleteDocumentSession.execute();
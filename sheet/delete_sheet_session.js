const { CreateSheetParameters } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_sheet_parameters");

const SDKInitializer = require("../SDKInitializer").SDKInitializer;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const SessionDeleteSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/session_delete_success_response").SessionDeleteSuccessResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class DeleteSheetSession {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createSheetParameters = new CreateSheetParameters();

            var newSessionObject = await sdkOperations.createSheet(createSheetParameters);
            var sessionId = newSessionObject.object.getSessionId();

            console.log("Sheet session id to be deleted - ", sessionId);

            var responseObject = await sdkOperations.deleteSheetSession(sessionId);

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

DeleteSheetSession.execute();
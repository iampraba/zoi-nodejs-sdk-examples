const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const CreatePresentationParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_presentation_parameters").CreatePresentationParameters;
const SessionDeleteSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/session_delete_success_response").SessionDeleteSuccessResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class DeletePresentationSession {

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
            var sdkOperations = new V1Operations();
            var createPresentationParameters = new CreatePresentationParameters();
            var newDocumentCreateResponse = await sdkOperations.createPresentation(createPresentationParameters);
            var sessionId = newDocumentCreateResponse.object.getSessionId();

            console.log("\nCreated a new presentation session to demonstrate the document delete api. Created session ID - " + sessionId);

            var responseObject = await sdkOperations.deletePresentationSession(sessionId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    if(writerResponseObject instanceof SessionDeleteSuccessResponse){
                        console.log("\nSession delete status - " + writerResponseObject.getSessionDelete());
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

DeletePresentationSession.execute();
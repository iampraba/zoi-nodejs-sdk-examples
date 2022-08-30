const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const FileDeleteSuccessResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/file_delete_success_response").FileDeleteSuccessResponse;
const CreatePresentationParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/create_presentation_parameters").CreatePresentationParameters;
const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class DeletePresentation {

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
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createPresentationParameters = new CreatePresentationParameters();

            var newPresentationCreateResponse = await sdkOperations.createPresentation(createPresentationParameters);
            var presentationId = newPresentationCreateResponse.object.getDocumentId();

            console.log("\nPresentation id to be deleted - ", presentationId);

            var responseObject = await sdkOperations.deletePresentation(presentationId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let presentationDeleteResponseObject = responseObject.object;
    
                if(presentationDeleteResponseObject != null){
                    if(presentationDeleteResponseObject instanceof FileDeleteSuccessResponse){
                        console.log("\nDocument delete status - " + presentationDeleteResponseObject.getDocDelete());
                    } else if (presentationDeleteResponseObject instanceof InvaildConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", presentationDeleteResponseObject);
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

DeletePresentation.execute();
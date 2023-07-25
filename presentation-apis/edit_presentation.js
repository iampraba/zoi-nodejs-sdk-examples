const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const UserInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/user_info").UserInfo;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_info").DocumentInfo;
const CallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/callback_settings").CallbackSettings;
const ZohoShowEditorSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/zoho_show_editor_settings").ZohoShowEditorSettings;
const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_response").CreateDocumentResponse;
const CreatePresentationParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_presentation_parameters").CreatePresentationParameters;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;

class EditPresentation {

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

            createPresentationParameters.setUrl("https://demo.office-integrator.com/samples/show/Zoho_Show.pptx");
            
            // var fileName = "Zoho_Show.pptx";
            // var filePath = __dirname + "/sample_documents/Zoho_Show.pptx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            // var streamWrapper = new StreamWrapper(null, null, filePath);
            
            // createPresentationParameters.setDocument(streamWrapper);
            
            var documentInfo = new DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Zoho_Show.pptx");

            createPresentationParameters.setDocumentInfo(documentInfo);

            var userInfo = new UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            createPresentationParameters.setUserInfo(userInfo);

            var editorSettings = new ZohoShowEditorSettings();

            editorSettings.setLanguage("en");

            createPresentationParameters.setEditorSettings(editorSettings);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);

            createPresentationParameters.setPermissions(permissions);

            var callbackSettings = new CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            /*var saveUrlHeaders = new Map();

            saveUrlHeaders.set("header1", "value1");
            saveUrlHeaders.set("header2", "value2");

            callbackSettings.setSaveUrlHeaders(saveUrlHeaders);*/

            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveFormat("pptx");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createPresentationParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createPresentation(createPresentationParameters);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                let presentationResponseObject = responseObject.object;
    
                if(presentationResponseObject != null){

                    if(presentationResponseObject instanceof CreateDocumentResponse){
                        console.log("\nPresentation ID - " + presentationResponseObject.getDocumentId());
                        console.log("\nPresentation session ID - " + presentationResponseObject.getSessionId());
                        console.log("\nPresentation session URL - " + presentationResponseObject.getDocumentUrl());
                        console.log("\nPresentation save URL - " + presentationResponseObject.getSaveUrl());
                        console.log("\nPresentation delete URL - " + presentationResponseObject.getDocumentDeleteUrl());
                        console.log("\nPresentation session delete URL - " + presentationResponseObject.getSessionDeleteUrl());
                    } else if (presentationResponseObject instanceof InvaildConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", presentationResponseObject);
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

EditPresentation.execute();
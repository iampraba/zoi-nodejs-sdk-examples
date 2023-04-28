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
const ShowCallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/show_callback_settings").ShowCallbackSettings;
const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_response").CreateDocumentResponse;
const ZohoShowEditorSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/zoho_show_editor_settings").ZohoShowEditorSettings;
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
            
            var documentInfo = new DocumentInfo();

            //To collaborate in existing document you need to provide the document id(e.g: 1000) alone is enough.
            //Note: Make sure the document already exist in Zoho server for below given document id.
            //Even if the document is added to this request, if document exist in zoho server for given document id,
            //then session will be create for document already exist with Zoho.
            documentInfo.setDocumentId("1000");

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

            var callbackSettings = new ShowCallbackSettings();

            callbackSettings.setSaveFormat("pptx");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createPresentationParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createPresentation(createPresentationParameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){

                    if(writerResponseObject instanceof CreateDocumentResponse){
                        console.log("\nPresentation ID - " + writerResponseObject.getDocumentId());
                        console.log("\nPresentation session ID - " + writerResponseObject.getSessionId());
                        console.log("\nPresentation session URL - " + writerResponseObject.getDocumentUrl());
                        console.log("\nPresentation save URL - " + writerResponseObject.getSaveUrl());
                        console.log("\nPresentation delete URL - " + writerResponseObject.getDocumentDeleteUrl());
                        console.log("\nPresentation session delete URL - " + writerResponseObject.getSessionDeleteUrl());
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

EditPresentation.execute();
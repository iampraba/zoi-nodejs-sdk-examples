const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const UiOptions = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/ui_options").UiOptions;
const UserInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/user_info").UserInfo;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/document_info").DocumentInfo;
const DocumentDefaults = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/document_defaults").DocumentDefaults;
const CallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/callback_settings").CallbackSettings;
const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/create_document_response").CreateDocumentResponse;
const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/create_document_parameters").CreateDocumentParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;
const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

class CoEditDocument {

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
            var createDocumentParameters = new CreateDocumentParameters();

            var documentInfo = new DocumentInfo();

            //To collaborate in existing document you need to provide the document id(e.g: 1000) alone is enough.
            //Note: Make sure the document already exist in Zoho server for below given document id.
            //Even if the document is added to this request, if document exist in zoho server for given document id,
            //then session will be create for document already exist with Zoho.
            documentInfo.setDocumentId("1000");

            createDocumentParameters.setDocumentInfo(documentInfo);

            var userInfo = new UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            createDocumentParameters.setUserInfo(userInfo);

            var documentDefaults = new DocumentDefaults();

            documentDefaults.setTrackChanges("enabled");

            createDocumentParameters.setDocumentDefaults(documentDefaults);

            var uiOptions = new UiOptions();

            uiOptions.setDarkMode("show");
            uiOptions.setFileMenu("show");
            uiOptions.setSaveButton("show");
            uiOptions.setChatPanel("show");
            createDocumentParameters.setUiOptions(uiOptions);

            var permissions = new Map();

            permissions.set("collab.chat", false);
            permissions.set("document.edit", true);
            permissions.set("document.fill", false);
            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("review.comment", false);
            permissions.set("review.changes.resolve", false);
            permissions.set("document.pausecollaboration", false);

            createDocumentParameters.setPermissions(permissions);

            var callbackSettings = new CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            callbackSettings.setSaveUrlParams(saveUrlParams);            
            callbackSettings.setRetries(1);
            callbackSettings.setSaveFormat("docx");
            callbackSettings.setHttpMethodType("post");
            callbackSettings.setTimeout(100000);
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createDocumentParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createDocument(createDocumentParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
    
                    //Check if expected CreateDocumentResponse instance is received
                    if(writerResponseObject instanceof CreateDocumentResponse){
                        console.log("\nDocument ID - " + writerResponseObject.getDocumentId());
                        console.log("\nDocument session ID - " + writerResponseObject.getSessionId());
                        console.log("\nDocument session URL - " + writerResponseObject.getDocumentUrl());
                        console.log("\nDocument save URL - " + writerResponseObject.getSaveUrl());
                        console.log("\nDocument delete URL - " + writerResponseObject.getDocumentDeleteUrl());
                        console.log("\nDocument session delete URL - " + writerResponseObject.getSessionDeleteUrl());
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

CoEditDocument.execute();
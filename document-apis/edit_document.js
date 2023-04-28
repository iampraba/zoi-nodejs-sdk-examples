const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const Margin = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/margin").Margin;
const UserInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/user_info").UserInfo;
const UiOptions = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/ui_options").UiOptions;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_info").DocumentInfo;
const EditorSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/editor_settings").EditorSettings;
const DocumentDefaults = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_defaults").DocumentDefaults;
const CallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/callback_settings").CallbackSettings;
const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_response").CreateDocumentResponse;
const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_parameters").CreateDocumentParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class EditDocument {

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
            var createDocumentParameters = new CreateDocumentParameters();

            createDocumentParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            //var fileName = "Graphic-Design-Proposal.docx";
            //var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            //var fileStream = fs.readFileSync(filePath);
            //var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            //createDocumentParameters.setDocument(streamWrapper);

            var documentInfo = new DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Graphic-Design-Proposal.docx");

            createDocumentParameters.setDocumentInfo(documentInfo);

            var userInfo = new UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            createDocumentParameters.setUserInfo(userInfo);

            var margin = new Margin();

            margin.setTop("2in");
            margin.setBottom("2in");
            margin.setLeft("2in");
            margin.setRight("2in");

            var documentDefaults = new DocumentDefaults();

            documentDefaults.setLanguage("ta");
            documentDefaults.setTrackChanges("enabled");

            createDocumentParameters.setDocumentDefaults(documentDefaults);

            var editorSettings = new EditorSettings();

            editorSettings.setUnit("mm");
            editorSettings.setLanguage("en");
            editorSettings.setView("pageview");

            createDocumentParameters.setEditorSettings(editorSettings);

            var uiOptions = new UiOptions();

            uiOptions.setDarkMode("show");
            uiOptions.setFileMenu("show");
            uiOptions.setSaveButton("show");
            uiOptions.setChatPanel("show");

            createDocumentParameters.setUiOptions(uiOptions);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);
            permissions.set("review.comment", false);
            permissions.set("review.changes.resolve", false);
            permissions.set("collab.chat", false);
            permissions.set("document.pausecollaboration", false);
            permissions.set("document.fill", false);

            createDocumentParameters.setPermissions(permissions);

            var callbackSettings = new CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            callbackSettings.setSaveUrlParams(saveUrlParams);

            callbackSettings.setRetries(1);
            callbackSettings.setTimeout(10000);
            callbackSettings.setSaveFormat("zdoc");
            callbackSettings.setHttpMethodType("post");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createDocumentParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createDocument(createDocumentParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
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

EditDocument.execute();
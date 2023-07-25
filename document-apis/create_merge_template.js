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
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_info").DocumentInfo;
const EditorSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/editor_settings").EditorSettings;
const DocumentDefaults = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_defaults").DocumentDefaults;
const CallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/callback_settings").CallbackSettings;
const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_response").CreateDocumentResponse;
const MailMergeTemplateParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/mail_merge_template_parameters").MailMergeTemplateParameters;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;

class CreateMergeTemplate {

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

        console.log("SDK initialized successfully.");
    }

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new V1Operations();
            var templateParameters = new MailMergeTemplateParameters();
            
            //Either use url as document source or attach the document in request body use below methods
            templateParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");
            templateParameters.setMergeDataJsonUrl("https://demo.office-integrator.com/data/candidates.json");

            // var fileName = "OfferLetter.zdoc";
            // var filePath = "./sample_documents/OfferLetter.zdoc";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            // var streamWrapper = new StreamWrapper(null, null, filePath)
            
            // templateParameters.setDocument(streamWrapper);

            // var jsonFileName = "candidates.json";
            // var jsonFilePath = "./sample_documents/candidates.json";
            // var jsonFileStream = fs.readFileSync(jsonFilePath);
            // var jsonStreamWrapper = new StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            // templateParameters.setMergeDataJsonContent(jsonStreamWrapper);

            var documentInfo = new DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Graphic-Design-Proposal.docx");

            templateParameters.setDocumentInfo(documentInfo);

            var userInfo = new UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            templateParameters.setUserInfo(userInfo);

            var margin = new Margin();

            margin.setTop("2in");
            margin.setBottom("2in");
            margin.setLeft("2in");
            margin.setRight("2in");

            var documentDefaults = new DocumentDefaults();

            documentDefaults.setFontName("Arial");
            documentDefaults.setFontSize(12);
            documentDefaults.setOrientation("landscape");
            documentDefaults.setPaperSize("A4");
            documentDefaults.setTrackChanges("enabled");
            documentDefaults.setMargin(margin);

            templateParameters.setDocumentDefaults(documentDefaults);

            var editorSettings = new EditorSettings();

            editorSettings.setUnit("mm");
            editorSettings.setLanguage("en");
            editorSettings.setView("pageview");

            templateParameters.setEditorSettings(editorSettings);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);
            permissions.set("review.comment", false);
            permissions.set("review.changes.resolve", false);
            permissions.set("collab.chat", false);
            permissions.set("document.pausecollaboration", false);
            permissions.set("document.fill", false);

            templateParameters.setPermissions(permissions);

            var callbackSettings = new CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            var saveUrlHeaders = new Map();

            saveUrlHeaders.set("header1", "value1");
            saveUrlHeaders.set("header2", "value2");

            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveUrlHeaders(saveUrlHeaders);

            callbackSettings.setHttpMethodType("post");
            callbackSettings.setRetries(1);
            callbackSettings.setTimeout(100000);
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");
            callbackSettings.setSaveFormat("pdf");

            templateParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createMailMergeTemplate(templateParameters);

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

CreateMergeTemplate.execute();
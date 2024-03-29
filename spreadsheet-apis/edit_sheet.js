const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_info").DocumentInfo;
const SheetUserSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/sheet_user_settings").SheetUserSettings;
const SheetEditorSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/sheet_editor_settings").SheetEditorSettings;
const CreateSheetResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_sheet_response").CreateSheetResponse;
const CreateSheetParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_sheet_parameters").CreateSheetParameters;
const SheetCallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/sheet_callback_settings").SheetCallbackSettings;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;

class EditSheet {

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
            var createSheetParameters = new CreateSheetParameters();

            createSheetParameters.setUrl("https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx");
            
            // var fileName = "Contact_List.xlsx";
            // var filePath = __dirname + "/sample_documents/Contact_List.xlsx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            // createSheetParameters.setDocument(streamWrapper);

            var documentInfo = new DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Contact_List.xlsx");

            createSheetParameters.setDocumentInfo(documentInfo);

            var userInfo = new SheetUserSettings();

            userInfo.setDisplayName("Prabakaran R");

            createSheetParameters.setUserInfo(userInfo);

            var editorSettings = new SheetEditorSettings();

            editorSettings.setCountry("IN");
            editorSettings.setLanguage("en");

            createSheetParameters.setEditorSettings(editorSettings);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);

            createSheetParameters.setPermissions(permissions);

            var callbackSettings = new SheetCallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            callbackSettings.setSaveFormat("xlsx");
            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createSheetParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createSheet(createSheetParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);

                //Get the api response object from responseObject
                let sheetResponseObject = responseObject.object;

                if(sheetResponseObject != null){

                    //Check if expected CreateSheetResponse instance is received
                    if(sheetResponseObject instanceof CreateSheetResponse){
                        console.log("\nSpreadSheet ID - " + sheetResponseObject.getDocumentId());
                        console.log("\nSpreadSheet session ID - " + sheetResponseObject.getSessionId());
                        console.log("\nSpreadSheet session URL - " + sheetResponseObject.getDocumentUrl());
                        console.log("\nSpreadSheet Grid View URL - " + sheetResponseObject.getGridviewUrl());
                        console.log("\nSpreadSheet save URL - " + sheetResponseObject.getSaveUrl());
                        console.log("\nSpreadSheet delete URL - " + sheetResponseObject.getDocumentDeleteUrl());
                        console.log("\nSpreadSheet session delete URL - " + sheetResponseObject.getSessionDeleteUrl());
                    } else if (sheetResponseObject instanceof InvaildConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", sheetResponseObject);
                    } else {
                        console.log("\nRequest not completed successfullly");
                    }
                }
            }
        } catch (error) {
            console.log("Exception while running sample code", error);
        }
    }
}

EditSheet.execute();
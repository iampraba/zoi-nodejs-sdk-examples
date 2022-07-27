const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const CreateSheetParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_sheet_parameters").CreateSheetParameters;
const SheetUserSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/sheet_user_settings").SheetUserSettings;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_info").DocumentInfo;
const SheetEditorSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/sheet_editor_settings").SheetEditorSettings;
const SheetCallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/sheet_callback_settings").SheetCallbackSettings;

const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_response").CreateDocumentResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class EditSheet {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createSheetParameters = new CreateSheetParameters();

            //createSheetParameters.setUrl("https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx");
            
            var fileName = "Contact_List.xlsx";
            var filePath = __dirname + "/sample_documents/Contact_List.xlsx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            createSheetParameters.setDocument(streamWrapper);

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
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
    
                    //Check if expected CreateDocumentResponse instance is received
                    if(writerResponseObject instanceof CreateDocumentResponse){
                        console.log("Document ID - " + writerResponseObject.getDocumentId());
                        console.log("Document save URL - " + writerResponseObject.getSaveUrl());
                        console.log("Document delete URL - " + writerResponseObject.getDocumentDeleteUrl());
                        console.log("Document session ID - " + writerResponseObject.getSessionId());
                        console.log("Document session URL - " + writerResponseObject.getDocumentUrl());
                        console.log("Document session delete URL - " + writerResponseObject.getSessionDeleteUrl());
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

EditSheet.execute();
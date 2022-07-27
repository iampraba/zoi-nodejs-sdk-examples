const { CreatePresentationParameters } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_presentation_parameters");
const { ShowCallbackSettings } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/show_callback_settings");
const { UserInfo } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/user_info");
const { ZohoShowEditorSettings } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/zoho_show_editor_settings");
const { DocumentInfo } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_info");

const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;



const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_response").CreateDocumentResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class EditPresentation {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createPresentationParameters = new CreatePresentationParameters();

            //TODO: Need to fix url import
            //createPresentationParameters.setUrl("https://demo.office-integrator.com/samples/show/Zoho_Show.pptx");
            
            var fileName = "Zoho_Show.pptx";
            var filePath = __dirname + "/sample_documents/Zoho_Show.pptx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            //var streamWrapper = new StreamWrapper(null, null, filePath);
            
            createPresentationParameters.setDocument(streamWrapper);
            
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

            var callbackSettings = new ShowCallbackSettings();

            callbackSettings.setSaveFormat("pptx");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createPresentationParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createPresentation(createPresentationParameters);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){

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

EditPresentation.execute();
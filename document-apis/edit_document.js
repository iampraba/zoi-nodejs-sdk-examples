const fs = require("fs");
const { UiOptions } = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/ui_options");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_parameters").CreateDocumentParameters;
const Margin = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/margin").Margin;
const UserInfo = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/user_info").UserInfo;
const DocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_info").DocumentInfo;
const EditorSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/editor_settings").EditorSettings;
const DocumentDefaults = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/document_defaults").DocumentDefaults;
const CallbackSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/callback_settings").CallbackSettings;

const CreateDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_response").CreateDocumentResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class EditDocument {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var createDocumentParameters = new CreateDocumentParameters();

            // createDocumentParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            var fileName = "Graphic-Design-Proposal.docx";
            var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            createDocumentParameters.setDocument(streamWrapper);

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

            //TODO: Need to add language option in this code
            documentDefaults.setTrackChanges("enabled");

            createDocumentParameters.setDocumentDefaults(documentDefaults);

            var editorSettings = new EditorSettings();

            editorSettings.setUnit("mm");
            editorSettings.setLanguage("en");
            editorSettings.setView("pageview");

            createDocumentParameters.setEditorSettings(editorSettings);

            var uiOptions = new UiOptions();

            uiOptions.setDarkMode("hide");
            uiOptions.setFileMenu("hide");
            uiOptions.setSaveButton("hide");
            uiOptions.setChatPanel("hide");
            //TODO: UIOptions object undefined exception
            //createDocumentParameters.setUiOptions(uiOptions);

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

            callbackSettings.setHttpMethodType("post");
            //callbackSettings.setRetries(BigInt(1));
            //callbackSettings.setTimeout(BigInt(100000));
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");
            callbackSettings.setSaveFormat("pdf");

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

EditDocument.execute();
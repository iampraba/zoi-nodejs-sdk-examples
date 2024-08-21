import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync } from 'fs';
const __dirname = import.meta.dirname;

class CreateMergeTemplate {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var templateParameters = new SDK.V1.MailMergeTemplateParameters();
            
            //Either use url as document source or attach the document in request body use below methods
            templateParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");
            templateParameters.setMergeDataJsonUrl("https://demo.office-integrator.com/data/candidates.json");

            // var fileName = "OfferLetter.zdoc";
            // var filePath = "./sample_documents/OfferLetter.zdoc";
            // var fileStream = readFileSync(filePath);
            // var streamWrapper = new SDK.StreamWrapper(fileName, fileStream, filePath);
            
            // templateParameters.setDocument(streamWrapper);

            // var jsonFileName = "candidates.json";
            // var jsonFilePath = "./sample_documents/candidates.json";
            // var jsonFileStream = readFileSync(jsonFilePath);
            // var jsonStreamWrapper = new SDK.StreamWrapper(jsonFileName, jsonFileStream, jsonFilePath);

            // templateParameters.setMergeDataJsonContent(jsonStreamWrapper);

            var documentInfo = new SDK.V1.DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Graphic-Design-Proposal.docx");

            templateParameters.setDocumentInfo(documentInfo);

            var userInfo = new SDK.V1.UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            templateParameters.setUserInfo(userInfo);

            var margin = new SDK.V1.Margin();

            margin.setTop("2in");
            margin.setBottom("2in");
            margin.setLeft("2in");
            margin.setRight("2in");

            var documentDefaults = new SDK.V1.DocumentDefaults();

            documentDefaults.setFontName("Arial");
            documentDefaults.setFontSize(12);
            documentDefaults.setOrientation("landscape");
            documentDefaults.setPaperSize("A4");
            documentDefaults.setTrackChanges("enabled");
            documentDefaults.setMargin(margin);

            templateParameters.setDocumentDefaults(documentDefaults);

            var editorSettings = new SDK.V1.EditorSettings();

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

            var callbackSettings = new SDK.V1.CallbackSettings();
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

            var uiOptions = new SDK.V1.UiOptions();

            uiOptions.setDarkMode("show");
            uiOptions.setSaveButton("show");
            uiOptions.setChatPanel("show");
            uiOptions.setFileMenu("show");

            templateParameters.setUiOptions(uiOptions);

            var responseObject = await sdkOperations.createMailMergeTemplate(templateParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
    
                    //Check if expected CreateDocumentResponse instance is received
                    if(writerResponseObject instanceof SDK.V1.CreateDocumentResponse){
                        console.log("\nDocument ID - " + writerResponseObject.getDocumentId());
                        console.log("\nDocument session ID - " + writerResponseObject.getSessionId());
                        console.log("\nDocument session URL - " + writerResponseObject.getDocumentUrl());
                        console.log("\nDocument save URL - " + writerResponseObject.getSaveUrl());
                        console.log("\nDocument delete URL - " + writerResponseObject.getDocumentDeleteUrl());
                        console.log("\nDocument session delete URL - " + writerResponseObject.getSessionDeleteUrl());
                    } else if (writerResponseObject instanceof SDK.V1.InvalidConfigurationException) {
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

    //Include office-integrator-sdk package in your package json and the execute this code.

    static async initializeSdk() {

        // Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
        let environment = await new SDK.ApiServer.Production("https://api.office-integrator.com");

        let auth = new SDK.AuthBuilder()
                        .addParam("apikey", "2ae438cf864488657cc9754a27daa480") //Update this apikey with your own apikey signed up in office inetgrator service
                        .authenticationSchema(await new SDK.V1.Authentication().getTokenFlow())
                        .build();

        let tokens = [ auth ];

        //Sdk application log configuration
        let logger = new SDK.LogBuilder()
            .level(SDK.Levels.INFO)
            //.filePath("<file absolute path where logs would be written>") //No I18N
            .build();

        let initialize = await new SDK.InitializeBuilder();

        await initialize.environment(environment).tokens(tokens).logger(logger).initialize();

        console.log("SDK initialized successfully.");
    }
}

CreateMergeTemplate.execute();
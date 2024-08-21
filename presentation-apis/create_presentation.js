import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync, writeFileSync } from 'fs';
const __dirname = import.meta.dirname;

class CreatePresentation {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var createPresentationParameters = new SDK.V1.CreatePresentationParameters();
            
            var documentInfo = new SDK.V1.DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("New Presentation");

            createPresentationParameters.setDocumentInfo(documentInfo);

            var userInfo = new SDK.V1.UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            createPresentationParameters.setUserInfo(userInfo);

            var editorSettings = new SDK.V1.ZohoShowEditorSettings();

            editorSettings.setLanguage("en");

            createPresentationParameters.setEditorSettings(editorSettings);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);

            createPresentationParameters.setPermissions(permissions);

            var callbackSettings = new SDK.V1.CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            /* var saveUrlHeaders = new Map();

            saveUrlHeaders.set("header1", "value1");
            saveUrlHeaders.set("header2", "value2");

            callbackSettings.setSaveUrlHeaders(saveUrlHeaders); */

            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveFormat("pptx");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            createPresentationParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createPresentation(createPresentationParameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){

                    if(writerResponseObject instanceof SDK.V1.CreateDocumentResponse){
                        console.log("\nPresentation ID - " + writerResponseObject.getDocumentId());
                        console.log("\nPresentation session ID - " + writerResponseObject.getSessionId());
                        console.log("\nPresentation session URL - " + writerResponseObject.getDocumentUrl());
                        console.log("\nPresentation save URL - " + writerResponseObject.getSaveUrl());
                        console.log("\nPresentation delete URL - " + writerResponseObject.getDocumentDeleteUrl());
                        console.log("\nPresentation session delete URL - " + writerResponseObject.getSessionDeleteUrl());
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

CreatePresentation.execute();
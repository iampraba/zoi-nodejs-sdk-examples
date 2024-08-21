import * as SDK from "@zoho-corp/office-integrator-sdk";

class CreateDocument {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var createDocumentParameters = new SDK.V1.CreateDocumentParameters();

            var documentInfo = new SDK.V1.DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("New Document");

            createDocumentParameters.setDocumentInfo(documentInfo);

            var userInfo = new SDK.V1.UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            createDocumentParameters.setUserInfo(userInfo);

            var margin = new SDK.V1.Margin();

            margin.setTop("2in");
            margin.setBottom("2in");
            margin.setLeft("2in");
            margin.setRight("2in");

            var documentDefaults = new SDK.V1.DocumentDefaults();

            documentDefaults.setFontSize(12);
            documentDefaults.setPaperSize("A4");
            documentDefaults.setFontName("Arial");
            documentDefaults.setTrackChanges("enabled");
            documentDefaults.setOrientation("landscape");

            documentDefaults.setMargin(margin);
            documentDefaults.setLanguage("ta");

            createDocumentParameters.setDocumentDefaults(documentDefaults);

            var editorSettings = new SDK.V1.EditorSettings();

            editorSettings.setUnit("in");
            editorSettings.setLanguage("en");
            editorSettings.setView("pageview");

            createDocumentParameters.setEditorSettings(editorSettings);

            var uiOptions = new SDK.V1.UiOptions();

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

            var callbackSettings = new SDK.V1.CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            var saveUrlHeaders = new Map();

            saveUrlHeaders.set("header1", "value1");
            saveUrlHeaders.set("header2", "value2");

            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveUrlHeaders(saveUrlHeaders);
            callbackSettings.setRetries(1);
            callbackSettings.setSaveFormat("zdoc");
            callbackSettings.setHttpMethodType("post");
            callbackSettings.setTimeout(100000);
            callbackSettings.setSaveUrl("https://bc59556ab64f536be787aeca1dd36571.m.pipedream.net");

            createDocumentParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.createDocument(createDocumentParameters);

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

CreateDocument.execute();
import * as SDK from "@zoho-corp/office-integrator-sdk";

class EditSheet {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var createSheetParameters = new SDK.V1.CreateSheetParameters();

            var documentInfo = new SDK.V1.DocumentInfo();

            //To collaborate in existing document you need to provide the document id(e.g: 1000) alone is enough.
            //Note: Make sure the document already exist in Zoho server for below given document id.
            //Even if the document is added to this request, if document exist in zoho server for given document id,
            //then session will be create for document already exist with Zoho.
            documentInfo.setDocumentId("1000");
            documentInfo.setDocumentName("Contact_List.xlsx");

            createSheetParameters.setDocumentInfo(documentInfo);

            var userInfo = new SDK.V1.SheetUserSettings();

            userInfo.setDisplayName("Prabakaran R");

            createSheetParameters.setUserInfo(userInfo);

            var editorSettings = new SDK.V1.SheetEditorSettings();

            editorSettings.setCountry("IN");
            editorSettings.setLanguage("en");

            createSheetParameters.setEditorSettings(editorSettings);

            var permissions = new Map();

            permissions.set("document.export", true);
            permissions.set("document.print", false);
            permissions.set("document.edit", true);

            createSheetParameters.setPermissions(permissions);

            var callbackSettings = new SDK.V1.SheetCallbackSettings();
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
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let sheetResponseObject = responseObject.object;
    
                if(sheetResponseObject != null){

                    //Check if expected CreateSheetResponse instance is received
                    if(sheetResponseObject instanceof SDK.V1.CreateSheetResponse){
                        console.log("\nSpreadSheet ID - " + sheetResponseObject.getDocumentId());
                        console.log("\nSpreadSheet session 1 ID - " + sheetResponseObject.getSessionId());
                        console.log("\nSpreadSheet session 1 URL - " + sheetResponseObject.getDocumentUrl());
                        console.log("\nSpreadSheet session 1 Grid View URL - " + sheetResponseObject.getGridviewUrl());
                        console.log("\nSpreadSheet session 1 save URL - " + sheetResponseObject.getSaveUrl());
                        console.log("\nSpreadSheet session 1 delete URL - " + sheetResponseObject.getSessionDeleteUrl());
                        console.log("\nSpreadSheet delete URL - " + sheetResponseObject.getDocumentDeleteUrl());

                        //Session 2 creation for collaboration demo
                        var user2Info = new SDK.V1.SheetUserSettings();

                        user2Info.setDisplayName("John");

                        createSheetParameters.setUserInfo(user2Info);

                        responseObject = await sdkOperations.createSheet(createSheetParameters);

                        if(responseObject != null) {
                            //Get the status code from response
                            console.log("\nStatus Code: " + responseObject.statusCode);
                
                            //Get the api response object from responseObject
                            var sheetSession2ResponseObject = responseObject.object;
                
                            if(sheetSession2ResponseObject != null){

                                //Check if expected CreateSheetResponse instance is received
                                if(sheetSession2ResponseObject instanceof SDK.V1.CreateSheetResponse){
                                    console.log("\nSpreadSheet ID - " + sheetSession2ResponseObject.getDocumentId());
                                    console.log("\nSpreadSheet session 2 ID - " + sheetSession2ResponseObject.getSessionId());
                                    console.log("\nSpreadSheet session 2 URL - " + sheetSession2ResponseObject.getDocumentUrl());
                                    console.log("\nSpreadSheet session 2 Grid View URL - " + sheetSession2ResponseObject.getGridviewUrl());
                                    console.log("\nSpreadSheet session 2 save URL - " + sheetSession2ResponseObject.getSaveUrl());
                                    console.log("\nSpreadSheet session 2 delete URL - " + sheetSession2ResponseObject.getSessionDeleteUrl());
                                    console.log("\nSpreadSheet delete URL - " + sheetSession2ResponseObject.getDocumentDeleteUrl());
                                } else if (sheetSession2ResponseObject instanceof SDK.V1.InvalidConfigurationException) {
                                    console.log("\nInvalid configuration exception. Exception json - ", sheetSession2ResponseObject);
                                } else {
                                    console.log("\nRequest not completed successfullly");
                                }
                            } else if (sheetSession2ResponseObject instanceof SDK.V1.InvalidConfigurationException) {
                                console.log("\nInvalid configuration exception in session 2 creation. Exception json - ", sheetResponseObject);
                            } else {
                                console.log("\nRequest not completed successfullly");
                            }
                        }
                    } else if (sheetResponseObject instanceof SDK.V1.InvalidConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", sheetResponseObject);
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

EditSheet.execute();
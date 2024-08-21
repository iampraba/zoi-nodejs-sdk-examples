import * as SDK from "@zoho-corp/office-integrator-sdk";

class DeleteDocumentSession {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var createDocumentParameters = new SDK.V1.CreateDocumentParameters();
            var sdkOperations = new SDK.V1.V1Operations();

            var createResponse = await sdkOperations.createDocument(createDocumentParameters);

            var sessionId = createResponse.object.getSessionId();

            console.log("\nCreated a session to demonstrate the session delete api. Created document session ID - " + sessionId);

            var responseObject = await sdkOperations.deleteSession(sessionId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    if(writerResponseObject instanceof SDK.V1.DocumentSessionDeleteSuccessResponse){
                        console.log("\nSession delete status - " + writerResponseObject.getSessionDeleted());
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

DeleteDocumentSession.execute();
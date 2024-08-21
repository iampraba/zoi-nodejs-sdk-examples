import * as SDK from "@zoho-corp/office-integrator-sdk";

class GetDocumentDetail {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var createDocumentParameters = new SDK.V1.CreateDocumentParameters();

            var createResponse = await sdkOperations.createDocument(createDocumentParameters);

            var documentId = createResponse.object.getDocumentId();

            console.log("\nCreated a new document to demonstrate get document details api. Created document ID - " + documentId);

            var responseObject = await sdkOperations.getDocumentInfo(documentId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    //TODO: Need to fix object type issue
                    if(writerResponseObject instanceof SDK.V1.DocumentMeta ){
                        console.log("\nDocument document ID - " + writerResponseObject.getDocumentId());
                        console.log("\nDocument Name - " + writerResponseObject.getDocumentName());
                        console.log("\nDocument Type - " + writerResponseObject.getDocumentType());
                        console.log("\nDocument Collaborators Count - " + writerResponseObject.getCollaboratorsCount());
                        console.log("\nDocument Created Time - " + writerResponseObject.getCreatedTime());
                        console.log("\nDocument Created Timestamp - " + writerResponseObject.getCreatedTimeMs());
                        console.log("\nDocument Expiry Time - " + writerResponseObject.getExpiresOn());
                        console.log("\nDocument Expiry Timestamp - " + writerResponseObject.getExpiresOnMs());
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

GetDocumentDetail.execute();
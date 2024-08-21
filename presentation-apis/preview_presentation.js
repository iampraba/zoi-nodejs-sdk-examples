import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync } from 'fs';
const __dirname = import.meta.dirname;

class PreviewPresentation {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var previewParameters = new SDK.V1.PresentationPreviewParameters();

            //Either use url as document source or attach the document in request body use below methods
            previewParameters.setUrl("https://demo.office-integrator.com/samples/show/Zoho_Show.pptx");
            
            // var fileName = "Zoho_Show.pptx";
            // var filePath = __dirname + "/sample_documents/Zoho_Show.pptx";
            // var fileStream = readFileSync(filePath);
            // var streamWrapper = new SDK.StreamWrapper(fileName, fileStream, filePath);
            
            //previewParameters.setDocument(streamWrapper);

            var documentInfo = new SDK.V1.DocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("Zoho_Show.pptx");

            previewParameters.setDocumentInfo(documentInfo);

            //TODO: Language params issue
            previewParameters.setLanguage("en");

            var responseObject = await sdkOperations.createPresentationPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let previewResponseObject = responseObject.object;
    
                if(previewResponseObject != null){

                    console.log("\nPreview URL : " + previewResponseObject.getPreviewUrl());

                    if(previewResponseObject instanceof SDK.V1.PreviewResponse){
                        console.log("\nPresentation ID - " + previewResponseObject.getDocumentId());
                        console.log("\nPresentation session ID - " + previewResponseObject.getSessionId());
                        console.log("\nPresentation preview URL - " + previewResponseObject.getPreviewUrl());
                        console.log("\nPresentation delete URL - " + previewResponseObject.getDocumentDeleteUrl());
                        console.log("\nPresentation session delete URL - " + previewResponseObject.getSessionDeleteUrl());
                    } else if (previewResponseObject instanceof SDK.V1.InvalidConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", previewResponseObject);
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

PreviewPresentation.execute();
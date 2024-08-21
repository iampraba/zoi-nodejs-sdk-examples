import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync } from 'fs';
const __dirname = import.meta.dirname;

class GetFields {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var parameter = new SDK.V1.GetMergeFieldsParameters();

            //Either use url as document source or attach the document in request body use below methods
            parameter.setFileUrl("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");

            // var fileName = "OfferLetter.zdoc";
            // var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            // var fileStream = readFileSync(filePath);
            // var streamWrapper = new SDK.StreamWrapper(fileName, fileStream, filePath)
            
            // parameter.setFileContent(streamWrapper);

            var responseObject = await sdkOperations.getMergeFields(parameter);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof SDK.V1.MergeFieldsResponse) {
                        writerResponseObject.getMerge().forEach(mergeField => {
                            console.log("\n", mergeField);
                        });
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

GetFields.execute();
import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync, writeFileSync } from 'fs';
const __dirname = import.meta.dirname;

class WatermarkDocument {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var watermarkParameters = new SDK.V1.WatermarkParameters();

            //Either use url as document source or attach the document in request body use below methods
            watermarkParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            // var fileName = "Graphic-Design-Proposal.docx";
            // var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            // var fileStream = readFileSync(filePath);
            // var streamWrapper = new SDK.StreamWrapper(fileName, fileStream, filePath)

            // watermarkParameters.setDocument(streamWrapper);

            var watermarkSettings = new SDK.V1.WatermarkSettings();

            watermarkSettings.setType("text");
            watermarkSettings.setFontSize(18);
            watermarkSettings.setOpacity(70.00);
            watermarkSettings.setFontName("Arial");
            watermarkSettings.setFontColor("#cd4544");
            watermarkSettings.setOrientation("horizontal");
            watermarkSettings.setText("Sample Water Mark Text");

            watermarkParameters.setWatermarkSettings(watermarkSettings);

            var responseObject = await sdkOperations.createWatermarkDocument(watermarkParameters);

            //SDK Error: TypeError: Cannot read properties of null (reading 'getWrappedResponse')

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    
                    //Check if expected FileBodyWrapper instance is received
                    if(writerResponseObject instanceof SDK.V1.FileBodyWrapper) {
                        var watermarkDocument = writerResponseObject.getFile();
                        var outputFilePath = __dirname + "/sample_documents/watermark_output.docx";

                        if (watermarkDocument instanceof SDK.StreamWrapper) {
                            writeFileSync(outputFilePath, watermarkDocument.getStream());
                            console.log("\nCheck watermarked pdf output file in file path - ", outputFilePath);
                        }
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

WatermarkDocument.execute();
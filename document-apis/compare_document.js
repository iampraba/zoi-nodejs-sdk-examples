import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync } from 'fs';
const __dirname = import.meta.dirname;

class CompareDocument {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var compareDocumentParameters = new SDK.V1.CompareDocumentParameters();

            //Either use url as document source or attach the document in request body use below methods
            compareDocumentParameters.setUrl1("https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx");
            compareDocumentParameters.setUrl2("https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx");

            var file1Name = "MS_Word_Document_v0.docx";
            // var file1Path = __dirname + "/sample_documents/MS_Word_Document_v0.docx";
            // var file1Stream = readFileSync(file1Path);
            // var stream1Wrapper = new SDK.StreamWrapper(file1Name, file1Stream, file1Path);

            var file2Name = "MS_Word_Document_v1.docx";
            // var file2Path = __dirname + "/sample_documents/MS_Word_Document_v1.docx";
            // var file2Stream = readFileSync(file2Path);
            // var stream2Wrapper = new SDK.StreamWrapper(file2Name, file2Stream, file2Path);
            
            // compareDocumentParameters.setDocument1(stream1Wrapper);
            // compareDocumentParameters.setDocument2(stream2Wrapper);

            compareDocumentParameters.setLang("en");
            compareDocumentParameters.setTitle(file1Name + " vs " + file2Name);

            var responseObject = await sdkOperations.compareDocument(compareDocumentParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    //Check if expected CompareDocumentResponse instance is received
                    if(writerResponseObject instanceof SDK.V1.CompareDocumentResponse) {
                        console.log("\nCompare URL - " + writerResponseObject.getCompareUrl());
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

CompareDocument.execute();
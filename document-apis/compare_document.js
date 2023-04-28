const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const CompareDocumentResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/compare_document_response").CompareDocumentResponse;
const CompareDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/compare_document_parameters").CompareDocumentParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class CompareDocument {

    //Include zoi-nodejs-sdk package in your package json and the execute this code.

    static async initializeSdk() {
        let user = new UserSignature("john@zylker.com");
        let environment = new Environment("https://api.office-integrator.com", null, null);
        let apikey = new APIKey("2ae438cf864488657cc9754a27daa480", Constants.PARAMS);
        let logger = new LogBuilder()
            .level(Levels.INFO)
            .filePath("./app.log")
            .build();
        let initialize = await new InitializeBuilder();

        await initialize.user(user).environment(environment).token(apikey).logger(logger).initialize();

        console.log("\nSDK initialized successfully.");
    }

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new V1Operations();
            var compareDocumentParameters = new CompareDocumentParameters();

            compareDocumentParameters.setUrl1("https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx");
            compareDocumentParameters.setUrl2("https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx");

            var file1Name = "MS_Word_Document_v0.docx";
            // var file1Path = __dirname + "/sample_documents/MS_Word_Document_v0.docx";
            // var file1Stream = fs.readFileSync(file1Path);
            // var stream1Wrapper = new StreamWrapper(file1Name, file1Stream, file1Path)
            // var stream1Wrapper = new StreamWrapper(null, null, file1Path)

            var file2Name = "MS_Word_Document_v1.docx";
            // var file2Path = __dirname + "/sample_documents/MS_Word_Document_v1.docx";
            // var file2Stream = fs.readFileSync(file2Path);
            // var stream2Wrapper = new StreamWrapper(file2Name, file2Stream, file2Path)
            // var stream2Wrapper = new StreamWrapper(null, null, file2Path)
            
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
                    if(writerResponseObject instanceof CompareDocumentResponse) {
                        console.log("\nCompare URL - " + writerResponseObject.getCompareUrl());
                        console.log("\nDocument session delete URL - " + writerResponseObject.getSessionDeleteUrl());
                    } else if (writerResponseObject instanceof InvaildConfigurationException) {
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
}

CompareDocument.execute();
const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const FileBodyWrapper = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/file_body_wrapper").FileBodyWrapper;
const DocumentConversionParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_conversion_parameters").DocumentConversionParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;
const DocumentConversionOutputOptions = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_conversion_output_options").DocumentConversionOutputOptions;

class ConvertDocument {

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

        console.log("SDK initialized successfully.");
    }

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new V1Operations();
            var documentConversionParameters = new DocumentConversionParameters();

            //Either use url as document source or attach the document in request body use below methods
            documentConversionParameters.setUrl("https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx");

            // var fileName = "Graphic-Design-Proposal.docx";
            // var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            // var streamWrapper = new StreamWrapper(null, null, filePath)

            // documentConversionParameters.setDocument(streamWrapper);

            var outputOptions = new DocumentConversionOutputOptions();

            outputOptions.setFormat("pdf");
            outputOptions.setDocumentName("conversion_output.pdf");
            outputOptions.setIncludeComments("all");
            outputOptions.setIncludeChanges("all");

            documentConversionParameters.setOutputOptions(outputOptions);
            documentConversionParameters.setPassword("***");
            var responseObject = await sdkOperations.convertDocument(documentConversionParameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;

                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof FileBodyWrapper) {
                        var convertedDocument = writerResponseObject.getFile();

                        if (convertedDocument instanceof StreamWrapper) {
                            var outputFilePath = __dirname + "/sample_documents/conversion_output.pdf";

                            fs.writeFileSync(outputFilePath, convertedDocument.getStream());
                            console.log("\nCheck converted output file in file path - ", outputFilePath);
                        }
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

ConvertDocument.execute();
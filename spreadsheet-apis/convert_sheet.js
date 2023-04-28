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
const SheetConversionParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/sheet_conversion_parameters").SheetConversionParameters;
const SheetConversionOutputOptions = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/sheet_conversion_output_options").SheetConversionOutputOptions;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class ConvertSheet {

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
            var sheetConversionParameters = new SheetConversionParameters();

            //Either use url as document source or attach the document in request body use below methods
            sheetConversionParameters.setUrl("https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx");

            // var fileName = "Contact_List.xlsx";
            // var filePath = __dirname + "/sample_documents/Contact_List.xlsx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            // var streamWrapper = new StreamWrapper(null, null, filePath)

            // sheetConversionParameters.setDocument(streamWrapper);

            var outputOptions = new SheetConversionOutputOptions();

            outputOptions.setFormat("pdf");
            outputOptions.setDocumentName("ConvertedSheet.pdf");

            sheetConversionParameters.setOutputOptions(outputOptions);

            var responseObject = await sdkOperations.convertSheet(sheetConversionParameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let sheetResponseObject = responseObject.object;

                if(sheetResponseObject != null) {
                    if(sheetResponseObject instanceof FileBodyWrapper) {
                        var convertedDocument = sheetResponseObject.getFile();

                        if (convertedDocument instanceof StreamWrapper) {
                            var outputFilePath = __dirname + "/sample_documents/conversion_output.pdf";

                            fs.writeFileSync(outputFilePath, convertedDocument.getStream());
                            console.log("\nCheck converted output file in file path - ", outputFilePath);
                        }
                    } else if (sheetResponseObject instanceof InvaildConfigurationException) {
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
}

ConvertSheet.execute();
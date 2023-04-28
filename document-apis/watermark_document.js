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
const WatermarkSettings = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/watermark_settings").WatermarkSettings;
const WatermarkParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/watermark_parameters").WatermarkParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class WatermarkDocument {

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
            var watermarkParameters = new WatermarkParameters();

            watermarkParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            // var fileName = "Graphic-Design-Proposal.docx";
            // var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath)
            //var streamWrapper = new StreamWrapper(null, null, filePath)

            // watermarkParameters.setDocument(streamWrapper);

            var watermarkSettings = new WatermarkSettings();

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
                    if(writerResponseObject instanceof FileBodyWrapper) {
                        var watermarkDocument = writerResponseObject.getFile();
                        var outputFilePath = __dirname + "/sample_documents/watermark_output.docx";

                        if (watermarkDocument instanceof StreamWrapper) {
                            fs.writeFileSync(outputFilePath, watermarkDocument.getStream());
                            console.log("\nCheck watermarked pdf output file in file path - ", outputFilePath);
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

WatermarkDocument.execute();
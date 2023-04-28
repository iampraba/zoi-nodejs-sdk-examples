const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const PreviewResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/preview_response").PreviewResponse;
const PreviewParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/preview_parameters").PreviewParameters;
const PreviewDocumentInfo = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/preview_document_info").PreviewDocumentInfo;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;

class PreviewDocument {

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
            var previewParameters = new PreviewParameters();

            previewParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            // var fileName = "Graphic-Design-Proposal.docx";
            // var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath)
            //var streamWrapper = new StreamWrapper(null, null, filePath)
            
            // previewParameters.setDocument(streamWrapper);

            var previewDocumentInfo = new PreviewDocumentInfo();

            //Time value used to generate unique document everytime. You can replace based on your application.
            previewDocumentInfo.setDocumentName("Graphic-Design-Proposal.docx");

            previewParameters.setDocumentInfo(previewDocumentInfo);

            var permissions = new Map();

            permissions.set("document.print", false);

            previewParameters.setPermissions(permissions);

            var responseObject = await sdkOperations.createDocumentPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    console.log("\nPreview URL : " + writerResponseObject.getPreviewUrl());

                    //Check if expected PreviewResponse instance is received
                    if(writerResponseObject instanceof PreviewResponse) {
                        console.log("\nDocument ID - " + writerResponseObject.getDocumentId());
                        console.log("\nDocument session ID - " + writerResponseObject.getSessionId());
                        console.log("\nDocument preview URL - " + writerResponseObject.getPreviewUrl());
                        console.log("\nDocument delete URL - " + writerResponseObject.getDocumentDeleteUrl());
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

PreviewDocument.execute();
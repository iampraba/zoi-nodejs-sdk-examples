const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const fs = require("fs");
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;
const PreviewResponse = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/preview_response").PreviewResponse;
const SheetPreviewParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/sheet_preview_parameters").SheetPreviewParameters;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;
const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

class PreviewSheet {

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
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var previewParameters = new SheetPreviewParameters();

            previewParameters.setUrl("https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx");
            
            // var fileName = "Contact_List.xlsx";
            // var filePath = __dirname + "/sample_documents/Contact_List.xlsx";
            // var fileStream = fs.readFileSync(filePath);
            // var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            // previewParameters.setDocument(streamWrapper);

            var permissions = new Map();

            permissions.set("document.export", false);
            permissions.set("document.print", false);

            previewParameters.setPermissions(permissions);

            previewParameters.setLanguage("en");

            var responseObject = await sdkOperations.createSheetPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: ", responseObject.statusCode);
    
                //Get the api response object from responseObject
                let previewResponseObject = responseObject.object;

                console.log("\nStatus Code: ", previewResponseObject);
    
                if(previewResponseObject != null){
                    if(previewResponseObject instanceof PreviewResponse){
                        console.log("\nSheet session ID - " + previewResponseObject.getSessionId());
                        console.log("\nSheet document ID - " + previewResponseObject.getDocumentId());
                        console.log("\nSheet preview URL - " + previewResponseObject.getPreviewUrl());
                        console.log("\nSheet delete URL - " + previewResponseObject.getDocumentDeleteUrl());
                        console.log("\nSheet session delete url - " + previewResponseObject.getSessionDeleteUrl());
                    } else if (previewResponseObject instanceof InvaildConfigurationException) {
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
}

PreviewSheet.execute();
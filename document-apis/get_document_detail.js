const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;

const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/create_document_parameters").CreateDocumentParameters;
const V1Operations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/v1_operations").V1Operations;
const DocumentMeta = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/document_meta").DocumentMeta;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/v1/invaild_configuration_exception").InvaildConfigurationException;

class GetDocumentDetail {

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
            var createDocumentParameters = new CreateDocumentParameters();

            var createResponse = await sdkOperations.createDocument(createDocumentParameters);

            var documentId = createResponse.object.getDocumentId();

            console.log("\nCreated a new document to demonstrate get document details api. Created document ID - " + documentId);

            var responseObject = await sdkOperations.getDocumentInfo(documentId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){
                    //TODO: Need to fix object type issue
                    if(writerResponseObject instanceof DocumentMeta ){
                        console.log("\nDocument document ID - " + writerResponseObject.getDocumentId());
                        console.log("\nDocument Name - " + writerResponseObject.getDocumentName());
                        console.log("\nDocument Type - " + writerResponseObject.getDocumentType());
                        console.log("\nDocument Collaborators Count - " + writerResponseObject.getCollaboratorsCount());
                        console.log("\nDocument Created Time - " + writerResponseObject.getCreatedTime());
                        console.log("\nDocument Created Timestamp - " + writerResponseObject.getCreatedTimeMs());
                        console.log("\nDocument Expiry Time - " + writerResponseObject.getExpiresOn());
                        console.log("\nDocument Expiry Timestamp - " + writerResponseObject.getExpiresOnMs());
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

GetDocumentDetail.execute();
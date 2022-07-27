const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;
const CreateDocumentParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/create_document_parameters").CreateDocumentParameters;

class SDKInitializer {

    static async initialize() {
        let user = new UserSignature("praburaji93@gmail.com");
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

}

module.exports = {
    SDKInitializer: SDKInitializer
};
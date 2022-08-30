const Constants = require("zoi-nodejs-sdk/utils/util/constants").Constants;
const UserSignature = require("zoi-nodejs-sdk/routes/user_signature").UserSignature;
const Levels = require("zoi-nodejs-sdk/routes/logger/logger").Levels;
const LogBuilder = require("zoi-nodejs-sdk/routes/logger/log_builder").LogBuilder;
const Environment = require("zoi-nodejs-sdk/routes/dc/environment").Environment;
const InitializeBuilder = require("zoi-nodejs-sdk/routes/initialize_builder").InitializeBuilder;
const APIKey = require("zoi-nodejs-sdk/models/authenticator/apikey").APIKey;

const PlanDetails = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/plan_details").PlanDetails;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;
const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/officeintegrator/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

class GetPlanDetails {

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
            var responseObject = await sdkOperations.getPlanDetails();

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let planResponseObject = responseObject.object;
    
                if(planResponseObject != null) {
                    if(planResponseObject instanceof PlanDetails){
                        console.log("\nPlan name - " + planResponseObject.getPlanName());
                        console.log("\nAPI usage limit - " + planResponseObject.getUsageLimit());
                        console.log("\nAPI usage so far - " + planResponseObject.getTotalUsage());
                        console.log("\nPlan upgrade payment link - " + planResponseObject.getPaymentLink());
                        console.log("\nSubscription period - " + planResponseObject.getSubscriptionPeriod());
                        console.log("\nSubscription interval - " + planResponseObject.getSubscriptionInterval());
                    } else if (planResponseObject instanceof InvaildConfigurationException) {
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

GetPlanDetails.execute();
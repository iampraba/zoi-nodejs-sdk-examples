import * as SDK from "@zoho-corp/office-integrator-sdk";



class GetPlanDetails {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in you application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var responseObject = await sdkOperations.getPlanDetails();

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let planResponseObject = responseObject.object;
    
                if(planResponseObject != null) {
                    if(planResponseObject instanceof SDK.V1.PlanDetails){
                        console.log("\nPlan name - " + planResponseObject.getPlanName());
                        console.log("\nAPI usage limit - " + planResponseObject.getUsageLimit());
                        console.log("\nAPI usage so far - " + planResponseObject.getTotalUsage());
                        console.log("\nPlan upgrade payment link - " + planResponseObject.getPaymentLink());
                        console.log("\nSubscription period - " + planResponseObject.getSubscriptionPeriod());
                        console.log("\nSubscription interval - " + planResponseObject.getSubscriptionInterval());
                    } else if (planResponseObject instanceof SDK.V1.InvalidConfigurationException) {
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

GetPlanDetails.execute();
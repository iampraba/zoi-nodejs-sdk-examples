const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const GetMergeFieldsParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/get_merge_fields_parameters").GetMergeFieldsParameters;

const MergeFieldsResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/merge_fields_response").MergeFieldsResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class GetFields {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var parameter = new GetMergeFieldsParameters();

            //TODO  //Set url not working
            //parameter.setFileUrl("https://demo.office-integrator.com/zdocs/LabReport.zdoc");

            var fileName = "OfferLetter.zdoc";
            var filePath = __dirname + "/sample_documents/OfferLetter.zdoc";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath)
            
            parameter.setFileContent(streamWrapper);

            var responseObject = await sdkOperations.getMergeFields(parameter);

            if(responseObject != null) {
                console.log("Status Code: " + responseObject.statusCode);
    
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    if(writerResponseObject instanceof MergeFieldsResponse) {
                        writerResponseObject.getMerge().forEach(mergeField => {
                            console.log(mergeField);
                        });
                    } else if (writerResponseObject instanceof InvaildConfigurationException) {
                        console.log("Invalid configuration exception. Exception json - ", writerResponseObject);
                    } else {
                        console.log("Request not completed successfullly");
                    }
                }
            }
        } catch (error) {
            console.log("Exception while running sample code", error);
        }
    }
}

GetFields.execute();
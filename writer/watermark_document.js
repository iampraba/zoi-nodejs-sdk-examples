const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const WatermarkParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/watermark_parameters").WatermarkParameters;

const WatermarkSettings = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/watermark_settings").WatermarkSettings;

const FileBodyWrapper = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/file_body_wrapper").FileBodyWrapper;

const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class WatermarkDocument {

    static async execute() {
        
        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var watermarkParameters = new WatermarkParameters();

            //watermarkParameters.setUrl("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx");

            var fileName = "Graphic-Design-Proposal.docx";
            var filePath = __dirname + "/sample_documents/Graphic-Design-Proposal.docx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath)
            //var streamWrapper = new StreamWrapper(null, null, filePath)

            watermarkParameters.setDocument(streamWrapper);

            var watermarkSettings = new WatermarkSettings();

            watermarkSettings.setType("text");
            watermarkSettings.setText("Sample Water Mark Text");
            watermarkSettings.setFontColor("#cd4544");
            watermarkSettings.setFontName("Arial");
            watermarkSettings.setOpacity("70");
            watermarkSettings.setOrientation("horizontal");
            watermarkSettings.setFontSize("18");

            watermarkParameters.setWatermarkSettings(watermarkSettings);

            var responseObject = await sdkOperations.createWatermarkDocument(watermarkParameters);

            //SDK Error: TypeError: Cannot read properties of null (reading 'getWrappedResponse')

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null) {
                    
                    //Check if expected FileBodyWrapper instance is received
                    if(writerResponseObject instanceof FileBodyWrapper) {
                        var watermarkDocument = writerResponseObject.getFile();

                        if (convertedDowatermarkDocumentcument instanceof StreamWrapper) {
                            fs.writeFileSync("./sample_documents/watermark_output.pdf", watermarkDocument.getStream());
                        }
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

WatermarkDocument.execute();
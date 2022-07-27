const fs = require("fs");
const SDKInitializer = require("../SDKInitializer").SDKInitializer;
const StreamWrapper = require("zoi-nodejs-sdk/utils/util/stream_wrapper").StreamWrapper;

const OfficeIntegratorSDKOperations = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/office_integrator_sdk_operations").OfficeIntegratorSDKOperations;

const SheetPreviewParameters = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/sheet_preview_parameters").SheetPreviewParameters;

const SheetPreviewResponse = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/sheet_preview_response").SheetPreviewResponse;
const InvaildConfigurationException = require("zoi-nodejs-sdk/core/com/zoho/crm/api/office_integrator_sdk/invaild_configuration_exception").InvaildConfigurationException;

class PreviewSheet {

    static async execute() {

        /** Initializing once is enough. Calling here since code sample will be tested standalone. 
          * You can place SDKInitializer in you application and call while start-up. 
          */
         await SDKInitializer.initialize();

        try {
            var sdkOperations = new OfficeIntegratorSDKOperations();
            var previewParameters = new SheetPreviewParameters();

            //previewParameters.setUrl("https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx");
            
            var fileName = "Contact_List.xlsx";
            var filePath = __dirname + "/sample_documents/Contact_List.xlsx";
            var fileStream = fs.readFileSync(filePath);
            var streamWrapper = new StreamWrapper(fileName, fileStream, filePath);
            
            previewParameters.setDocument(streamWrapper);

            var permissions = new Map();

            permissions.set("document.export", false);
            permissions.set("document.print", false);

            previewParameters.setPermissions(permissions);

            previewParameters.setLanguage("en");

            var responseObject = await sdkOperations.createSheetPreview(previewParameters);

            if(responseObject != null) {
                //Get the status code from response
                console.log("Status Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let writerResponseObject = responseObject.object;
    
                if(writerResponseObject != null){

                    if(writerResponseObject instanceof SheetPreviewResponse){
                        console.log("Document ID - " + writerResponseObject.getDocumentId());
                        console.log("Document preview URL - " + writerResponseObject.getPreviewUrl());
                        console.log("Document delete URL - " + writerResponseObject.getDocumentDeleteUrl());
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

PreviewSheet.execute();
const express = require('express');
const multer = require('multer');
const fs = require('fs');
//To get absolute file path to writer file read from post request
const { resolve } = require('path');

const app = express();
const upload = multer();

// Start the server
const port = 3000;

app.listen(port, () => {
  console.log(`\nServer listening on port ${port}`);

  console.log(`\nMake a post request to this end point to test save callback from local file - http://localhost/${port}${callbackEndpoint}`);

  console.log(`\nKeep the multipart key as content for file. If you send the file in different key then change the key in save_callback code as well before testing`);

  console.log(`\nYou should not use http://localhost/${port} as save_url in callback settings. You should only use it as publicly accessible(with your application domain) end point.`);
});

var showCallback = async (req, res) => {
  try {
    // multer middleware will read the file and store it to file key in req object
    const fileData = req.file;
    const outputFilePath = resolve(".") + "/" + fileData.originalname;

    //Writing file to localdisk. Use can write the file in any of your storage service. Modify the code accordingly
    fs.writeFileSync( outputFilePath, fileData.buffer); 
    console.log('\nFile uploaded successfully in file path - ' + outputFilePath);

    // Send a response back to Zoho API to acknowledge the callback successfully completed.
    // String value given for message key below will shown as save status in editor UI.
    res.status(200).json({ message: "Document Save Successfully in NodeJS Server."});
  } catch (err) {
    console.error('Error in saving file.', err);

    // You can send an appropriate error response back to Zoho Show API here if needed.
    // String value given for message key below will shown as save status in editor UI.
    res.status(500).json({ message: 'Error in saving file' });
  }
};

//saveurl callback api end point
const callbackEndpoint = "/zoho/file/uploader";

app.post(callbackEndpoint, upload.single('content'), showCallback);
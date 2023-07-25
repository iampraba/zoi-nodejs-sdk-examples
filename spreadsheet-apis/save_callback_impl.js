const express = require('express');
const multer = require('multer');

const app = express();
const upload = multer();

app.post('/zoho/file/uploader', upload.single('content'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  var filename = req.body.filename;
  var format = req.body.format;
  //if you configured save_url_params while session creation, e.g "key1" : "value1"
  // then you can read that key from request body
  var customkey1 = req.body.key1;

  console.log("filename - ", filename);
  console.log("format - ", format);
  console.log("customkey1 - ", customkey1);

  // Access the file using req.file
  const file = req.file;

  // Perform operations on the file
  // For example, you can save it to disk or process it in some way

  // Send a response indicating successful upload
  res.json({ message: 'File uploaded successfully' });
});

// Start the server
const port = 3000;

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
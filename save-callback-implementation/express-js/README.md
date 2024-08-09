Save the [code](https://raw.githubusercontent.com/zoho/office-integrator-nodejs-sdk-examples/main/save-callback-implementation/express-js/save_callback.js) in a file. for example, save_callback.js.

To run the following commands from your Mac terminal, you should have nodejs installed on you machine. 

Open the terminal, navigate to the directory save-callback-implementation, and run the following command:

 ``` sh
npm install
  ```

  ``` sh
node save_callback.js
  ```

This will start the api end point to receive file. You can access the API endpoint at http://localhost:3000/zoho/file/uploader.

Test above api end point and verify if the file store without any error. You could free to edit the code to fit your application needs. 

To host this code on a server so that it can be accessed from the internet, you would need a web server with nodejs support. Here are the general steps:

- Choose a hosting provider or set up your own server with a web server software like Apache or Nginx.
- Upload the your save_callback.js file to your server.
- Make sure the server has nodejs installed and properly configured.
- Ensure that the destination folder for uploaded files has appropriate write permissions.
- Access the API endpoint using the server's domain or IP address, e.g., http://zylker.com//zoho/file/uploader.

Note that the exact steps may vary depending on your hosting environment and configuration. It's recommended to refer to the documentation provided by your hosting provider or the server software you choose for more detailed instructions.  

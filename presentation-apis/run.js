import { exec } from 'child_process';

// List all the Node.js files you want to run in an array
const files = ["coedit_presentation.js", "create_presentation.js", "delete_presentation_session.js", "get_session_detail.js", "convert_presentation.js", "delete_presentation.js", "edit_presentation.js", "preview_presentation.js"];

// Loop through the array and run each file using the 'node' command
var i = 1;
files.forEach(file => {
  exec(`node ${file}`, (error, stdout, stderr) => {
    if (error) {
      console.error(`Error running ${file}: ${error}`);
      return;
    }
    console.log(`${i++} - ${file} output: ${stdout}`);
  });
});

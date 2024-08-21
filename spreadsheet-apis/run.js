import { exec } from 'child_process';

// List all the Node.js files you want to run in an array
const files = ["coedit_sheet.js", "create_sheet.js", "delete_sheet_session.js", "get_session_detail.js", "convert_sheet.js",	"delete_sheet.js", "edit_sheet.js", "preview_sheet.js"];

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

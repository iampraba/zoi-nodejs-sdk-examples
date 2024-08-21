import { exec } from 'child_process';

// List all the Node.js files you want to run in an array
const files = ["create_document.js", "edit_document.js", "get_session_detail.js", "coedit_document.js", "create_merge_template.js", "get_all_sessions.js", "merge_and_deliver.js", "watermark_document.js", "compare_document.js", "delete_document.js", "get_document_detail.js", "merge_and_download.js", "convert_document.js", "delete_document_session.js", "get_merge_fields.js", "preview_document.js"];

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

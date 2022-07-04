let http = require('http');
let formidable = require('formidable');
let fs = require('fs');

http.createServer(function (req, res) {

  //Create an instance of the form object
  let form = new formidable.IncomingForm();

  //Process the file upload in Node
  form.parse(req, function (error, fields, file) {
    let filepath = file.fileupload.filepath;
    let newpath = 'wordDocument.docx';
    

    //Copy the uploaded file to a custom folder
    fs.rename(filepath, newpath, function () {
      //Send a NodeJS file upload confirmation message
      res.write('NodeJS File Upload Success!');
      res.write(newpath);
      res.end();
    });
    
  });

  
}).listen(80);
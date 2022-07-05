let http = require('http');
let formidable = require('formidable');
let fs = require('fs');
let path = require('path');
const ParserBoundaries = require('./ParserBoundaries');
const ExcellWriterEngine=require('./excelWriterEngine');
const ParserEngine=require('./parserEngine');
const express=require('express');
const app=express();
const PORT=3000;
app.listen(PORT,()=>{
  console.log(`Listen on port ${PORT}`)
})
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
    const textract = require('textract');

textract.fromFileWithPath('./wordDocument.docx',function(error,text){
  // console.log(text);
    let protoParser=new ParserEngine(text);
    let currentAccountType=false;
    if(text.includes('ЗА ОСОБЕН ЗАЛОГ НА ВЗЕМАНИЯ ЗА НАЛИЧНОСТИ ПО СМЕТКА')){
        currentAccountType=true;
    }
    let protoBoundaries=new ParserBoundaries();
    let text1=protoParser.extractTextByBoundaryStrings(protoBoundaries.protoParser.lowerBoundary,protoBoundaries.protoParser.upperBoundary)
    let parser=new ParserEngine(text1);
    let boundaries=new ParserBoundaries();
    let requestorName=parser.extractTextByBoundaryStrings(boundaries.requestorName.lowerBoundary,boundaries.requestorName.upperBoundary);
    let requestorEIK=parser.extractTextByBoundaryStrings(boundaries.requestorEIK.lowerBoundary,boundaries.requestorEIK.upperBoundary);
    let requestorAddress=parser.extractTextByBoundaryStrings(boundaries.requestorAddress.lowerBoundary,boundaries.requestorAddress.upperBoundary);
    let pledgerName=parser.extractTextByBoundaryStrings(boundaries.pledgerName.lowerBoundary,boundaries.pledgerName.upperBoundary);
    let pledgerEIK=parser.extractTextByBoundaryStrings(boundaries.pledgerEIK.lowerBoundary,boundaries.pledgerEIK.upperBoundary)
    let pledgerAdress=parser.extractTextByBoundaryStrings(boundaries.pledgerAdress.lowerBoundary,boundaries.pledgerAdress.upperBoundary);
    let loanBL=parser.extractTextByBoundaryStrings(boundaries.loanBL.lowerBoundary,boundaries.loanBL.upperBoundary);
    let loanInterestBase=parser.extractTextByBoundaryStrings(boundaries.loanInterestBase.lowerBoundary,boundaries.loanInterestBase.upperBoundary);
    loanInterestBase=loanInterestBase.replace('Бизнес клиенти','БК')
    let loanInterestSpread=parser.extractTextByBoundaryStrings(boundaries.loanInterestSpread.lowerBoundary,boundaries.loanInterestSpread.upperBoundary);
    let loanInterestMinimum=parser.extractTextByBoundaryStrings(boundaries.loanInterestMinimum.lowerBoundary,boundaries.loanInterestMinimum.upperBoundary);
    let finalInterestRateString=`${loanInterestBase} + ${loanInterestSpread}, минимум ${loanInterestMinimum}`
    let finalInterestRateStringOverdue=`${loanInterestBase} + ${Number(ExcellWriterEngine.trimStringBeforeExcelUpload(loanInterestSpread))+10}, минимум ${Number(ExcellWriterEngine.trimStringBeforeExcelUpload(loanInterestMinimum))+10}`
    let loanAmount=parser.extractTextByBoundaryStrings(boundaries.loanAmount.lowerBoundary,boundaries.loanAmount.upperBoundary);
    let loanCollateral=parser.extractTextByBoundaryStrings(boundaries.loanCollateral.lowerBoundary,boundaries.loanAmount.upperBoundary);
    let representators=parser.extractTextByBoundaryStrings(boundaries.representators.lowerBoundary,boundaries.representators.upperBoundary);
    let representatorsArray=[];
    representators.split(' и ').forEach(representaor=>{representatorsArray.push(representaor.split('ЕГН'))});
    
    let excellFilename='./Form.xlsx'
    let excellWritter=new ExcellWriterEngine(excellFilename,requestorName,requestorEIK,requestorAddress,
    pledgerName,pledgerEIK,pledgerAdress, loanBL,finalInterestRateString,finalInterestRateStringOverdue,
    loanAmount,loanCollateral,representatorsArray,currentAccountType)
    excellWritter.prepareExcelFile()
    
    app.get('/',(request1,resource)=>{
      resource.attachment(path.resolve('./newfile25.js'))
      resource.send(200)
      
    });


})

  });


}).listen(80);



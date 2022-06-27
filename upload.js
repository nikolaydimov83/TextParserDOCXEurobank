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
  const textract = require('textract');
const reader = require('xlsx');
const ExcelJS = require('exceljs');
let _requestorName=new WeakMap();
let _requestorEIK=new WeakMap();
let _requestorAddress=new WeakMap();
let _pledgerName=new WeakMap();
let _pledgerEIK=new WeakMap();
let _pledgerAdress=new WeakMap();
let _loanBL=new WeakMap();
let _finalInterestRateString=new WeakMap();
let _finalInterestRateStringOverdue=new WeakMap();
let _loanAmount=new WeakMap();
let _loanCollateral=new WeakMap();
let _roughText=new WeakMap();
let _representatorsArray=new WeakMap();
let _fileName=new WeakMap();
let _currentAccountType=new WeakMap();


textract.fromFileWithPath('./wordDocument.docx',function(error,text){
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
    
    let excellFilename='./Образец 1 Заявление за вписване на договор за залог и за удостоверение.xlsx'
    let excellWritter=new ExcellWriterEngine(excellFilename,requestorName,requestorEIK,requestorAddress,
    pledgerName,pledgerEIK,pledgerAdress, loanBL,finalInterestRateString,finalInterestRateStringOverdue,
    loanAmount,loanCollateral,representatorsArray,currentAccountType)
    excellWritter.prepareExcelFile();
    

  
    
   

})

class ParserBoundaries{
    constructor(){
        this.protoParser={lowerBoundary:`"БАНКАТА", `,
        upperBoundary:`Чл. 3.`}
        this.requestorName={lowerBoundary:`от една страна, и`,upperBoundary:`, вписано в Търговския регистър, `}
        this.requestorEIK={lowerBoundary:`ЕИК`,upperBoundary:`, със седалище и адрес на управление`}
        this.requestorAddress={lowerBoundary:`със седалище и адрес на управление`,upperBoundary:`, представлявано от`}
        this.pledgerName={lowerBoundary:`(по-долу Договор за кредит") между`,upperBoundary:`в качеството му на Кредитополучател`}
        this.pledgerEIK={lowerBoundary:`чакам долна граница`,upperBoundary:`чакам горна граница`}
        this.pledgerAdress={lowerBoundary:`oчаквам долна граница`,upperBoundary:`очаквам горна граница`}
        this.loanBL={lowerBoundary:`е сключен Договор за банков кредит`,upperBoundary:`(по-долу Договор за кредит")`}
        this.loanInterestBase={lowerBoundary:`бизнес клиенти на Юробанк България" АД`,upperBoundary:`плюс фиксирана договорна надбавка в размер на`}
        this.loanInterestSpread={lowerBoundary:`плюс фиксирана договорна надбавка в размер на`,upperBoundary:`процентни пункта, но не по ниска от`}
        this.loanInterestMinimum={lowerBoundary:`но не по ниска от`,upperBoundary:`/ ( Долна граница на дължимата лихва"`}
        this.loanAmount={lowerBoundary:`(максимален разрешен размер на кредита):`,upperBoundary:`. Срок за издължаване:`}
        this.loanCollateral={lowerBoundary:`Предмет на настоящия договор е учредяването на`,upperBoundary:`/по-долу "Договора за сметка"/`}
        this.representators={lowerBoundary:`представлявано от`,upperBoundary:`в качеството им на управители, наричана по-долу за краткост "ЗАЛОГОДАТЕЛ"`}
    }
    
}

class ParserEngine{
constructor(roughText){
    _roughText.set(this,roughText);
}
get roughText(){
    return _roughText.get(this);
}
extractTextByBoundaryStrings(lowerBoundary,upperBoundary){
    let intermediateArrayOfStrings=this.roughText.split(lowerBoundary)
    let finalArray=intermediateArrayOfStrings[1].split(upperBoundary)
    return finalArray[0];
}
}


class ExcellWriterEngine{
    constructor(a,requestorName,requestorEIK,requestorAddress,
        pledgerName,pledgerEIK, pledgerAdress, loanBL,finalInterestRateString,
        finalInterestRateStringOverdue,loanAmount,loanCollateral,representatorsArray,currentAccountType){

        //this.file = reader.readFile(filename);
        /*let workbook = new ExcelJS.Workbook();
       workbook.xlsx.readFile(filename);
        let a=workbook.worksheets[0];*/
        _fileName.set(this,a);
        _requestorName.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(requestorName));
        _requestorEIK.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(requestorEIK));
        _requestorAddress.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(requestorAddress));
        _pledgerName.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(pledgerName));
        _pledgerEIK.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(pledgerEIK));
        _pledgerAdress.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(pledgerAdress));
        _loanBL.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(loanBL));
        _finalInterestRateString.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(finalInterestRateString)+'%');
        _finalInterestRateStringOverdue.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(finalInterestRateStringOverdue)+'%');
        _loanAmount.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(loanAmount))
        _loanCollateral.set(this,ExcellWriterEngine.trimStringBeforeExcelUpload(loanCollateral));
        _currentAccountType.set(this,currentAccountType);
        let trimmedRepArray=[];
        for (const key in representatorsArray) {
            let interArr=[];
            if (Object.hasOwnProperty.call(representatorsArray, key)) {
                
                const representatorNameEGN = representatorsArray[key];
                for (const key in representatorNameEGN) {
                    if (Object.hasOwnProperty.call(representatorNameEGN, key)) {
                        let element = representatorNameEGN[key];
                        interArr.push(ExcellWriterEngine.trimStringBeforeExcelUpload(element));
                        
                    }
                }
                trimmedRepArray.push(interArr);
            }
        }    
        _representatorsArray.set(this,trimmedRepArray);
    }

    get fileName(){
        return _fileName.get(this)
    }
    
    get requestorName(){
        return _requestorName.get(this)
    }


    get requestorEIK(){
        return _requestorEIK.get(this)
    }

    get requestorAddress(){
        return _requestorAddress.get(this)
    }

    get pledgerName(){
        return _pledgerName.get(this)
    }

    get pledgerEIK(){
        return _pledgerEIK.get(this)
    }

    get pledgerAdress(){
        return _pledgerAdress.get(this)
    }

    get loanBL(){
        return _loanBL.get(this)
    }

    get finalInterestRateString(){
        return _finalInterestRateString.get(this)
    }
    get finalInterestRateStringOverdue(){
        return _finalInterestRateStringOverdue.get(this)
    }
    get loanAmount(){
        return _loanAmount.get(this);
    }
    get loanCollateral(){
        return _loanCollateral.get(this);
    }
    get representatorsArray(){
        return _representatorsArray.get(this);
    }
    get currentAccountType(){
        return _currentAccountType.get(this);
    }

    static trimStringBeforeExcelUpload(string){
        let trimmedString=string;
        let a=trimmedString.replace(/^[^a-zа-я\d]*|[^a-zа-я\d]*$/gi, '');
        return a; 
    }
    prepareExcelFile(){
        const wb = new ExcelJS.Workbook();

        const fileName = 'Образец 1 Заявление за вписване на договор за залог и за удостоверение.xlsx';
        
        wb.xlsx.readFile(fileName).then(() => {
            
            const ws = wb.getWorksheet('Вписване');
        
            let cellActive = ws.getCell('A5');
            cellActive.value = this.requestorName;
            this.splitAndFulfillEIK('Z',5,this.requestorEIK,11,ws);
            cellActive=ws.getCell('A7');
            cellActive.value=this.requestorAddress;
            cellActive=ws.getCell('A17');
            cellActive.value=this.requestorName;
            this.splitAndFulfillEIK('Z',17,this.requestorEIK,11,ws);
            cellActive=ws.getCell('A19');
            cellActive.value=this.requestorAddress;
            cellActive=ws.getCell('A24');
            cellActive.value=this.pledgerName;
            this.splitAndFulfillEIK('Z',24,this.pledgerEIK,11,ws);
            cellActive=ws.getCell('A26');
            cellActive.value=this.pledgerAdress;
            cellActive=ws.getCell('A37');
            cellActive.value=this.loanBL;
            cellActive=ws.getCell('AC37');
            cellActive.value=this.finalInterestRateString;
            cellActive=ws.getCell('A39');
            cellActive.value=this.loanAmount;
            cellActive=ws.getCell('AC39');
            cellActive.value=this.finalInterestRateStringOverdue;
            cellActive=ws.getCell('A47');
            cellActive.value=this.loanCollateral;
            cellActive=ws.getCell('A64');
            cellActive.value=this.representatorsArray[0][0];
            this.splitAndFulfillEIK('A',66,this.representatorsArray[0][1],10,ws);
            if (this.representatorsArray.length>1){
                cellActive=ws.getCell('S64');
                cellActive.value=this.representatorsArray[1][0];
                this.splitAndFulfillEIK('S',66,this.representatorsArray[1][1],10,ws);
            }
            if (this.currentAccountType===true){
                cellActive=ws.getCell('D53');
                cellActive.value='НЕ'
            }

           /* let numberOfCells=11;
            let cellsEIKArray=ExcellWriterEngine.createValidExcellColumnNames('Z',numberOfCells);
            let requestorEIKAsArr=this.trimStringBeforeExcelUpload(this.requestorEIK).split('');
            let lastPositionWithTilda=numberOfCells-requestorEIKAsArr.length-1;
            for (const key in cellsEIKArray) {
                if (Object.hasOwnProperty.call(cellsEIKArray, key)) {
                    const element = cellsEIKArray[key]+'5';
                    cellActive=ws.getCell(element)
                    if(key<=lastPositionWithTilda){     
                        cellActive.value='~'
                    }else{
                        cellActive.value=requestorEIKAsArr[key-(lastPositionWithTilda+1)]
                    }
                }
            }*/
            wb.xlsx
            .writeFile('newfile25.xlsx')
            .then(() => {
              console.log('file created');
              // documentLink=document.createElement()
              
            })
            .catch(err => {
              console.log(err.message);
            });
        }).catch(err => {
            console.log(err.message);
        });
    }
    splitAndFulfillEIK(startCol,startRow,eik,numberOfCells,ws){
        let cellsEIKArray=ExcellWriterEngine.createValidExcellColumnNames(startCol,numberOfCells);
        let requestorEIKAsArr=eik.split('');
        let lastPositionWithTilda=numberOfCells-requestorEIKAsArr.length-1;
        for (const key in cellsEIKArray) {
            if (Object.hasOwnProperty.call(cellsEIKArray, key)) {
                const element = cellsEIKArray[key]+startRow;
                let cellActive=ws.getCell(element)
                if(key<=lastPositionWithTilda){     
                    cellActive.value='~'
                }else{
                    cellActive.value=requestorEIKAsArr[key-(lastPositionWithTilda+1)]
                }
            }
        }
    }
    static createValidExcellColumnNames(startString,numberOfColumns){
        let arrayOfCharsAsNumber=[]
        startString.split('').forEach(char=>{
            arrayOfCharsAsNumber.push(char.charCodeAt(0))
        })
        let finalArray=[];
        for (let i = 1; i <= numberOfColumns; i++) {
            
            for (let j=arrayOfCharsAsNumber.length-1;j>=0;j--){
                if(arrayOfCharsAsNumber[j]<=89){
                    arrayOfCharsAsNumber[j]++
                    finalArray.push(arrayOfCharsAsNumber.toString())
                    break
                }else if(arrayOfCharsAsNumber[j]===90){
                    if(j===0){
                       arrayOfCharsAsNumber.unshift(65); 
                    }
                    for (let k = j; k < arrayOfCharsAsNumber.length; k++) {
                        arrayOfCharsAsNumber[k]=65;
                        if(j!==0){
                            arrayOfCharsAsNumber[k-1]++
                        }
                        
                        
                    }
                    finalArray.push(arrayOfCharsAsNumber.toString())
                }
                
            }
            
        }
    let excelValidColsArray=[]
    excelValidColsArray.push(startString)
    finalArray.forEach(string=>{
        let colString='';
        string.split(`,`).forEach(num=>{
            colString+=String.fromCharCode(Number.parseInt(num))
        })
        excelValidColsArray.push(colString)
    })
    return excelValidColsArray;
    }
}
  
}).listen(80);
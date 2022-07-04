let _fileName=new WeakMap();
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
let _currentAccountType=new WeakMap();
let _representatorsArray=new WeakMap();
const ExcelJS = require('exceljs');

class ExcellWriterEngine{
    constructor(a,requestorName,requestorEIK,requestorAddress,
        pledgerName,pledgerEIK, pledgerAdress, loanBL,finalInterestRateString,
        finalInterestRateStringOverdue,loanAmount,loanCollateral,representatorsArray,currentAccountType){

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

        const fileName = 'Form.xlsx';
        
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

            wb.xlsx
            .writeFile('newfile25.xlsx')
            .then(() => {
              console.log('file created');
              
              
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
module.exports=ExcellWriterEngine;
let textract = require('textract');
let _requestorName=new WeakMap();
let _requestorEIK=new WeakMap();
let _requestorAddress=new WeakMap();
let _pledgerName=new WeakMap();
let _pledgerEIK=new WeakMap();
let _pledgerAdress=new WeakMap();
let _loanBL=new WeakMap();
let _loanInterestBase=new WeakMap();
let _loanInterestSpread=new WeakMap();
let _loanAmount=new WeakMap();
let _loanCollateral=new WeakMap();
let _roughText=new WeakMap();

;
textract.fromFileWithPath('./Договор за залог на сметка_фирма_БЕЗ блокировка .docx',function(error,text){
    let protoParser=new ParserEngine(text);
    let protoBoundaries=new ParserBoundaries();
    let text1=protoParser.extractTextByBoundaryStrings(protoBoundaries.protoParser.lowerBoundary,protoBoundaries.protoParser.upperBoundary)
    let parser=new ParserEngine(text);
    let boundaries=new ParserBoundaries();
    let requestorName=parser.extractTextByBoundaryStrings(boundaries.requestorName.lowerBoundary,boundaries.requestorName.upperBoundary);
    let requestorEIK=parser.extractTextByBoundaryStrings(boundaries.requestorEIK.lowerBoundary,boundaries.requestorEIK.upperBoundary);
    console.log(requestorName)
    console.log(requestorEIK)

})

class ParserBoundaries{
    constructor(){
        this.protoParser={lowerBoundary:`БАНКАТА`,
        upperBoundary:`II. ПРАВА И ЗАДЪЛЖЕНИЯ  НА ЗАЛОГОДАТЕЛЯ`}
        this.requestorName={lowerBoundary:`от една страна, и`,upperBoundary:`, вписано в Търговския регистър, `}
        this.requestorEIK={lowerBoundary:`ЕИК`,upperBoundary:`със седалище и адрес на управление …………`}
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

}
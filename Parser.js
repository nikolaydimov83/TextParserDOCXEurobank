const textract = require('textract');
const reader = require('xlsx');
let _requestorName=new WeakMap();
let _requestorEIK=new WeakMap();
let _requestorAddress=new WeakMap();
let _pledgerName=new WeakMap();
let _pledgerEIK=new WeakMap();
let _pledgerAdress=new WeakMap();
let _loanBL=new WeakMap();
let _finalInterestRateString=new WeakMap();
let _loanAmount=new WeakMap();
let _loanCollateral=new WeakMap();
let _roughText=new WeakMap();


textract.fromFileWithPath('./Договор за залог на сметка_фирма_БЕЗ блокировка_.docx',function(error,text){
    let protoParser=new ParserEngine(text);
    let protoBoundaries=new ParserBoundaries();
    let text1=protoParser.extractTextByBoundaryStrings(protoBoundaries.protoParser.lowerBoundary,protoBoundaries.protoParser.upperBoundary)
    let parser=new ParserEngine(text1);
    let boundaries=new ParserBoundaries();
    let requestorName=parser.extractTextByBoundaryStrings(boundaries.requestorName.lowerBoundary,boundaries.requestorName.upperBoundary);
    let requestorEIK=parser.extractTextByBoundaryStrings(boundaries.requestorEIK.lowerBoundary,boundaries.requestorEIK.upperBoundary);
    let requestorAddress=parser.extractTextByBoundaryStrings(boundaries.requestorAddress.lowerBoundary,boundaries.requestorAddress.upperBoundary);
    let pledgerName=parser.extractTextByBoundaryStrings(boundaries.pledgerName.lowerBoundary,boundaries.pledgerName.upperBoundary);
    let loanBL=parser.extractTextByBoundaryStrings(boundaries.loanBL.lowerBoundary,boundaries.loanBL.upperBoundary);
    let loanInterestBase=parser.extractTextByBoundaryStrings(boundaries.loanInterestBase.lowerBoundary,boundaries.loanInterestBase.upperBoundary);
    loanInterestBase=loanInterestBase.replace('Бизнес клиенти','БК')
    let loanInterestSpread=parser.extractTextByBoundaryStrings(boundaries.loanInterestSpread.lowerBoundary,boundaries.loanInterestSpread.upperBoundary);
    let loanInterestMinimum=parser.extractTextByBoundaryStrings(boundaries.loanInterestMinimum.lowerBoundary,boundaries.loanInterestMinimum.upperBoundary);
    let finalInterestRateString=`${loanInterestBase} + ${loanInterestSpread}, минимум ${loanInterestMinimum}`
    let loanAmount=parser.extractTextByBoundaryStrings(boundaries.loanAmount.lowerBoundary,boundaries.loanAmount.upperBoundary);
    let loanCollateral=parser.extractTextByBoundaryStrings(boundaries.loanCollateral.lowerBoundary,boundaries.loanAmount.upperBoundary)
    let excellFilename='./Образец 1 Заявление за вписване на договор за залог и за удостоверение.xls'
    let excellWritter=new ExcellWriterEngine(excellFilename,requestorName,requestorEIK,requestorAddress,
    pledgerName,loanBL,finalInterestRateString,loanAmount,loanCollateral)
    

  
    
   

})

class ParserBoundaries{
    constructor(){
        this.protoParser={lowerBoundary:`"БАНКАТА", `,
        upperBoundary:`Чл. 3.`}
        this.requestorName={lowerBoundary:`от една страна, и`,upperBoundary:`, вписано в Търговския регистър, `}
        this.requestorEIK={lowerBoundary:`ЕИК`,upperBoundary:`, със седалище и адрес на управление`}
        this.requestorAddress={lowerBoundary:`със седалище и адрес на управление`,upperBoundary:`, представлявано от`}
        this.pledgerName={lowerBoundary:`(по-долу Договор за кредит") между`,upperBoundary:`в качеството му на Кредитополучател`}
        this.loanBL={lowerBoundary:`е сключен Договор за банков кредит`,upperBoundary:`(по-долу Договор за кредит")`}
        this.loanInterestBase={lowerBoundary:`бизнес клиенти на Юробанк България" АД`,upperBoundary:`плюс фиксирана договорна надбавка в размер на`}
        this.loanInterestSpread={lowerBoundary:`плюс фиксирана договорна надбавка в размер на`,upperBoundary:`процентни пункта, но не по ниска от`}
        this.loanInterestMinimum={lowerBoundary:`но не по ниска от`,upperBoundary:`/ ( Долна граница на дължимата лихва"`}
        this.loanAmount={lowerBoundary:`(максимален разрешен размер на кредита):`,upperBoundary:`. Срок за издължаване:`}
        this.loanCollateral={lowerBoundary:`Предмет на настоящия договор е учредяването на`,upperBoundary:`/по-долу "Договора за сметка"/`}
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
    constructor(filename,requestorName,requestorEIK,requestorAddress,
        pledgerName,loanBL,finalInterestRateString,
        loanAmount,loanCollateral){

        this.file = reader.readFile(filename);
        _requestorName.set(this,requestorName);
        _requestorEIK.set(this,requestorEIK);
        _requestorAddress.set(this,requestorAddress);
        _pledgerName.set(this,pledgerName);
        _loanBL.set(this,loanBL);
        _finalInterestRateString.set(this,finalInterestRateString);
        _loanAmount.set(this,loanAmount)
        _loanCollateral.set(this,loanCollateral);    
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
        return _requestorAddress.get(this)
    }

    get loanBL(){
        return _loanBL.get(this)
    }

    get finalInterestRateString(){
        return _finalInterestRateString.get(this)
    }

    get loanAmount(){
        return _loanAmount.get(this)
    }
    get loanCollateral(){
        return _loanCollateral.get(this)
    }

}
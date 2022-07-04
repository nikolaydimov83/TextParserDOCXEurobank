
let _roughText=new WeakMap();
class ParserEngine{
    constructor(roughText){
        _roughText.set(this,roughText);
    }
    get roughText(){
        return _roughText.get(this);
    }
    extractTextByBoundaryStrings(lowerBoundary,upperBoundary){
        let intermediateArrayOfStrings=this.roughText.split(lowerBoundary)
        try{
            let finalArray=intermediateArrayOfStrings[1].split(upperBoundary);
            return finalArray[0];
        }catch{
            return `No Info for pledger EIK added`
        }
          
    }
    }
    module.exports=ParserEngine;
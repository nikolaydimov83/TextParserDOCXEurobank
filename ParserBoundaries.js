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

module.exports=ParserBoundaries;
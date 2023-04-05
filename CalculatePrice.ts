import {Data} from "./Data";

namespace SetUnitPrice
{
    const fs = require('fs');
    const AdmZip = require('adm-zip');
    const Big = require('big.js');

    export class CalculatePrice
    {
        private static docBID : JSON;
        private static eleBID : JSON;
        private static maxBID : JSON;
        public static myPercent;
        public static balancedUnitPriceRate;
        public static targetRate;
        private static exSum = 0;
        private static exCount = 0;

        public static Calculation() : void
        {
            const bidString : string = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", 'utf-8');
            this.docBID = JSON.parse(bidString);
            this.eleBID = this.docBID['data'];

            // this.ApplyStandardPriceOption();

            BidHandling.JsonToBid();

            // this.myPercent = Big(0.85)

            // this.CalculateRate(-1, 2.9932);

            // this.SetBusinessInfo();

            //this.CreateFile();
        }

        public static Reset() : void
        {

        }

        public static ApplyStandardPriceOption() : void
        {
            const bidT3 : object = this.eleBID['T3'];
            let code : string;
            let type : string;

            for(let key in bidT3){
                code = JSON.stringify(bidT3[key]['C9']['_text']);
                type = JSON.stringify(bidT3[key]['C5']['_text'])[1];

                if(code !== undefined && type === 'S')
                {
                    let constNum : string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                    let numVal : string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                    let detailVal : string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                    // let curObject = Data.Dic.get(constNum);



                    console.log(constNum + numVal + detailVal);
                }
            }
        }

        public static GetFixedPriceRate() : void
        {
            // const directConstPrice = Data.Investigation['직공비'];
            // const fixCostSum = Data.InvestigateFixedPriceDirectMaterial + Data.InvestigateFixedPriceDirectLabor + Data.InvestigateFixedPriceOutputExpense;

            // Data.FixedPricePercent = Math.trunc(((fixCostSum / directConstPrice) * 100) * 10000) / 10000;
        }

        public static FindMyPercent() : void
        {

        }

        public static GetWeight() : void
        {

        }

        public static CalculateRate(presonalRate : any, balancedRate : any) : void
        {
            const unitPrice = Big(100);
            presonalRate = Big(presonalRate);
            balancedRate = Big(balancedRate);

            // this.balancedUnitPriceRate = unitPrice.times(0.9).times(balancedRate.div(100).plus(1.0).times(this.myPercent)).div(this.myPercent.times(0.1).neg().plus(1.0)).div(100);
            // this.targetRate = unitPrice.times(presonalRate.div(100).add(1.0)).times(0.9).plus(unitPrice.times(this.balancedUnitPriceRate).times(1.0)).times(this.myPercent).div(100);
            //this.targetRate =  ;

            this.balancedUnitPriceRate = ((0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) / (1.0 - 0.1 * this.myPercent)) / 100;
            this.targetRate = ((unitPrice * (1.0 + presonalRate / 100) * 0.9 + unitPrice * this.balancedUnitPriceRate * 0.1) * this.myPercent) / 100;
            this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000;

            console.log(this.targetRate);
        }

        public static RoundOrTruncate(Rate : any, Object : object, myMaterialUnit : any, myLaborUnit : any, myExpenseUnit : any) : void
        {
            // myMaterialUnit = Math.trunc(Object.MaterialUnit * Rate * 10) / 10;
        }

        public static Recalculation() : void
        {

        }

        public static SetExcludingPrice() : void
        {

        }

        public static GetAdjustedExcludePrice() : void
        {

        }

        public static SetPriceOfSuperConstruction() : void
        {

        }

        public static SetBusinessInfo() : void
        {
            const bidT1 : object = this.eleBID['T1'];

            // bidT1['C17']['_text'] = '0000';
            // bidT1['C18']['_text'] = 'test';

            fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", JSON.stringify(this.docBID));
        }

        public static SubstitutePrice() : void
        {

        }

        public static CreateZipFile(xlsfiles : Array<string>) : void
        {
            if (fs.existsSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip"))
            {
                fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip");
            }

            const zip = new AdmZip();

            for(let idx in xlsfiles)
            {
                zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\" + xlsfiles[idx]);
            }

            zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\WORK DIRECTORY" + "\\입찰내역.zip");
        }

        public static CreateFile() : void
        {
            // CreateResultFile.Create();
            const files : Array<string> = fs.readdirSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID");
            const xlsFiles : Array<string> = files.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === '.xls');
            this.CreateZipFile(xlsFiles);
        }
    }

}
SetUnitPrice.CalculatePrice.Calculation();
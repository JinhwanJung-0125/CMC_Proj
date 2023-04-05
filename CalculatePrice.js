var SetUnitPrice;
(function (SetUnitPrice) {
    var fs = require('fs');
    var AdmZip = require('adm-zip');
    var Big = require('big.js');
    var CalculatePrice = /** @class */ (function () {
        function CalculatePrice() {
        }
        CalculatePrice.Calculation = function () {
            var bidString = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", 'utf-8');
            this.docBID = JSON.parse(bidString);
            this.eleBID = this.docBID['data'];
            // this.ApplyStandardPriceOption();
            BidHandling.JsonToBid();
            // this.myPercent = Big(0.85)
            // this.CalculateRate(-1, 2.9932);
            // this.SetBusinessInfo();
            //this.CreateFile();
        };
        CalculatePrice.Reset = function () {
        };
        CalculatePrice.ApplyStandardPriceOption = function () {
            var bidT3 = this.eleBID['T3'];
            var code;
            var type;
            for (var key in bidT3) {
                code = JSON.stringify(bidT3[key]['C9']['_text']);
                type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
                if (code !== undefined && type === 'S') {
                    var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                    var numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                    var detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                    // let curObject = Data.Dic.get(constNum);
                    console.log(constNum + numVal + detailVal);
                }
            }
        };
        CalculatePrice.GetFixedPriceRate = function () {
            // const directConstPrice = Data.Investigation['직공비'];
            // const fixCostSum = Data.InvestigateFixedPriceDirectMaterial + Data.InvestigateFixedPriceDirectLabor + Data.InvestigateFixedPriceOutputExpense;
            // Data.FixedPricePercent = Math.trunc(((fixCostSum / directConstPrice) * 100) * 10000) / 10000;
        };
        CalculatePrice.FindMyPercent = function () {
        };
        CalculatePrice.GetWeight = function () {
        };
        CalculatePrice.CalculateRate = function (presonalRate, balancedRate) {
            var unitPrice = Big(100);
            presonalRate = Big(presonalRate);
            balancedRate = Big(balancedRate);
            // this.balancedUnitPriceRate = unitPrice.times(0.9).times(balancedRate.div(100).plus(1.0).times(this.myPercent)).div(this.myPercent.times(0.1).neg().plus(1.0)).div(100);
            // this.targetRate = unitPrice.times(presonalRate.div(100).add(1.0)).times(0.9).plus(unitPrice.times(this.balancedUnitPriceRate).times(1.0)).times(this.myPercent).div(100);
            //this.targetRate =  ;
            this.balancedUnitPriceRate = ((0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) / (1.0 - 0.1 * this.myPercent)) / 100;
            this.targetRate = ((unitPrice * (1.0 + presonalRate / 100) * 0.9 + unitPrice * this.balancedUnitPriceRate * 0.1) * this.myPercent) / 100;
            this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000;
            console.log(this.targetRate);
        };
        CalculatePrice.RoundOrTruncate = function (Rate, Object, myMaterialUnit, myLaborUnit, myExpenseUnit) {
            // myMaterialUnit = Math.trunc(Object.MaterialUnit * Rate * 10) / 10;
        };
        CalculatePrice.Recalculation = function () {
        };
        CalculatePrice.SetExcludingPrice = function () {
        };
        CalculatePrice.GetAdjustedExcludePrice = function () {
        };
        CalculatePrice.SetPriceOfSuperConstruction = function () {
        };
        CalculatePrice.SetBusinessInfo = function () {
            var bidT1 = this.eleBID['T1'];
            // bidT1['C17']['_text'] = '0000';
            // bidT1['C18']['_text'] = 'test';
            fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", JSON.stringify(this.docBID));
        };
        CalculatePrice.SubstitutePrice = function () {
        };
        CalculatePrice.CreateZipFile = function (xlsfiles) {
            if (fs.existsSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip")) {
                fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip");
            }
            var zip = new AdmZip();
            for (var idx in xlsfiles) {
                zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\" + xlsfiles[idx]);
            }
            zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\WORK DIRECTORY" + "\\입찰내역.zip");
        };
        CalculatePrice.CreateFile = function () {
            // CreateResultFile.Create();
            var files = fs.readdirSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID");
            var xlsFiles = files.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.xls'; });
            this.CreateZipFile(xlsFiles);
        };
        CalculatePrice.exSum = 0;
        CalculatePrice.exCount = 0;
        return CalculatePrice;
    }());
    SetUnitPrice.CalculatePrice = CalculatePrice;
})(SetUnitPrice || (SetUnitPrice = {}));
SetUnitPrice.CalculatePrice.Calculation();

import { Data } from "./Data";
import { FillCostAccount } from "./FillCostAccount"

const fs = require('fs');
const AdmZip = require('adm-zip');
const Big = require('big.js');

export class CalculatePrice {
    private static docBID : JSON;
    private static eleBID : JSON;
    private static maxBID : JSON;
    public static myPercent : number;
    public static balancedUnitPriceRate : number;
    public static targetRate : number;
    private static exSum  : number = 0;
    private static exCount : number = 0;

    public static Calculation(): void {
        const bidString: string = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", 'utf-8');
        this.docBID = JSON.parse(bidString);
        this.eleBID = this.docBID['data'];

        this.ApplyStandardPriceOption();

        // this.myPercent = Big(0.85)

        // this.CalculateRate(-1, 2.9932);

        // this.SetBusinessInfo();

        //this.CreateFile();
    }

    public static Reset(): void {

    }

    public static ApplyStandardPriceOption(): void {
        const bidT3: object = this.eleBID['T3'];
        let code: string;
        let type: string;

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                let curObject = Data.Dic.get(constNum).find(x => x.WorkNum === numVal && x.DetailWorkNum === detailVal);

                if (curObject.Item.localeCompare('표준시장단가')) {
                    Data.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data.FixedPriceDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.FixedPriceDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.FixedPriceOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data.StandardMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.StandardLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.StandardExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);

                    if (curObject.MaterialUnit !== 0)
                        curObject.MaterialUnit = (Math.trunc(curObject.MaterialUnit * 0.997 * 10) / 10) + 1.0;
                    if (curObject.LaborUnit !== 0)
                        curObject.LaborUnit = (Math.trunc(curObject.LaborUnit * 0.997 * 10) / 10) + 0.1;
                    if (curObject.ExpenseUnit !== 0)
                        curObject.ExpenseUnit = (Math.trunc(curObject.ExpenseUnit * 0.997 * 10) / 10) + 0.1;

                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                    bidT3[key]['C20']['_text'] = curObject.Material.toString();
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();

                    Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data.FixedPriceDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.FixedPriceDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.FixedPriceOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data.StandardMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.StandardLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.StandardExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                }
            }
        }
    }

    public static GetFixedPriceRate(): void {
        const directConstPrice : number = Data.Investigation['직공비'];
        const fixCostSum : number = Data.InvestigateFixedPriceDirectMaterial + Data.InvestigateFixedPriceDirectLabor + Data.InvestigateFixedPriceOutputExpense;

        Data.FixedPricePercent = Math.trunc(((fixCostSum / directConstPrice) * 100) * 10000) / 10000;
    }

    public static FindMyPercent(): void {
        if (Data.FixedPricePercent < 20.0)
            this.myPercent = 0.85;
        else if (Data.FixedPricePercent < 25.0)
            this.myPercent = 0.84;
        else if (Data.FixedPricePercent < 30.0)
            this.myPercent = 0.83;
        else this.myPercent = 0.82;
    }

    public static GetWeight(): void {
        let varCostSum = Data.RealPriceDirectMaterial + Data.RealPriceDirectLabor + Data.RealPriceOutputExpense;
        let weight: number;
        let maxWeight: number = 0;
        let weightSum: number = 0;
        let max : Data = new Data();

        Data.Dic.forEach((value, _) => {
            for(let idx in value)
            {
                if(value[idx].Item.localeCompare('일반'))
                {
                    let material = value[idx].Material;
                    let labor = value[idx].Labor;
                    let expense = value[idx].Expense;

                    weight = Math.round(((material + labor + expense) / varCostSum) * 1000000) / 1000000;
                    weightSum += weight;

                    if(maxWeight < weight){
                        maxWeight = weight;
                        max = value[idx];
                    }

                    value[idx].Weight = weight;
                }
            }
        })

        if(weightSum !== 1.0){
            let lack : number = 1.0 - weightSum;
            max.Weight += lack;
        }
    }

    public static CalculateRate(presonalRate: any, balancedRate: any): void {
        const unitPrice = Big(100);
        presonalRate = Big(presonalRate);
        balancedRate = Big(balancedRate);

        this.balancedUnitPriceRate = ((0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) / (1.0 - 0.1 * this.myPercent)) / 100;
        this.targetRate = ((unitPrice * (1.0 + presonalRate / 100) * 0.9 + unitPrice * this.balancedUnitPriceRate * 0.1) * this.myPercent) / 100;
        this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000;

        console.log(this.targetRate);
    }

    public static RoundOrTruncate(Rate: number, Object: Data, refMyMaterialUnit: {value: number}, refMyLaborUnit: {value: number}, refMyExpenseUnit: {value: number}): void {
        if(Data.UnitPriceTrimming.localeCompare('1')){
            refMyMaterialUnit.value = Math.trunc(Object.MaterialUnit * Rate * 10) / 10;
            refMyLaborUnit.value = Math.trunc(Object.LaborUnit * Rate * 10) / 10;
            refMyExpenseUnit.value = Math.trunc(Object.ExpenseUnit * Rate * 10) /10;
        }
        else if(Data.UnitPriceTrimming.localeCompare('2')){
            refMyMaterialUnit.value = Math.ceil(Object.MaterialUnit * Rate);
            refMyLaborUnit.value = Math.ceil(Object.LaborUnit * Rate);
            refMyExpenseUnit.value = Math.ceil(Object.ExpenseUnit * Rate);
        }
    }

    public static CheckLaborLimit80(Object: Data, refMyMaterialUnit: {value: number}, refMyLaborUnit: {value: number}, refMyExpenseUnit: {value: number}): void{
        if(Object.LaborUnit * 0.8 > refMyLaborUnit.value){
            const deficiency: number = Object.LaborUnit * 0.8 - refMyLaborUnit.value;

            if(refMyLaborUnit.value !== 0)
                refMyMaterialUnit.value -= deficiency;
            else if(refMyExpenseUnit.value !== 0)
                refMyExpenseUnit.value -= deficiency;

            refMyLaborUnit.value = Object.LaborUnit * 0.8;
        }
    }

    public static Recalculation(): void {
        const bidT3: object = this.eleBID['T3'];
        let code: string;
        let type: string;
        this.exCount = 0;
        this.exSum = 0;

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                let curObject = Data.Dic.get(constNum).find(x => x.WorkNum === numVal && x.DetailWorkNum === detailVal);

                if (curObject.Item.localeCompare('일반')) {
                    Data.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);

                    let targetPrice: number = (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) * this.targetRate;

                    let myMaterialUnit = {value : 0};
                    let myLaborUnit = {value: 0};
                    let myExpenseUnit = {value: 0};
                    let myPrice: number;

                    if(Data.ZeroWeightDeduction.localeCompare('1')){
                        if(curObject.Weight === 0 && curObject.LaborUnit === 0){
                            curObject.MaterialUnit = Math.ceil(curObject.MaterialUnit * 0.5);
                            curObject.ExpenseUnit = Math.ceil(curObject.ExpenseUnit * 0.5);

                            bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                            bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                            bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                            bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                            bidT3[key]['C20']['_text'] = curObject.Material.toString();
                            bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                            bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                            bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();

                            Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                            Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                            Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);

                            continue;
                        }
                        else{
                            this.RoundOrTruncate(this.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                            this.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        }
                    }
                    else if(Data.ZeroWeightDeduction.localeCompare('2')){
                        this.RoundOrTruncate(this.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        this.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                    }

                    myPrice = myMaterialUnit.value + myLaborUnit.value + myExpenseUnit.value;

                    if(Data.LaborCostLowBound.localeCompare('1')){
                        let Excess: number = myPrice - targetPrice;
                        let laborExcess: number = myLaborUnit.value - curObject.LaborUnit * 0.8;
                        laborExcess = Math.trunc(laborExcess * 10) / 10;

                        if(laborExcess > 0){
                            if(myExpenseUnit.value !== 0){
                                myLaborUnit.value -= laborExcess;
                                myExpenseUnit.value += laborExcess + Excess;
                            }
                            else{
                                if(myMaterialUnit.value !== 0){
                                    myLaborUnit.value -= laborExcess;
                                    myMaterialUnit.value += laborExcess + Excess;
                                }
                                else{
                                    myLaborUnit.value -= laborExcess;
                                    myExpenseUnit.value += laborExcess + Excess;
                                }
                            }
                        }
                        else if(laborExcess < 0){
                            myLaborUnit.value = curObject.LaborUnit * 0.8;

                            if(myMaterialUnit.value !== 0){
                                myMaterialUnit.value += laborExcess + Excess;
                            }
                            else{
                                myExpenseUnit.value += laborExcess + Excess;
                            }
                        }
                    }

                    curObject.MaterialUnit = myMaterialUnit.value;
                    curObject.LaborUnit = myLaborUnit.value;
                    curObject.ExpenseUnit = myExpenseUnit.value;

                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                    bidT3[key]['C20']['_text'] = curObject.Material.toString();
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();

                    Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                }
                else if(curObject.Item === '제요율적용제외'){
                    curObject.MaterialUnit = Math.trunc(curObject.MaterialUnit * this.targetRate * 10) / 10;
                    curObject.LaborUnit = Math.trunc(curObject.LaborUnit * this.targetRate * 10) / 10;
                    curObject.ExpenseUnit = Math.trunc(curObject.ExpenseUnit * this.targetRate * 10) / 10;

                    this.exSum += curObject.PriceSum;
                    this.exCount++;
                }
            }
        }

    }

    public static SetExcludingPrice(): void {
        let TempInvestDirectSum: number = Data.Investigation['직공비'];
        let TempRealDirectSum: number = FillCostAccount.ToLong(Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense); 
    }

    public static GetAdjustedExcludePrice(): void {

    }

    public static SetPriceOfSuperConstruction(): void {

    }

    public static SetBusinessInfo(): void {
        const bidT1: object = this.eleBID['T1'];

        // bidT1['C17']['_text'] = '0000';
        // bidT1['C18']['_text'] = 'test';

        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", JSON.stringify(this.docBID));
    }

    public static SubstitutePrice(): void {

    }

    public static CreateZipFile(xlsfiles: Array<string>): void {
        if (fs.existsSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip")) {
            fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\WORK DIRCETORY\\입찰내역.zip");
        }

        const zip = new AdmZip();

        for (let idx in xlsfiles) {
            zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\" + xlsfiles[idx]);
        }

        zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\WORK DIRECTORY" + "\\입찰내역.zip");
    }

    public static CreateFile(): void {
        // CreateResultFile.Create();
        const files: Array<string> = fs.readdirSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID");
        const xlsFiles: Array<string> = files.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === '.xls');
        this.CreateZipFile(xlsFiles);
    }
}


CalculatePrice.Calculation();
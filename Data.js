"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Data = void 0;
var path = require('path');
var fs = require('fs');
var Data = exports.Data = /** @class */ (function () {
    function Data() {
        this.materialUnit = 0; //재료비 단가
        this.laborUnit = 0; //노무비 단가
        this.expenseUnit = 0; //경비 단가
        this.Item = ''; //항목 구분(공종(입력불가), 무대(입력불가), 일반, 관급, PS, 제요율적용제외, 안전관리비, PS내역, 표준시장단가)
        this.ConstructionNum = ''; //공사 인덱스
        this.WorkNum = ''; //세부 공사 인덱스
        this.DetailWorkNum = ''; //세부 공종 인덱스
        this.Code = ''; //코드
        this.Name = ''; //품명
        this.Standard = ''; //규격
        this.Unit = ''; //단위
        this.Quantity = 0; //수량
        this.Weight = 0; //가중치
        this.PriceScore = 0; //세부 점수
    }
    Object.defineProperty(Data.prototype, "MaterialUnit", {
        //재료비 단가
        get: function () {
            //사용자가 단가 정수처리를 원한다면("2") 정수 값으로 return / Reset 함수를 쓰지 않은 경우의 조건 추가 (23.02.06)
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Math.ceil(this.materialUnit);
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                // 사용자가 단가 소수점 처리를 원하거나 Reset 함수를 썼다면 소수 첫째 자리 아래로 절사 (23.02.06)
                return Math.floor(this.materialUnit * 10) / 10;
            return this.materialUnit; //Default는 있는 그대로의 값을 return
        },
        set: function (value) {
            this.materialUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "LaborUnit", {
        //노무비 단가
        get: function () {
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Math.ceil(this.laborUnit);
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                return Math.floor(this.laborUnit * 10) / 10;
            return this.laborUnit;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "LavorUnit", {
        set: function (value) {
            this.laborUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "ExpenseUnit", {
        //경비 단가
        get: function () {
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Math.ceil(this.expenseUnit);
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                return Math.floor(this.expenseUnit * 10) / 10;
            return this.expenseUnit;
        },
        set: function (value) {
            this.expenseUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Material", {
        get: function () {
            return Math.floor(this.Quantity * this.MaterialUnit);
        } //재료비 (수량 x 단가)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Labor", {
        get: function () {
            return Math.floor(this.Quantity * this.LaborUnit);
        } //노무비
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Expense", {
        get: function () {
            return Math.floor(this.Quantity * this.ExpenseUnit);
        } //경비
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "UnitPriceSum", {
        get: function () {
            return this.MaterialUnit + this.LaborUnit + this.ExpenseUnit;
        } //합계단가
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "PriceSum", {
        get: function () {
            return this.Material + this.Labor + this.Expense;
        } //합계(세부공종별 금액의 합계)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Score", {
        get: function () {
            return this.PriceScore * this.Weight;
        } //단가 점수(세부 점수 * 가중치)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data, "BalancedRate", {
        //업체 평균 예측율
        get: function () {
            //decimal로 반환. 소수 계산 라이브러리 필요
            return Data.BalanceRateNum; //입력받은 BalancedRateNum(double? 형)을 decimal로 바꿈
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data, "PersonalRate", {
        //내 예가 사정률
        get: function () {
            //decimal로 반환. 소수 계산 라이브러리 필요
            return Data.PersonalRateNum; //입력받은 PersonalRateNum(double? 형)을 decimal로 바꿈
        },
        enumerable: false,
        configurable: true
    });
    Data.folder = path.resolve(__dirname, '../AutoBID'); //내 문서 폴더의 AutoBID 폴더로 지정 (23.02.02)
    // WPF 앱 파일 관리 변수
    Data.XlsText = '';
    Data.CanCovertFile = false; // 새로운 파일 업로드 시 변환 가능
    Data.IsConvert = false; // 변환을 했는지 안했는지
    Data.IsBidFileOk = true; // 정상적인 공내역 파일인지
    Data.IsFileMatch = true; // 공내역 파일과 실내역 파일의 공사가 일치하는지
    Data.CompanyRegistrationNum = ''; //1.31 사업자등록번호 추가
    Data.CompanyRegistrationName = ''; // 2.02 회사명 추가
    // 프로그램 폴더로 위치 변경
    Data.work_path = path.join(Data.folder, 'WORK DIRECTORY'); //작업폴더(WORK DIRECTORY) 경로
    Data.Dic = new Map(); //key : 세부공사별 번호 / value : 세부공사별 리스트
    Data.ConstructionNums = new Map(); //세부 공사별 번호 저장
    Data.MatchedConstNum = new Map(); //실내역과 세부공사별 번호의 매칭 결과
    Data.Fixed = new Map(); //고정금액 항목별 금액 저장
    Data.Rate1 = new Map(); //적용비율1 저장
    Data.Rate2 = new Map(); //적용비율2 저장
    Data.RealPrices = new Map(); //원가계산서 항목별 금액 저장
    Data.Investigation = new Map(); //세부결과_원가계산서 항목별 조사금액 저장
    Data.Bidding = new Map(); //세부결과_원가계산서 항목별 입찰금액 저장
    Data.Correction = new Map(); //원가계산서 조사금액 보정 항목 저장
    //사용자의 옵션 및 사정률 데이터
    Data.UnitPriceTrimming = '0'; //단가 소수 처리 (defalut = "0")
    Data.StandardMarketDeduction = '2'; //표준시장단가 99.7% 적용
    Data.ZeroWeightDeduction = '2'; //가중치 0% 공종 50% 적용
    Data.CostAccountDeduction = '2'; //원가계산 제경비 99.7% 적용
    Data.BidPriceRaise = '2'; //투찰금액 천원 절상
    Data.LaborCostLowBound = '2'; //노무비 하한 80%
    Data.ExecuteReset = '0'; //Reset 함수 사용시 단가 소수처리 옵션과 별개로 소수 첫째자리 아래로 절사
    return Data;
}());

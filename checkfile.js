const XLSX = require("xlsx");

let workbook, sheetNameList, sheet1, sheet2, sheet3, sheet4;
let sheets = [];

let osBeginKeyChangesMap = {
    CLASS: "insuranceClass",
    "POLICY NO": "policyNo",
    "CLAIM NO": "claimNo",
    "INSURED NAME": "insuredName",
    "DATE REPORTED": "dateReported",
    "DATE OF LOSS": "dateOfLoss",
    "PERIOD FROM": "periodFrom",
    "PERIOD TO": "periodTo",
    ESTIMATE: "osBeginEstimate"
};

let osEndMonthKeyChangesMap = {
    CLASS: "insuranceClass",
    "POLICY NO": "policyNo",
    "CLAIM NO": "claimNo",
    "INSURED NAME": "insuredName",
    "DATE REPORTED": "dateReported",
    "DATE OF LOSS": "dateOfLoss",
    "PERIOD FROM": "periodFrom",
    "PERIOD TO": "periodTo",
    ESTIMATE: "osEndMonthEstimate"
};

let intimatedKeyChangesMap = {
    CLASS: "insuranceClass",
    INSURED: "insured",
    AGENCY: "Agency",
    "POLICY NO": "policyNo",
    "CLAIM NO": "claimNo",
    "INTIMATION RESERVE": "intimationReserve",
    "DATE OF LOSS": "dateOfLoss",
    "DATE REPORTED": "dateReported"
};

let paymentKeyChangesMap = {
    CLASS: "insuranceClass",
    "DATE OF CHEQUE": "dateOfcheque",
    "CHEQUE NO": "chequeNo",
    "CLAIM NO": "claimNo",
    "POLICY HOLDER": "policyHolder",
    "UW YEAR": "uwYear",
    PAYEE: "payee",
    "PAID AMOUNT": "paidAmount"
};


passFileNameForLoading = function (file) {

    workbook = XLSX.readFile(file.toString());
    sheetNameList = workbook.SheetNames;
    sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
    sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
    sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[2]]);
    sheet4 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[3]]);
    sheets = [sheet1, sheet2, sheet3, sheet4];

    return checkSheetFields();

};

// check files
checkSheetFields = function () {

    let status = [];

    if (sheets.length < 4) {
        status.push({
            sheet: "no sheets",
            status: 'error',
            message: "one or more sheets are missing"
        });
    } else {

        let sheetName = "";
        sheets.map(function (sheet, index) {
            let theIndex = index + 1;
            let keysMap;

            if (theIndex == 1) {
                sheetName = "Begining OS estimate";
                keysMap = osBeginKeyChangesMap;
            } else if (theIndex == 2) {
                sheetName = "Intimated claims";
                keysMap = intimatedKeyChangesMap;
            } else if (theIndex == 3) {
                sheetName = "Paid claims";
                keysMap = paymentKeyChangesMap;
            } else {
                sheetName = "End month OS estimate";
                keysMap = osEndMonthKeyChangesMap;
            }

            for (const prop in sheet[0]) {
                if (!keysMap.hasOwnProperty(prop)) {
                    status.push({
                        sheet: sheetName,
                        status: 'error',
                        message: `column name is mispelled or missing *${prop}*`
                    });
                }
            }
        });
    }

    if (status.length < 1) {
        status.push({
            status: 'success',
            sheet: "all",
            message: "all sheets are OKAY"
        });
    }


    return status;
};



module.exports = {

    passFileNameForLoading: passFileNameForLoading

};
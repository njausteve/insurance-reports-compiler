const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");
const _ = require("lodash");
const del = require("del");
const path = require('path');
const objectRenameKeys = require("object-rename-keys");




let workbook, sheetNameList,  sheet1, sheet2, sheet3, sheet4;
let sheets = [];



exports.passFileNameForLoading = function(file){



workbook = XLSX.readFile(file.toString());

sheetNameList = workbook.SheetNames;

sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[2]]);
sheet4 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[3]]);

sheets = [sheet1, sheet2, sheet3, sheet4];


   return checkSheetFields();


};




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

let printToExcelKeysmap = {
  insuranceClass: "CLASS",
  policyNo: "POLICY NO",
  claimNo: "CLAIM NO",
  insuredName: "INSURED NAME",
  dateReported: "DATE REPORTED",
  dateOfLoss: "DATE OF LOSS",
  periodFrom: "PERIOD FROM",
  periodTo: "PERIOD TO",
  osBeginEstimate: "BEGINING OS ESTIMATE",
  osEndMonthEstimate: "END OS ESTIMATE",
  insured: "INSURED",
  Agency: "AGENCY",
  intimationReserve: "INTIMATION RESERVE",
  dateOfcheque: "DATE OF CHEQUE",
  chequeNo: "CHEQUE NO",
  policyHolder: "POLICY HOLDER",
  uwYear: "UW YEAR",
  payee: "PAYEE",
  paidAmount: "PAID AMOUNT",
  difference: "DIFFERENCE",
  revived: "REVIVED",
  movement: "MOVED",
  noClaim: "NO CLAIM",
  settled: "SETTTLED"
};

// check sheet Fields

checkSheetFields = function() {

  let status = [];

  if (sheets.length < 4) {
    status.push({
      sheet: "no sheets",
      error: "one or more sheets are missing"
    });
  } else {
    let sheetName = "";
    sheets.map(function(sheet, index) {
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
            error: `column name is mispelled or missing *${prop}*`
          });
        }
      }
    });
  }

  if (status.length < 1) {
    status.push({ sheet: "all", message: "all sheets OKAY" });
  }

 
  return status;
};

// console.log(checkSheetFields());


var XLSX = require('xlsx');
var workbook = XLSX.readFile('./os as at 31032018.xlsx');
var sheet_name_list = workbook.SheetNames;
var fs = require('fs');
var objectRenameKeys = require('object-rename-keys');

//console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]))
let combineOS =[];
let insuranceItem = {};
var sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
var sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]]);
var sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[2]]);

osBeginKeyChangesMap = {
       'CLASS' : 'insuranceClass',
       'POLICY NO':'policyNo',
       'CLAIM NO': 'claimNo',
       'INSURED NAME':'insuredName' ,
       'DATE REPORTED': 'dateReported',
       'D.O.L': 'dateOfLoss'  ,
       'DATE OF LOSS': 'dateOfLoss',
       'PERIOD FROM':'periodFrom',
       'PERIOD TO': 'periodTo',
       'ESTIMATE': 'osBeginEstimate',
       ' O/S ESTIMATE ':'osBeginEstimate',
       'MANDATORY CLAIM': 'osBeginMandatoryClaim',
       'COMPANY CLAIM':'osBeginCompanyClaim',        
};

osEndMonthKeyChangesMap = {
  'CLASS' : 'insuranceClass',
  'POLICY NO':'policyNo',
  'CLAIM NO': 'claimNo',
  'INSURED NAME':'insuredName' ,
  'DATE REPORTED': 'dateReported',
  'D.O.L': 'dateOfLoss'  ,
  'DATE OF LOSS': 'dateOfLoss',
  'PERIOD FROM':'periodFrom',
  'PERIOD TO': 'periodTo',
  'ESTIMATE': 'osEndMonthEstimate',
  ' O/S ESTIMATE ':'osEndMonthEstimate',
  'MANDATORY CLAIM': 'osEndMonthMandatoryClaim',
  'COMPANY CLAIM':'osEndMonthCompanyClaim',        
};

var osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
var osEndMonth = objectRenameKeys(sheet2, osEndMonthKeyChangesMap);
//var intimated = objectRenameKeys(sheet3, osKeyChangesMap);


console.log(osEndMonth);





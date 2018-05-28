const XLSX = require("xlsx");
let workbook = XLSX.readFile("./source file/insurance.xlsx");
let sheetNameList = workbook.SheetNames;
const _ = require("lodash");
const del = require("del");
const objectRenameKeys = require("object-rename-keys");

//console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]))
let sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
let sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
let sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[2]]);

let osBeginKeyChangesMap = {
    CLASS: "insuranceClass",
    "POLICY NO": "policyNo",
    "CLAIM NO": "claimNo",
    "INSURED NAME": "insuredName",
    "DATE REPORTED": "dateReported",
    "D.O.L": "dateOfLoss",
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
    "D.O.L": "dateOfLoss",
    "DATE OF LOSS": "dateOfLoss",
    "PERIOD FROM": "periodFrom",
    "PERIOD TO": "periodTo",
    ESTIMATE: "osEndMonthEstimate"
};


let intimatedKeyChangesMap = {
    CLASS: "insuranceClass",
    "INSURED" : "insured",
    "AGENCY" : "Agency",
    "POLICY NO": "policyNo",
    "CLAIM NO": "claimNo",
    "INTIMATION RESERVE": "intimationReserve"
};


let osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
let osEndMonth = objectRenameKeys(sheet2, osEndMonthKeyChangesMap);
let intimated = objectRenameKeys(sheet3, intimatedKeyChangesMap);
// //let intimated = objectRenameKeys(sheet3, osKeyChangesMap);

let combinedOsWithDuplicates = _.concat(osEndMonth, osBeginMonth, "claimNo");

// [Removed OS] : to find those in the osEndMonth but not in the osBeginMonth
let addedOsEndFromBeginMonth = _.differenceBy(
    osEndMonth,
    osBeginMonth,
    "claimNo"
);

// [Added OS]: to find those in the osBeginMonth but not in the osEndMonth
let removedOsBeginToEndMonth = _.differenceBy(
    osBeginMonth,
    osEndMonth,
    "claimNo"
);

let osNoChange = _.intersectionBy(osBeginMonth, osEndMonth, "claimNo");

// console.log( "unionby ---> : " + combinedOs.length + "\n concat -----> :", combined.length + "\n added ----> : " + addedOsEndFromBeginMonth.length + "\n removed ---> : " + removedOsBeginToEndMonth.length + "\n no change---> : " + osNoChange.length );

let osRepeatedClaimNo = osNoChange.map(function (claim) {
    let newObj = {};

    combinedOsWithDuplicates.map(function (combineClaim) {
        if (combineClaim.claimNo === claim.claimNo) {
            for (const prop in combineClaim) {
                newObj[prop] = combineClaim[prop];
            }
        }
    });

    return newObj;
});


// combined sheet OS begining and end no dubplictes:

let combinedSheets12 = _.concat(
    addedOsEndFromBeginMonth,
    removedOsBeginToEndMonth,
    osRepeatedClaimNo,
    "claimNO"
).map(function (claim) {
    let newObj = {};

    for (const prop in claim) {

        newObj[prop] = claim[prop];

        if (claim.osEndMonthEstimate == null) {
            newObj.osEndMonthEstimate = 0;

        }else if(claim.osBeginEstimate == null) {
            newObj.osBeginEstimate = 0;
        }
    }

    return newObj;
});




// combined  (OS end & OS begining) and intimated









// find values that are revived : present in Added but not intimated for this month

let revivedClaims = _.differenceBy(addedOsEndFromBeginMonth, intimated, "claimNo");




// create work book
let wb = XLSX.utils.book_new();

// create sheetsNames
wb.SheetNames.push("Combined OS");
wb.SheetNames.push("Removed OS");
wb.SheetNames.push("Added OS");
wb.SheetNames.push("Revived claims");

let wsRemovedOs = XLSX.utils.json_to_sheet(removedOsBeginToEndMonth);
let wsAddedOs = XLSX.utils.json_to_sheet(addedOsEndFromBeginMonth);
let wsCombinedOs = XLSX.utils.json_to_sheet(combinedSheets12);
let wsRevivedOs = XLSX.utils.json_to_sheet(revivedClaims);


wb.Sheets["Combined OS"] = wsCombinedOs;
wb.Sheets["Added OS"] = wsAddedOs;
wb.Sheets["Removed OS"] = wsRemovedOs;
wb.Sheets["Revived claims"] = wsRevivedOs;

XLSX.write(wb, { bookType: "xlsx", type: "binary" });




del.sync(["*.xlsx"]);
XLSX.writeFile(wb, "Final Report.xlsx");


// console.log(revivedClaims);
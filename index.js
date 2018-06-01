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
let sheet4 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[3]]);

let summary = [];

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
  INSURED: "insured",
  AGENCY: "Agency",
  "POLICY NO": "policyNo",
  "CLAIM NO": "claimNo",
  "INTIMATION RESERVE": "intimationReserve"
};


let paymentKeyChangesMap = {
    CLASS : "insuranceClass",
    "DATE OF CHEQUE" : "dateOfcheque",
    "CHEQUE NO"	: "chequeNo",
    "CLAIM NO":	"claimNo",
    "POLICY HOLDER"	:"placeHolder",
    "UW YEAR": "uwYear",
    "PAYEE":"payee",
    "PAID AMOUNT": "paidAmount"
};	

let osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
let osEndMonth = objectRenameKeys(sheet2, osEndMonthKeyChangesMap);
let intimated = objectRenameKeys(sheet3, intimatedKeyChangesMap);
let payment = objectRenameKeys(sheet4, paymentKeyChangesMap);
// //let intimated = objectRenameKeys(sheet3, osKeyChangesMap);

console.log(payment);

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

let osRepeatedClaimNo = osNoChange.map(function(claim) {
  let newObj = {};

  combinedOsWithDuplicates.map(function(combineClaim) {
    if (combineClaim.claimNo === claim.claimNo) {
      for (const prop in combineClaim) {
        newObj[prop] = combineClaim[prop];
      }
    }
  });

  return newObj;
});

// cover camelCase to NORMAL CASE Uppercase

function unCamelCase(str){
    str = str.replace(/([a-z\xE0-\xFF])([A-Z\xC0\xDF])/g, '$1 $2');
    str = str.toLowerCase(); //add space between camelCase text
    return str.toUpperCase();
  }
  

// convert "200,000.75" to 200000.75

function toFloat(stringValue) {
  return parseFloat(stringValue.replace(/,/g, ""));
}

function calcMovement(claim) {
  claim.difference =
    toFloat(claim.osEndMonthEstimate) - toFloat(claim.osBeginEstimate);
  return claim;
}

function movementTotal(movementArray) {
  return movementArray.reduce(function(prev, cur) {
    return prev + cur.difference;
  }, 0);
}

//  sheet with upward movement + differences
let movementUp = osRepeatedClaimNo
  .filter(function(claim) {
    if (toFloat(claim.osBeginEstimate) < toFloat(claim.osEndMonthEstimate)) {
      return claim;
    }
  })
  .map(calcMovement);

//  sheet with downward movement + differences
let movementDown = osRepeatedClaimNo
  .filter(function(claim) {
    if (toFloat(claim.osBeginEstimate) > toFloat(claim.osEndMonthEstimate)) {
      return claim;
    }
  })
  .map(calcMovement);

//   sheet with total up + down

let movementUpDown = _.concat(movementUp, movementDown);

function calculateMovementPerclass(targetArray) {
  let totalMovement = {};

  let motorPrivate = [];
  let motorPsvHire = [];
  let miscellaneous = [];
  let fireDomestic = [];
  let marine = [];
  let fireIndustrial = [];
  let liabilities = [];
  let motorCommercial = [];
  let accident = [];
  let engineering = [];
  let theft = [];
  let wiba = [];
  let medical = [];
  let count = {};

  targetArray.map(function(claim) {
    claimPrefixShort = claim.policyNo.slice(0, 6);
    claimPrefixLong = claim.policyNo.slice(0, 10);

    if (claimPrefixShort == "MGL/07") {
      motorPrivate.push(claim.difference);
      // motor private
    } else if (claimPrefixLong === "MGL/08/084") {
      // MOTOR PSV HIRE
      motorPsvHire.push(claim.difference);
    } else if (claimPrefixShort === "MGL/12") {
      // MISCELLANEOUS

      miscellaneous.push(claim.difference);
    } else if (claimPrefixShort === "MGL/03") {
      // FIRE DOMESTIC

      fireDomestic.push(claim.difference);
    } else if (claimPrefixShort === "MGL/06") {
      // MARINE

      marine.push(claim.difference);
    } else if (claimPrefixShort === "MGL/04") {
      fireIndustrial.push(claim.difference);
      // FIRE INDUSTRIAL
    } else if (claimPrefixShort === "MGL/05") {
      liabilities.push(claim.difference);
      // LIABILITIES
    } else if (claimPrefixShort === "MGL/08") {
      // MOTOR COMMERCIAL
      motorCommercial.push(claim.difference);
    } else if (claimPrefixShort === "MGL/02") {
      // ENGINEERING
      engineering.push(claim.difference);
    } else if (claimPrefixShort === "MGL/10") {
      // THEFT
      theft.push(claim.difference);
    } else if (claimPrefixShort === "MGL/11") {
      // WIBA
      wiba.push(claim.difference);
    } else if (
      claimPrefixLong === "MGL/09/096" ||
      claimPrefixLong === "MGL/09/091" ||
      claimPrefixLong === "MGL/09/099"
    ) {
      // medical
      medical.push(claim.difference);
    } else {
      // ACCIDENT
      accident.push(claim.difference);
    }
  });

  totalMovement.motorPrivate = _.sum(motorPrivate);
  totalMovement.motorPsvHire = _.sum(motorPsvHire);
  totalMovement.miscellaneous = _.sum(miscellaneous);
  totalMovement.fireDomestic = _.sum(fireDomestic);
  totalMovement.marine = _.sum(marine);
  totalMovement.fireIndustrial = _.sum(fireIndustrial);
  totalMovement.liabilities = _.sum(liabilities);
  totalMovement.motorCommercial = _.sum(motorCommercial);
  totalMovement.accident = _.sum(accident);
  totalMovement.engineering = _.sum(engineering);
  totalMovement.theft = _.sum(theft);
  totalMovement.wiba = _.sum(wiba);
  totalMovement.medical = _.sum(medical);

  totalMovement.count = {
    motorPrivate: motorPrivate.length,
    motorPsvHire: motorPsvHire.length,
    miscellaneous: miscellaneous.length,
    fireDomestic: fireDomestic.length,
    marine: marine.length,
    fireIndustrial: fireIndustrial.length,
    liabilities: liabilities.length,
    motorCommercial: motorCommercial.length,
    accident: accident.length,
    engineering: engineering.length,
    theft: theft.length,
    wiba: wiba.length,
    medical: medical.length
  };

  return totalMovement;
}

// total movement summary (up + down)
let movementSummary = [];

function getMovementSummary() {
  let newObj = {};

  newObj.CLASS = "TOTAL MOVEMENT";
  newObj.COUNT = movementUp.length + movementDown.length;
  newObj.TOTAL = movementTotal(movementUp) + movementTotal(movementDown);
  movementSummary.push(newObj);

  let movementObj = calculateMovementPerclass(movementUpDown);

  for (const prop in movementObj) {
    
    if(prop != 'count'){
      movementSummary.push({
        CLASS: unCamelCase(prop),
        COUNT: movementObj.count[prop],
        TOTAL: movementObj[prop]
      });
    }
  }
}

getMovementSummary();

/*  Insurance Classes 

ACCIDENT - GROUP PERSONAL ACCIDENT (MGL/09/092), INBOUND TRAVEL INSURANCE POLICY (MGL/09/095), INDIVIDUAL PERSONAL ACCIDENT (MGL/09/090), OVERSEAS TRAVEL INSURANCE COVER  (M1GL/09/097)- 

ENGINEERING - CONTRACTORS ALL RISK, ELECTRONIC EQUIPMENT, ERECTION ALL RISK, L.O.P. FOLLOWING MACHINERY B/DOWN, MACHINERY BREAKDOWN -MGL/02/

MOTOR PRIVATE - MOTOR CYCLE, MOTOR PRIVATE, MOTOR PRIVATE ENHANCED - MGL/07

MOTOR COMMERCIAL - MOTOR COMMERCIAL, MOTOR GENERAL CARTAGE, MOTOR TRACTORS, MOTOR TRADE - MGL/08

MOTOR PSV HIRE - MOTOR(PSV) PRIVATE HIRE - MGL/08/084

MISCELLANEOUS - BONDS ( I A TA) FINANCIAL GUARA, GOLFERS/SPORTSMAN INSURANCE - MGL/12

LIABILITIES - CARRIERS LIABILITY POLICY, CONTRACTUAL LIABILITY POLICY, PORT LIABILITY POLICY, PUBLIC LIABILITY, WAREHOUSE LIABILITY POLICY - MGL/05

FIRE DOMESTIC - FIRE DOMESTIC(HOC) - MGL/03/

MARINE - GOODS IN TRANSIT, MARINE CARGO, MARINE HULL, MARINE OPEN COVER - MGL/06

FIRE INDUSTRIAL - FIRE INDUSTRIAL, INDUSTRIAL ALL RISKS - MGL/04

MEDICAL - ACCIDENT HOSPITALISATION INS.P (MGL/09/099), HEALTH/MEDICAL EXPENSES INSURANCE (MGL/09/091), INDIVIDUAL MEDICAL INSURANCE (MGL/09/096)

WIBA - WORKERS INJURUY BENEFIT ACT, WORKMEN'S COMP (COMMON LAW) COVER, WORKMEN'S COMPENSATION(ACT) CO - MGL/11/

THEFT - ALL RISKS, BANKERS BLANKET INSURANCE, BUGRLARY, CASH IN TRANSIT, FIDELITY GUARANTEE - MGL/10/

*/

// combined sheet OS begining and end no dubplictes:

let combinedSheets12 = _.concat(
  addedOsEndFromBeginMonth,
  removedOsBeginToEndMonth,
  osRepeatedClaimNo,
  "claimNO"
).map(function(claim) {
  let newObj = {};
  for (const prop in claim) {
    newObj[prop] = claim[prop];
    if (claim.osEndMonthEstimate == null) {
      newObj.osEndMonthEstimate = 0;
    } else if (claim.osBeginEstimate == null) {
      newObj.osBeginEstimate = 0;
    }
  }

  return newObj;
});

// find values that are revived : present in Added but not intimated for this month

let revivedClaims = _.differenceBy(
  addedOsEndFromBeginMonth,
  intimated,
  "claimNo"
);



let closedAsNoClaim = _.differenceBy(removedOsBeginToEndMonth, payment, "claimNo");

console.log( "removed ---> : " + removedOsBeginToEndMonth.length + "\n  paid---> : " + payment.length );


// create work book
let wb = XLSX.utils.book_new();

// create sheetsNames
wb.SheetNames.push(
  "Combined OS",
  "Removed OS",
  "Added OS",
  "Revived claims",
  "in OSBegining & OSend",
  "Movement up",
  "Movement down",
  "summary"
);

let wsRemovedOs = XLSX.utils.json_to_sheet(removedOsBeginToEndMonth);
let wsAddedOs = XLSX.utils.json_to_sheet(addedOsEndFromBeginMonth);
let wsCombinedOs = XLSX.utils.json_to_sheet(combinedSheets12);
let wsRevivedOs = XLSX.utils.json_to_sheet(revivedClaims);
let wsInBeginingEnd = XLSX.utils.json_to_sheet(osRepeatedClaimNo);
let wsUpMovement = XLSX.utils.json_to_sheet(movementUp);
let wsDownMovement = XLSX.utils.json_to_sheet(movementDown);
let wsSummary = XLSX.utils.json_to_sheet(movementSummary, {
  header: ["CLASS", "COUNT", "TOTAL"]
});

wb.Sheets["Combined OS"] = wsCombinedOs;
wb.Sheets["Added OS"] = wsAddedOs;
wb.Sheets["Removed OS"] = wsRemovedOs;
wb.Sheets["in OSBegining & OSend"] = wsInBeginingEnd;
wb.Sheets["Movement up"] = wsUpMovement;
wb.Sheets["Movement down"] = wsDownMovement;
wb.Sheets["Revived claims"] = wsRevivedOs;
wb.Sheets["summary"] = wsSummary;

XLSX.write(wb, { bookType: "xlsx", type: "binary" });

del.sync(["*.xlsx"]);
XLSX.writeFile(wb, "Final Report.xlsx");

// console.log(revivedClaims);

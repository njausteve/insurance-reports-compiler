const XLSX = require("xlsx");
let workbook = XLSX.readFile("./source file/insurance.xlsx");
let sheetNameList = workbook.SheetNames;
const _ = require("lodash");
const del = require("del");
const objectRenameKeys = require("object-rename-keys");

let sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
let sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
let sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[2]]);
let sheet4 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[3]]);

let sheets = [sheet1, sheet2, sheet3, sheet4];



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
  "INTIMATION RESERVE": "intimationReserve"
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
  difference: "DIFFERENCE"
};

// check sheet Fields

function checkSheetFields() {
  let status = [];

  if (sheets.length < 4) {
    status.push({ sheet: "no sheets", error: "one or more sheets are missing" });
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

   if(status.length < 1){
      status.push({sheet: "all", message: "all sheets OKAY"});   
   }

  return status;
}

console.log(checkSheetFields());

let osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
let intimated = objectRenameKeys(sheet2, intimatedKeyChangesMap);
let payment = objectRenameKeys(sheet3, paymentKeyChangesMap);
let osEndMonth = objectRenameKeys(sheet4, osEndMonthKeyChangesMap);

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

// sheet with those that appear ib begining and End OS estimate
let osNoChange = _.intersectionBy(osBeginMonth, osEndMonth, "claimNo");

// sheet with those that appear ib begining and End OS estimate : repeated
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

// convert all keys to original
function toExcelSheet(sheetToprint) {
  return objectRenameKeys(sheetToprint, printToExcelKeysmap);
}

// cover camelCase to NORMAL CASE Uppercase

function unCamelCase(str) {
  str = str.replace(/([a-z\xE0-\xFF])([A-Z\xC0\xDF])/g, "$1 $2");
  str = str.toLowerCase(); //add space between camelCase text
  return str.toUpperCase();
}

// convert to currency format

function toCurrency(value) {
  return value.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, "$1,");
}

// convert "200,000.75" to 200000.75

function toFloat(stringValue) {
  return parseFloat(stringValue.toString().replace(/,/g, ""));
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

function calculatePerclass(targetArray, valueUsedToCalculate) {
  let ValuesPerClass = {};
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

    valueToPush = toFloat(claim[valueUsedToCalculate]);

    if (claimPrefixShort == "MGL/07") {
      motorPrivate.push(valueToPush);

      // motor private
    } else if (claimPrefixLong === "MGL/08/084") {
      // MOTOR PSV HIRE
      motorPsvHire.push(valueToPush);
    } else if (claimPrefixShort === "MGL/12") {
      // MISCELLANEOUS

      miscellaneous.push(valueToPush);
    } else if (claimPrefixShort === "MGL/03") {
      // FIRE DOMESTIC

      fireDomestic.push(valueToPush);
    } else if (claimPrefixShort === "MGL/06") {
      // MARINE

      marine.push(valueToPush);
    } else if (claimPrefixShort === "MGL/04") {
      fireIndustrial.push(valueToPush);
      // FIRE INDUSTRIAL
    } else if (claimPrefixShort === "MGL/05") {
      liabilities.push(valueToPush);
      // LIABILITIES
    } else if (claimPrefixShort === "MGL/08") {
      // MOTOR COMMERCIAL
      motorCommercial.push(valueToPush);
    } else if (claimPrefixShort === "MGL/02") {
      // ENGINEERING
      engineering.push(valueToPush);
    } else if (claimPrefixShort === "MGL/10") {
      // THEFT
      theft.push(valueToPush);
    } else if (claimPrefixShort === "MGL/11") {
      // WIBA
      wiba.push(valueToPush);
    } else if (
      claimPrefixLong === "MGL/09/096" ||
      claimPrefixLong === "MGL/09/091" ||
      claimPrefixLong === "MGL/09/099"
    ) {
      // medical
      medical.push(valueToPush);
    } else {
      // ACCIDENT
      accident.push(valueToPush);
    }
  });

  ValuesPerClass.motorPrivate = _.sum(motorPrivate);
  ValuesPerClass.motorPsvHire = _.sum(motorPsvHire);
  ValuesPerClass.miscellaneous = _.sum(miscellaneous);
  ValuesPerClass.fireDomestic = _.sum(fireDomestic);
  ValuesPerClass.marine = _.sum(marine);
  ValuesPerClass.fireIndustrial = _.sum(fireIndustrial);
  ValuesPerClass.liabilities = _.sum(liabilities);
  ValuesPerClass.motorCommercial = _.sum(motorCommercial);
  ValuesPerClass.accident = _.sum(accident);
  ValuesPerClass.engineering = _.sum(engineering);
  ValuesPerClass.theft = _.sum(theft);
  ValuesPerClass.wiba = _.sum(wiba);
  ValuesPerClass.medical = _.sum(medical);

  ValuesPerClass.count = {
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

  return ValuesPerClass;
}

// total movement summary (up + down)

function getMovementSummary() {
  let movementSummary = [];
  let newObj = {};
  let movementObj = calculatePerclass(movementUpDown, "difference");

  for (const prop in movementObj) {
    if (prop != "count") {
      movementSummary.push({
        CLASS: unCamelCase(prop),
        COUNT: movementObj.count[prop],
        TOTAL: toCurrency(movementObj[prop])
      });
    }
  }

  movementSummary.push({
    CLASS: "TOTAL SUM",
    COUNT: movementUp.length + movementDown.length,
    TOTAL: toCurrency(movementTotal(movementUp) + movementTotal(movementDown))
  });

  return movementSummary;
}

// get summary for any sheet
function getSummary(targetSheet, valueToRefer) {
  let summary = [];
  let summaryObj = calculatePerclass(targetSheet, valueToRefer);

  let sumTotal = 0;

  for (const prop in summaryObj) {
    if (prop != "count") {
      sumTotal = sumTotal + summaryObj[prop];

      summary.push({
        CLASS: unCamelCase(prop),
        COUNT: summaryObj.count[prop],
        TOTAL: toCurrency(summaryObj[prop])
      });
    }
  }

  summary.push({
    CLASS: "TOTAL SUM",
    COUNT: _.sum(_.values(summaryObj.count)),
    TOTAL: toCurrency(sumTotal)
  });

  return summary;
}

// claims in intimated and paid without payment data

let intimatedAndPaidIncomplete = _.intersectionBy(
  intimated,
  payment,
  "claimNo"
);

// intimated and paid all {duplicates}

let intimatedPaidMovementwithDuplicates = _.concat(intimated, payment);

let intimatedPaidMovement = intimatedAndPaidIncomplete.map(function(claim) {
  let newObj = {};

  intimatedPaidMovementwithDuplicates.map(function(combineClaim) {
    if (combineClaim.claimNo === claim.claimNo) {
      for (const prop in combineClaim) {
        newObj[prop] = combineClaim[prop];

        if (newObj.paidAmount != undefined) {
          newObj.difference =
            toFloat(newObj.paidAmount) - toFloat(newObj.intimationReserve);
        }
      }
    }
  });

  return newObj;
});

let intimatedPaidSummary = getSummary(intimatedPaidMovement, "difference");

// combined sheet OS begining and end no dubplictes:

let combinedSheets12 = _
  .concat(
    addedOsEndFromBeginMonth,
    removedOsBeginToEndMonth,
    osRepeatedClaimNo,
    "claimNO"
  )
  .map(function(claim) {
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

// sheet for claims closed as having no claim
let closedAsNoClaim = _.differenceBy(
  removedOsBeginToEndMonth,
  payment,
  "claimNo"
);

// summary sheets

let closedAsNoClaimSummary = getSummary(closedAsNoClaim, "osBeginEstimate");

let revivedClaimsSummary = getSummary(revivedClaims, "osEndMonthEstimate");

let movementSummary = getMovementSummary();

let totalMovementSummary = movementSummary.map(function(moveClass) {
  let newObj = {};

  intimatedPaidSummary.map(function(intPaidClass) {
    for (const prop in intPaidClass) {
      if (intPaidClass[prop] == moveClass.CLASS) {
        newObj = {
          CLASS: moveClass.CLASS,
          COUNT: intPaidClass.COUNT + moveClass.COUNT,
          TOTAL: toCurrency(
            toFloat(intPaidClass.TOTAL) + toFloat(moveClass.TOTAL)
          )
        };
      }
    }
  });

  return newObj;
});

// create work book
let wb = XLSX.utils.book_new();

// create sheetsNames
wb.SheetNames.push(
  "MOVEMENT SUMMARY",
  "CLOSED AS NO CLAIM SUMMARY",
  "REVIVED CLAIMS SUMMARY",
  "REMOVED CLAIMS",
  "ADDED CLAIMS",
  "REVIVED CLAIMS",
  "CLOSED AS NO CLAIM",
  "CLAIMS IN LAST AND CURRENT OS",
  "MOVED UP CLAIMS",
  "MOVED DOWN CLAIMS",
  "ALL COMBINED SORTED"
);

let summaryHeader = {
  header: ["CLASS", "COUNT", "TOTAL"]
};

let movementHeader = {
  header: [
    "CLASS",
    "POLICY NO",
    "CLAIM NO",
    "INSURED NAME",
    "DATE REPORTED",
    "DATE OF LOSS",
    "PERIOD FROM",
    "PERIOD TO",
    "BEGINING OS ESTIMATE",
    "END OS ESTIMATE",
    "DIFFERENCE"
  ]
};

let closedAsnoClaimHeader = {
  header: [
    "CLASS",
    "POLICY NO",
    "CLAIM NO",
    "INSURED NAME",
    "DATE REPORTED",
    "DATE OF LOSS",
    "PERIOD FROM",
    "PERIOD TO",
    "BEGINING OS ESTIMATE"
  ]
};

let revivedHeader = {
  header: [
    "CLASS",
    "POLICY NO",
    "CLAIM NO",
    "INSURED NAME",
    "DATE REPORTED",
    "DATE OF LOSS",
    "PERIOD FROM",
    "PERIOD TO",
    "END OS ESTIMATE"
  ]
};

let wsRemovedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(removedOsBeginToEndMonth)
);
let wsAddedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(addedOsEndFromBeginMonth)
);
let wsCombinedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(intimatedPaidMovement)
);
let wsRevivedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(revivedClaims),
  revivedHeader
);
let wsClosedASNoClaim = XLSX.utils.json_to_sheet(
  toExcelSheet(closedAsNoClaim),
  closedAsnoClaimHeader
);
let wsInBeginingEnd = XLSX.utils.json_to_sheet(toExcelSheet(osRepeatedClaimNo));

let wsUpMovement = XLSX.utils.json_to_sheet(
  toExcelSheet(movementUp),
  movementHeader
);
let wsDownMovement = XLSX.utils.json_to_sheet(
  toExcelSheet(movementDown),
  movementHeader
);
let wsMovementSummary = XLSX.utils.json_to_sheet(
  totalMovementSummary,
  summaryHeader
);
let wsClosedAsNoClaimSummary = XLSX.utils.json_to_sheet(
  closedAsNoClaimSummary,
  summaryHeader
);
let wsRevivedClaimSummary = XLSX.utils.json_to_sheet(
  revivedClaimsSummary,
  summaryHeader
);

wb.Sheets["ALL COMBINED SORTED"] = wsCombinedOs;
wb.Sheets["ADDED CLAIMS"] = wsAddedOs;
wb.Sheets["REMOVED CLAIMS"] = wsRemovedOs;
wb.Sheets["CLOSED AS NO CLAIM"] = wsClosedASNoClaim;
wb.Sheets["CLAIMS IN LAST AND CURRENT OS"] = wsInBeginingEnd;
wb.Sheets["MOVED UP CLAIMS"] = wsUpMovement;
wb.Sheets["MOVED DOWN CLAIMS"] = wsDownMovement;
wb.Sheets["REVIVED CLAIMS"] = wsRevivedOs;
wb.Sheets["MOVEMENT SUMMARY"] = wsMovementSummary;
wb.Sheets["CLOSED AS NO CLAIM SUMMARY"] = wsClosedAsNoClaimSummary;
wb.Sheets["REVIVED CLAIMS SUMMARY"] = wsRevivedClaimSummary;

XLSX.write(wb, { bookType: "xlsx", type: "binary" });

del.sync(["*.xlsx"]);
XLSX.writeFile(wb, "Final Report.xlsx");

// TODO: separate code into modules.

//TODO: change formating of values to end sheets

// TODO: remove difference col in CLAIMS IN LAST AND CURRENT OS

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

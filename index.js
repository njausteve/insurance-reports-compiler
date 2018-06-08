const XLSX = require("xlsx");
const workbook = XLSX.readFile("./source file/insurance.xlsx");
const XlsxPopulate = require("xlsx-populate");
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
  difference: "DIFFERENCE"
};

// check sheet Fields

function checkSheetFields() {
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
}

console.log(checkSheetFields());

// assignment

let osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
let intimatedWithZeros = objectRenameKeys(sheet2, intimatedKeyChangesMap);
let paymentWithDuplicate = objectRenameKeys(sheet3, paymentKeyChangesMap);
let osEndMonth = objectRenameKeys(sheet4, osEndMonthKeyChangesMap);

/* ============== utility/ helper functions here =============*/

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
  let newObj = {};

  for (const prop in claim) {
    newObj[prop] = claim[prop];
  }

  newObj.difference = (
    toFloat(newObj.osEndMonthEstimate) - toFloat(newObj.osBeginEstimate)
  ).toFixed(2);

  return newObj;
}

function movementTotal(movementArray) {
  return movementArray.reduce(function(prev, cur) {
    return prev + cur.difference;
  }, 0);
}

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
    if (claim.policyNo) {
      claimPrefixShort = claim.policyNo.slice(0, 6);
      claimPrefixLong = claim.policyNo.slice(0, 10);
    }

    insClass = claim.insuranceClass.trim().toString();

    valueToPush = toFloat(claim[valueUsedToCalculate]);

    if (
      insClass === "MOTOR CYCLE" ||
      insClass === "MOTOR PRIVATE" ||
      insClass === "MOTOR PRIVATE ENHANCED"
    ) {
      motorPrivate.push(valueToPush);

      // motor private
    } else if (insClass === "MOTOR (PSV) PRIVATE HIRE") {
      // MOTOR PSV HIRE
      motorPsvHire.push(valueToPush);
    } else if (
      insClass === "BONDS ( I A TA) FINANCIAL GUARA" ||
      insClass === "GOLFERS/SPORTSMAN INSURANCE"
    ) {
      // MISCELLANEOUS

      miscellaneous.push(valueToPush);
    } else if (insClass === "FIRE DOMESTIC (HOC)") {
      // FIRE DOMESTIC

      fireDomestic.push(valueToPush);
    } else if (
      insClass === "GOODS IN TRANSIT" ||
      insClass === "MARINE CARGO" ||
      insClass === "MARINE OPEN COVER" ||
      insClass === "MARINE HULL"
    ) {
      // MARINE

      marine.push(valueToPush);
    } else if (
      insClass === "INDUSTRIAL ALL RISKS" ||
      insClass === "FIRE INDUSTRIAL"
    ) {
      fireIndustrial.push(valueToPush);
      // FIRE INDUSTRIAL
    } else if (
      insClass === "CARRIERS LIABILITY POLICY" ||
      insClass === "CONTRACTUAL LIABILITY POLICY" ||
      insClass === "PORT LIABILITY POLICY" ||
      insClass === "PUBLIC LIABILITY" ||
      insClass === "WAREHOUSE LIABILITY POLICY"
    ) {
      liabilities.push(valueToPush);
      // LIABILITIES
    } else if (
      insClass === "MOTOR COMMERCIAL" ||
      insClass === "MOTOR GENERAL CARTAGE" ||
      insClass === "MOTOR TRACTORS" ||
      insClass === "MOTOR TRADE"
    ) {
      // MOTOR COMMERCIAL
      motorCommercial.push(valueToPush);
    } else if (
      insClass === "CONTRACTORS ALL RISKS" ||
      insClass === "ELECTRONIC EQUIPMENT" ||
      insClass === "ERECTION ALL RISKS" ||
      insClass === "L.O.P. FOLLOWING MACHINERY B/DOWN" ||
      insClass === "MACHINERY BREAKDOWN"
    ) {
      // ENGINEERING
      engineering.push(valueToPush);
    } else if (
      insClass === "ALL RISKS" ||
      insClass === "BANKERS BLANKET INSURANCE" ||
      insClass === "BURGLARY" ||
      insClass === "CASH IN TRANSIT" ||
      insClass === "FIDELITY GUARANTEE"
    ) {
      // THEFT
      theft.push(valueToPush);
    } else if (
      insClass === "WORKERS INJURY BENEFIT ACT" ||
      insClass === "WORKMEN'S COMP (COMMON LAW) COVER" ||
      insClass === "WORKMEN'S COMPENSATION (ACT) CO"
    ) {
      // WIBA
      wiba.push(valueToPush);
    } else if (
      insClass === "ACCIDENT HOSPITALISATION INS. P" ||
      insClass === "HEALTH/MEDICAL EXPENSES INSURANCE" ||
      insClass === "INDIVIDUAL MEDICAL INSURANCE"
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

/* ============== movement calcultions section  =============*/

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

// sheet with those that appear in begining and End OS estimate : repeated
let osRepeatedClaimNo = osNoChange
  .map(function(claim) {
    let newObj = {};

    combinedOsWithDuplicates.map(function(combineClaim) {
      if (combineClaim.claimNo === claim.claimNo) {
        for (const prop in combineClaim) {
          newObj[prop] = combineClaim[prop];
        }
      }
    });

    return newObj;
  })
  .filter(
    claim =>
      toFloat(claim.osBeginEstimate) - toFloat(claim.osEndMonthEstimate) != 0
  );

// payments adjustments

let uniquePaidClaimNo = _
  .uniqBy(paymentWithDuplicate, "claimNo")
  .map(claim => claim.claimNo);

let payments = uniquePaidClaimNo
  .map(function(claimNo) {
    let totalPaid = 0;
    let newObj = {};

    paymentWithDuplicate.map(function(dupClaim) {
      for (const prop in dupClaim) {
        if (claimNo === dupClaim.claimNo) {
          if (prop === "paidAmount") {
            totalPaid = totalPaid + toFloat(dupClaim.paidAmount);

            newObj.paidAmount = totalPaid;
          } else {
            newObj[prop] = dupClaim[prop];
          }
        }
      }
    });

    return newObj;
  })
  .filter(claim => claim.paidAmount != 0);

// remove zero values claims from intimated
let intimated = intimatedWithZeros.filter(
  claim => claim.intimationReserve != 0
);

//============== BeginingEndMovement without Paid =============

//  sheet with upward movement + differences
let movedUpWithDifference = osRepeatedClaimNo
  .filter(function(claim) {
    return toFloat(claim.osBeginEstimate) < toFloat(claim.osEndMonthEstimate);
  })
  .map(calcMovement);

let movementUpNoPaid = _.differenceBy(
  movedUpWithDifference,
  payments,
  "claimNo"
);

//  sheet with downward movement + differences - paid
let movementDownWithDifference = osRepeatedClaimNo

  .filter(function(claim) {
    return toFloat(claim.osBeginEstimate) > toFloat(claim.osEndMonthEstimate);
  })
  .map(calcMovement);

let movementDownNoPaid = _.differenceBy(
  movementDownWithDifference,
  payments,
  "claimNo"
);

//   sheet with total up + down that do not appear in Payments

let beginEndMovementNoPaid = _.concat(movementUpNoPaid, movementDownNoPaid);

//============== BeginingEndMovement with Paid =============

//  Begining - End movement
let movementWithPaidUniqueClaimNo = _
  .intersectionBy(osRepeatedClaimNo, payments, "claimNo")
  .map(claim => claim.claimNo);

// combine all those in Begining, End and In paid (with duplicates);
let begEndPaidCombined = _.concat(osRepeatedClaimNo, payments);

//  Begining - End movement with claim paid (paidAmount + EndOsEstimate) - beginingOsEstimate
let beginEndmovementWithPaid = movementWithPaidUniqueClaimNo
  .map(function(claimNo) {
    let newObj = {};

    begEndPaidCombined.map(function(dupClaim) {
      if (dupClaim.claimNo === claimNo) {
        for (const prop in dupClaim) {
          newObj[prop] = dupClaim[prop];
        }
      }
    });

    return newObj;
  })
  .map(function(claim) {
    claim.difference = (
      toFloat(claim.paidAmount) +
      toFloat(claim.osEndMonthEstimate) -
      toFloat(claim.osBeginEstimate)
    ).toFixed(2);

    return claim;
  })
  .filter(claim => claim.difference != 0);


//============== BeginingPaidEndMovement without End OS estimate =============

// claims in Begining OS estimates and paid (payments) without paid amount data
let beginPaidIncomplete = _.intersectionBy(osBeginMonth, payments, "claimNo");

// Begining OS estimates and paid (payments) all {duplicates}
let beginPaidMovementDuplicates = _.concat(osBeginMonth, payments);

// sheet with Begining OS estimates and paid (payments) movements
let beginPaid = beginPaidIncomplete
  .map(function(claim) {
    let newObj = {};

    beginPaidMovementDuplicates.map(function(combineClaim) {
      if (combineClaim.claimNo === claim.claimNo) {
        for (const prop in combineClaim) {
          newObj[prop] = combineClaim[prop];

          if (newObj.paidAmount != undefined) {
            newObj.difference =
              toFloat(newObj.paidAmount) - toFloat(newObj.osBeginEstimate);
          }
        }
      }
    });

    return newObj;
  })
  .filter(claim => claim.difference != 0);

// sheet with Begining OS estimates and paid (payments) movements that are not In OsEndmonth
let beginPaidMovementNoEndOS = _.differenceBy(beginPaid, osEndMonth, "claimNo");





// ============== intimation related movements =============

// claimsNo for claims in intimated and paid
let intimatedAndPaidClaimNo = _
  .intersectionBy(intimated, payments, "claimNo")
  .map(claim => claim.claimNo);

// intimated and paid all {duplicates}

let intimatedPaidMovementwithDuplicates = _.concat(intimated, payments);

// claims intimimated + paid (including paidAmount)
let intimatedPaidNoDifference = intimatedAndPaidClaimNo
  .map(function(claimNo) {
    let newObj = {};

    intimatedPaidMovementwithDuplicates.map(function(combineClaim) {
      if (combineClaim.claimNo === claimNo) {
        for (const prop in combineClaim) {
          newObj[prop] = combineClaim[prop];
        }
      }
    });

    return newObj;
  })
  .filter(
    claim => toFloat(claim.intimationReserve) != toFloat(claim.paidAmount)
  );

// intimatedPaidmovement with Difference included
let intimatedPaidMovement = intimatedPaidNoDifference.map(function(claim) {
  let newObj = {};

  for (const prop in claim) {
    newObj[prop] = claim[prop];
  }

  newObj.difference = (
    toFloat(newObj.paidAmount) - toFloat(newObj.intimationReserve)
  ).toFixed(2);

  return newObj;
});

// intimatedPaidmovement without EndOS

let intimatedPaidmovementNoEndOS = _.differenceBy(
  intimatedPaidMovement,
  osEndMonth,
  "claimNo"
);

let intimatedPaidEndOsWithDuplicates = _.concat(
  intimatedPaidNoDifference,
  osEndMonth
);

// intimatedPaid movement with EndOS

let intimatedPaidMovementWithEndOs = _
  .intersectionBy(intimatedPaidNoDifference, osEndMonth, "claimNo")
  .map(claim => claim.claimNo)
  .map(function(claimNo) {
    let newObj = {};

    intimatedPaidEndOsWithDuplicates.map(function(combineClaim) {
      if (combineClaim.claimNo === claimNo) {
        for (const prop in combineClaim) {
          newObj[prop] = combineClaim[prop];
        }
      }
    });
    return newObj;
  })
  .map(function(claim) {
    let newObj = {};

    for (const prop in claim) {
      newObj[prop] = claim[prop];
    }

    newObj.difference = (
      toFloat(newObj.osEndMonthEstimate) +
      toFloat(newObj.paidAmount) -
      toFloat(newObj.intimationReserve)
    ).toFixed(2);

    return newObj;
  });

//=============== intimated End no paid =====================

// claims in intimated and EndOs Estimates without Endestimate amount data
let intimatedEndOsIncomplete = _.intersectionBy(
  intimated,
  osEndMonth,
  "claimNo"
);

// intimated and End Os all {duplicates}
let intimatedEndOsMovementDuplicates = _.concat(intimated, osEndMonth);

// sheet with intimated - EndOS movement
let intimatedEndOsMovement = intimatedEndOsIncomplete
  .map(function(claim) {
    let newObj = {};
    intimatedEndOsMovementDuplicates.map(function(combineClaim) {
      if (combineClaim.claimNo === claim.claimNo) {
        for (const prop in combineClaim) {
          newObj[prop] = combineClaim[prop];
          if (newObj.osEndMonthEstimate != undefined) {
            newObj.difference =
              toFloat(newObj.osEndMonthEstimate) -
              toFloat(newObj.intimationReserve);
          }
        }
      }
    });

    return newObj;
  })
  .filter(claim => claim.difference != 0);

// intimated-End movement with no claims from paid

let intimatedEndOsMovementNoPaid = _.differenceBy(
  intimatedEndOsMovement,
  payments,
  "claimNo"
);


/* ================== closed As No claim  =============== */

let intimatedNotPaidclaims = _.differenceBy(intimated, payments, 'claimNo');


// sheet for claims closed as having no claim From Intimated 

let intimatedClosedAsNoClaim = _.differenceBy(intimatedNotPaidclaims, osEndMonth, 'claimNo');

// sheet for claims closed as having no claim From Begining OS
let beginingClosedAsNoClaim = _.differenceBy(
  removedOsBeginToEndMonth,
  payments,
  "claimNo"
);



console.log("beginingClosedAsNoClaim", beginingClosedAsNoClaim);



// total movement =   beginPaidMovement + intimatedEndOsMovement + intimatedPaidMovement + movementUpDown










/*



let totalMovementWithDuplicates = _.concat(
  movementUpDown,
  beginPaidMovementNoEndOS,
  intimatedEndOsMovement,
  intimatedPaidMovement
);

let totalMovementUniqueClaimNo = _
  .uniqBy(totalMovementWithDuplicates, "claimNo")
  .map(claim => claim.claimNo);

let totalMovement = totalMovementUniqueClaimNo
  .map(function(claimNo) {
    let totaldifference = 0;
    let newObj = {};

    totalMovementWithDuplicates.map(function(dupClaim) {
      for (const prop in dupClaim) {
        if (claimNo === dupClaim.claimNo) {
          if (prop === "difference") {
            totaldifference = totaldifference + toFloat(dupClaim.difference);
            newObj.difference = totaldifference;
          } else {
            newObj[prop] = dupClaim[prop];
          }
        }
      }
    });

    return newObj;
  })
  .filter(claim => claim.difference != 0);

// combined sheet OS begining and end no dubplictes:

// let combinedSheets = _
//   .concat(
//     addedOsEndFromBeginMonth,
//     removedOsBeginToEndMonth,
//     osRepeatedClaimNo,
//     "claimNo"
//   )
//   .map(function(claim) {
//     let newObj = {};
//     for (const prop in claim) {
//       newObj[prop] = claim[prop];
//       if (claim.osEndMonthEstimate == null) {
//         newObj.osEndMonthEstimate = 0;
//       } else if (claim.osBeginEstimate == null) {
//         newObj.osBeginEstimate = 0;
//       }
//     }

//     return newObj;
//   });

let combined12 = _.concat(osBeginMonth, intimated);

// find values that are revived : present in Added but not intimated for this month

let revivedClaims = _.differenceBy(
  addedOsEndFromBeginMonth,
  intimated,
  "claimNo"
);



// summary sheets

let closedAsNoClaimSummary = getSummary(closedAsNoClaim, "osBeginEstimate");

let revivedClaimsSummary = getSummary(revivedClaims, "osEndMonthEstimate");

let movementSummary = getMovementSummary();

let paidSummary = getSummary(payments, "paidAmount");

let intimatedSummary = getSummary(intimated, "intimationReserve");

// let totalMovementSummary = movementSummary.map(function(moveClass) {
//   let newObj = {};

//   intimatedPaidSummary.map(function(intPaidClass) {
//     for (const prop in intPaidClass) {
//       if (intPaidClass[prop] == moveClass.CLASS) {
//         newObj = {
//           CLASS: moveClass.CLASS,
//           COUNT: intPaidClass.COUNT + moveClass.COUNT,
//           TOTAL: toCurrency(
//             toFloat(intPaidClass.TOTAL) + toFloat(moveClass.TOTAL)
//           )
//         };
//       }
//     }
//   });

//   return newObj;
// });

let totalMovementSummary = getSummary(totalMovement, "difference");

console.log("intimatedSummary", intimatedSummary);

// create work book
let wb = XLSX.utils.book_new();

// create sheetsNames
wb.SheetNames.push(
  "PAID SUMMARY",
  "INTIMATED SUMMARY",
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
  "INT-OS END MOVEMENT",
  "INT-PAID MOVEMENT",
  "OS BEGIN-PAID MOVEMENT",
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

let wsIntimatedEndOSmovement = XLSX.utils.json_to_sheet(
  toExcelSheet(intimatedEndOsMovement)
);

let wsIntimatedPaidMovement = XLSX.utils.json_to_sheet(
  toExcelSheet(intimatedPaidMovement)
);

let wsbeginPaidMovement = XLSX.utils.json_to_sheet(
  toExcelSheet(beginPaidMovementNoEndOS)
);

let wsRemovedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(removedOsBeginToEndMonth)
);
let wsAddedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(addedOsEndFromBeginMonth)
);
let wsCombinedOs = XLSX.utils.json_to_sheet(
  toExcelSheet(totalMovementWithDuplicates)
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

let wsPaidSummary = XLSX.utils.json_to_sheet(paidSummary, summaryHeader);

let wsIntimatedSummary = XLSX.utils.json_to_sheet(
  intimatedSummary,
  summaryHeader
);

let wsheets = [
  wsMovementSummary,
  wsClosedAsNoClaimSummary,
  wsRevivedClaimSummary,
  wsPaidSummary,
  wsIntimatedSummary,
  wsRemovedOs,
  wsAddedOs,
  wsCombinedOs,
  wsRevivedOs,
  wsClosedASNoClaim,
  wsInBeginingEnd,
  wsUpMovement,
  wsDownMovement
];

// formating of column widths

wsheets.map(function(sheet, index) {
  if (index < 6) {
    sheet["!cols"] = [{ wch: 20 }, { wch: 10 }, { wch: 20 }];
    sheet.A1.s = {
      patternType: "solid",
      fgColor: { theme: 8, tint: 0.3999755851924192 },
      bgColor: { indexed: 64 }
    };
  } else {
    sheet["!cols"] = [
      { wch: 40 },
      { wch: 30 },
      { wch: 30 },
      { wch: 30 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 }
    ];
  }

  sheet["!margins"] = {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3
  };
});

wb.Sheets["PAID SUMMARY"] = wsPaidSummary;
wb.Sheets["INTIMATED SUMMARY"] = wsIntimatedSummary;
wb.Sheets["MOVEMENT SUMMARY"] = wsMovementSummary;
wb.Sheets["CLOSED AS NO CLAIM SUMMARY"] = wsClosedAsNoClaimSummary;
wb.Sheets["REVIVED CLAIMS SUMMARY"] = wsRevivedClaimSummary;
wb.Sheets["ALL COMBINED SORTED"] = wsCombinedOs;
wb.Sheets["ADDED CLAIMS"] = wsAddedOs;
wb.Sheets["REMOVED CLAIMS"] = wsRemovedOs;
wb.Sheets["CLOSED AS NO CLAIM"] = wsClosedASNoClaim;
wb.Sheets["CLAIMS IN LAST AND CURRENT OS"] = wsInBeginingEnd;
wb.Sheets["MOVED UP CLAIMS"] = wsUpMovement;
wb.Sheets["MOVED DOWN CLAIMS"] = wsDownMovement;
wb.Sheets["INT-OS END MOVEMENT"] = wsIntimatedEndOSmovement;
wb.Sheets["INT-PAID MOVEMENT"] = wsIntimatedPaidMovement;
wb.Sheets["OS BEGIN-PAID MOVEMENT"] = wsbeginPaidMovement;
wb.Sheets["REVIVED CLAIMS"] = wsRevivedOs;

XLSX.write(wb, { bookType: "xlsx", type: "binary" });

del.sync(["./tmp/*.xlsx", "*.xlsx"]);

XLSX.writeFile(wb, "./tmp/Unstyled Report.xlsx");

XlsxPopulate.fromFileAsync("./tmp/Unstyled Report.xlsx")
  .then(function(otherWorkBook) {
    const sheets = otherWorkBook.sheets();

    sheets.map(function(sheet) {
      sheet.row(1).style({ bold: true, fill: "ffff00", fontSize: 14 });
    });

    return otherWorkBook.toFileAsync("./Final Report.xlsx");
  })
  .catch(err => console.error(err));

// TODO: separate code into modules.

//TODO: change formating of values to end sheets

// TODO: remove difference col in CLAIMS IN LAST AND CURRENT OS
//TODO: remove empty  row on sheets

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

// console.log("paid summary", calculatePerclass(payment, "paidAmount"));

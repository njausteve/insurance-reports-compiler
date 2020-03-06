const XLSX = require("xlsx");
const XlsxPopulate = require("xlsx-populate");
const _ = require("lodash");
const del = require("del");
const path = require("path");
const objectRenameKeys = require("object-rename-keys");

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

// headers for excel sheets

let summaryHeader = {
  header: ["CLASS", "COUNT", "TOTAL"]
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

let combinedAllHeader = {
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
    "INTIMATION RESERVE",
    "PAID AMOUNT",
    "END OS ESTIMATE",
    "DIFFERENCE",
    "REVIVED",
    "MOVED",
    "NO CLAIM",
    "INSURED",
    "AGENCY",
    "DATE OF CHEQUE",
    "CHEQUE NO",
    "POLICY HOLDER",
    "UW YEAR",
    "PAYEE"
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

function calculatePerclass(targetArray, valueUsedToCalculate) {
  let ValuesPerClass = {};
  let motorPrivate = [];
  let miscellaneous = [];
  let fireDomestic = [];
  let marine = [];
  let fireIndustrial = [];
  let liability = [];
  let motorCommercial = [];
  let personalAccident = [];
  let engineering = [];
  let theft = [];
  let workmensCompensation = [];
  let medical = [];
  let noCategory = [];
  let psv = [];

  targetArray.map(claim => {
    const insClass = claim.insuranceClass.trim().toString();
    let valueToPush = toFloat(claim[valueUsedToCalculate]);

    if (insClass === "MOTOR PRIVATE") {
      motorPrivate.push(valueToPush);
      // motor private
    } else if (insClass === "MISCELLANEOUS") {
      // MISCELLANEOUS
      miscellaneous.push(valueToPush);
    } else if (insClass === "FIRE DOMESTIC") {
      // FIRE DOMESTIC
      fireDomestic.push(valueToPush);
    } else if (insClass === "MARINE AND TRANSIT GOODS") {
      // MARINE
      marine.push(valueToPush);
    } else if (insClass === "FIRE INDUSTRIAL") {
      fireIndustrial.push(valueToPush);
      // FIRE INDUSTRIAL
    } else if (insClass === "LIABILITY") {
      liability.push(valueToPush);
      // LIABILITIES
    } else if (insClass === "MOTOR COMMERCIAL") {
      // MOTOR COMMERCIAL
      motorCommercial.push(valueToPush);
    } else if (insClass === "ENGINEERING") {
      // ENGINEERING
      engineering.push(valueToPush);
    } else if (insClass === "THEFT") {
      // THEFT
      theft.push(valueToPush);
    } else if (insClass === "PERSONAL ACCIDENT") {
      // PERSONAL ACCIDENT
      personalAccident.push(valueToPush);
    } else if (
      insClass === "WORKMENS COMPENSATION" ||
      insClass === "WORKMEN'S COMPENSATION" ||
      insClass === "WORKMEN COMPENSATION"
    ) {
      // WORKMENS COMPENSATION
      workmensCompensation.push(valueToPush);
    } else if (insClass === "MEDICAL") {
      // MEDICAL
      medical.push(valueToPush);
    } else if (insClass === "PSV") {
      // PSV
      psv.push(valueToPush);
    }else {
      // noCategory
      noCategory.push(valueToPush);
    }
  });

  ValuesPerClass.motorPrivate = _.sum(motorPrivate);
  ValuesPerClass.miscellaneous = _.sum(miscellaneous);
  ValuesPerClass.fireDomestic = _.sum(fireDomestic);
  ValuesPerClass.marine = _.sum(marine);
  ValuesPerClass.fireIndustrial = _.sum(fireIndustrial);
  ValuesPerClass.liability = _.sum(liability);
  ValuesPerClass.motorCommercial = _.sum(motorCommercial);
  ValuesPerClass.noCategory = _.sum(noCategory);
  ValuesPerClass.engineering = _.sum(engineering);
  ValuesPerClass.theft = _.sum(theft);
  ValuesPerClass.personalAccident = _.sum(personalAccident);
  ValuesPerClass.workmensCompensation = _.sum(workmensCompensation);
  ValuesPerClass.medical = _.sum(medical);
  ValuesPerClass.psv = _.sum(psv);

  ValuesPerClass.count = {
    motorPrivate: motorPrivate.length,
    miscellaneous: miscellaneous.length,
    fireDomestic: fireDomestic.length,
    marine: marine.length,
    fireIndustrial: fireIndustrial.length,
    liability: liability.length,
    personalAccident: personalAccident.length,
    motorCommercial: motorCommercial.length,
    noCategory: noCategory.length,
    engineering: engineering.length,
    theft: theft.length,
    workmensCompensation: workmensCompensation.length,
    medical: medical.length,
    psv: psv.length
  };

  return ValuesPerClass;
}

function getSummary(targetSheet, valueToRefer) {
  let summary = [];
  let summaryObj = calculatePerclass(targetSheet, valueToRefer);

  let sumTotal = 0;

  for (const prop in summaryObj) {
    if (prop !== "count") {
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

  summary = _.sortBy(summary, "CLASS");

  // move TOTAL SUM to last index
  summary.push(summary.splice(13, 1)[0]);

  return summary;
}

function runCalculationsFromIndex(sourcefile, reportDestination) {
  return new Promise((resolve, reject) => {
    let fileName = path.basename(sourcefile.toString(), ".xlsx");
    let finalReportName = `Final Report (${fileName}).xlsx`;
    let finalReportPath = `${reportDestination.toString()}/${finalReportName}`;

    let workbook = XLSX.readFile(sourcefile.toString());

    let sheetNameList = workbook.SheetNames;

    let sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
    let sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[1]]);
    let sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[2]]);
    let sheet4 = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList[3]]);

    let osBeginMonth = objectRenameKeys(sheet1, osBeginKeyChangesMap);
    let intimatedWithZeros = objectRenameKeys(sheet2, intimatedKeyChangesMap);
    let paymentWithDuplicate = objectRenameKeys(sheet3, paymentKeyChangesMap);
    let osEndMonth = objectRenameKeys(sheet4, osEndMonthKeyChangesMap);

    let claimData = {};

    //* ============== movement calcultions section ================* /
    let combinedOsWithDuplicates = _.concat(
      osEndMonth,
      osBeginMonth,
      "claimNo"
    );

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

    // sheet with those that appear in begining and End OS estimate
    let osNoChange = _.intersectionBy(osBeginMonth, osEndMonth, "claimNo");

    // sheet with those that appear in begining and End OS estimate : repeated
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

    // payments adjustments
    let uniquePaidClaimNo = _.uniqBy(paymentWithDuplicate, "claimNo").map(
      claim => claim.claimNo
    );

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
      .filter(claim => claim.paidAmount !== 0);

    // remove zero values claims from intimated
    let intimated = intimatedWithZeros.filter(
      claim => claim.intimationReserve !== 0
    );

    //============== BeginingEndMovement without Paid =============

    //  sheet with upward movement + differences
    let movedUpWithDifference = osRepeatedClaimNo
      .filter(function(claim) {
        return (
          toFloat(claim.osBeginEstimate) < toFloat(claim.osEndMonthEstimate)
        );
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
        return (
          toFloat(claim.osBeginEstimate) > toFloat(claim.osEndMonthEstimate)
        );
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
    let movementWithPaidUniqueClaimNo = _.intersectionBy(
      osRepeatedClaimNo,
      payments,
      "claimNo"
    ).map(claim => claim.claimNo);

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
      .filter(claim => claim.difference !== 0);

    //============== BeginingPaidEndMovement without End OS estimate =============

    // claims in Begining OS estimates and paid (payments) without paid amount data
    let beginPaidIncomplete = _.intersectionBy(
      osBeginMonth,
      payments,
      "claimNo"
    );

    console.log("-----------------------hiut ---------------");

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

              if (newObj.paidAmount !== undefined) {
                newObj.difference =
                  toFloat(newObj.paidAmount) - toFloat(newObj.osBeginEstimate);
              }
            }
          }
        });

        return newObj;
      })
      .filter(claim => claim.difference !== 0);

    // sheet with Begining OS estimates and paid (payments) movements that are not In OsEndmonth
    let beginPaidMovementNoEndOS = _.differenceBy(
      beginPaid,
      osEndMonth,
      "claimNo"
    );

    // ============== intimation related movements =============

    // claimsNo for claims in intimated and paid
    let intimatedAndPaidClaimNo = _.intersectionBy(
      intimated,
      payments,
      "claimNo"
    ).map(claim => claim.claimNo);

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
        claim => toFloat(claim.intimationReserve) !== toFloat(claim.paidAmount)
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

    let intimatedPaidMovementWithEndOs = _.intersectionBy(
      intimatedPaidNoDifference,
      osEndMonth,
      "claimNo"
    )
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

    console.log("-----------------------hiut 2---------------");
    // sheet with intimated - EndOS movement
    let intimatedEndOsMovement = intimatedEndOsIncomplete
      .map(function(claim) {
        let newObj = {};
        intimatedEndOsMovementDuplicates.map(function(combineClaim) {
          if (combineClaim.claimNo === claim.claimNo) {
            for (const prop in combineClaim) {
              newObj[prop] = combineClaim[prop];
              if (newObj.osEndMonthEstimate !== undefined) {
                newObj.difference =
                  toFloat(newObj.osEndMonthEstimate) -
                  toFloat(newObj.intimationReserve);
              }
            }
          }
        });

        return newObj;
      })
      .filter(claim => claim.difference !== 0);

    // intimated-End movement with no claims from paid

    let intimatedEndOsMovementNoPaid = _.differenceBy(
      intimatedEndOsMovement,
      payments,
      "claimNo"
    );

    // ================== closed As No claim  =============== //

    let intimatedNotPaidclaims = _.differenceBy(intimated, payments, "claimNo");

    // sheet for claims closed as having no claim From Intimated

    let intimatedClosedAsNoClaim = _.differenceBy(
      intimatedNotPaidclaims,
      osEndMonth,
      "claimNo"
    );

    console.log("-----------------------hiut 3---------------");

    // sheet for claims closed as having no claim From Begining OS
    let beginingClosedAsNoClaim = _.differenceBy(
      removedOsBeginToEndMonth,
      payments,
      "claimNo"
    );

    let totalClosedAsNoClaim = _.concat(
      intimatedClosedAsNoClaim,
      beginingClosedAsNoClaim
    );

    // ================== Revived claim  =============== //

    let combinedIntimatedBegininingOs = _.concat(osBeginMonth, intimated);

    let EndOSNotInBeginIntimatedRevived = _.differenceBy(
      addedOsEndFromBeginMonth,
      intimated,
      "claimNo"
    );

    let paidNotInIntimatedBeginingOsRevived = _.differenceBy(
      payments,
      combinedIntimatedBegininingOs,
      "claimNo"
    );

    let revivedClaims = _.concat(
      EndOSNotInBeginIntimatedRevived,
      paidNotInIntimatedBeginingOsRevived
    );

    let totalMovementWithDuplicates = _.concat(
      beginEndmovementWithPaid,
      beginEndMovementNoPaid,
      beginPaidMovementNoEndOS,
      intimatedPaidmovementNoEndOS,
      intimatedPaidMovementWithEndOs,
      intimatedEndOsMovementNoPaid
    );

    let totalMovementUniqueClaimNo = _.uniqBy(
      totalMovementWithDuplicates,
      "claimNo"
    ).map(claim => claim.claimNo);

    let totalMovement = totalMovementUniqueClaimNo
      .map(function(claimNo) {
        let totaldifference = 0;
        let newObj = {};
        totalMovementWithDuplicates.map(function(dupClaim) {
          for (const prop in dupClaim) {
            if (claimNo === dupClaim.claimNo) {
              if (prop === "difference") {
                totaldifference =
                  totaldifference + toFloat(dupClaim.difference);
                newObj.difference = totaldifference;
              } else {
                newObj[prop] = dupClaim[prop];
              }
            }
          }
        });

        return newObj;
      })
      .filter(claim => claim.difference !== 0);

    let combinedAllDuplicate = _.concat(
      osBeginMonth,
      intimated,
      payments,
      osEndMonth
    );

    let combinedAllClaimUniqueClaimNo = _.uniqBy(
      combinedAllDuplicate,
      "claimNo"
    ).map(claim => claim.claimNo);

    let combinedAll = combinedAllClaimUniqueClaimNo.map(function(claimNo) {
      let newObj = {};

      combinedAllDuplicate.map(function(dupClaim) {
        for (const prop in dupClaim) {
          if (claimNo === dupClaim.claimNo) {
            newObj[prop] = dupClaim[prop];
          }
        }
      });

      return newObj;
    });

    console.log("-----------------------hiut 4---------------");


    let paidSettled = combinedAll
      .map(function(claim) {
        let newObj = {};

        for (const prop in claim) {
          newObj[prop] = claim[prop];
        }

        if (
          claim.paidAmount !== undefined &&
          claim.osEndMonthEstimate === undefined
        ) {
          newObj.settled = "YES";
        } else {
          newObj.settled = "NO";
        }

        return newObj;
      })
      .filter(claim => claim.settled === "YES");

    let combinedAllwithLabels = combinedAll.map(function(claim) {
      let newObj = {};

      // movement flag

      if (totalMovement.some(e => e.claimNo === claim.claimNo)) {
        newObj.movement = "YES";

        newObj.difference = _.find(totalMovement, function(o) {
          return o.claimNo === claim.claimNo;
        }).difference;
      } else {
        newObj.movement = "NO";
      }
      // totalClosedAsNoClaim flag
      totalClosedAsNoClaim.some(e => e.claimNo === claim.claimNo)
        ? (newObj.noClaim = "YES")
        : (claim.noClaim = "NO");

      // revived claim flag

      revivedClaims.some(e => e.claimNo === claim.claimNo)
        ? (newObj.revived = "YES")
        : (claim.revived = "NO");

      // settled claim flag

      paidSettled.some(e => e.claimNo === claim.claimNo)
        ? (newObj.settled = "YES")
        : (claim.settled = "NO");

      for (const prop in claim) {
        newObj[prop] = claim[prop];
      }

      return newObj;
    });


    console.log("-----------------------hiut 5 ---------------");

    // summary sheets

    let oSbeginMonthSummary = getSummary(osBeginMonth, "osBeginEstimate");

    let oSEndMonthSummary = getSummary(osEndMonth, "osEndMonthEstimate");

    let totalMovementSummary = getSummary(totalMovement, "difference");

    let paidSummary = getSummary(payments, "paidAmount");

    let intimatedSummary = getSummary(intimated, "intimationReserve");

    let intimatedClosedAsNoClaimSummary = getSummary(
      intimatedClosedAsNoClaim,
      "intimationReserve"
    );

    let beginingClosedAsNoClaimSummary = getSummary(
      beginingClosedAsNoClaim,
      "osBeginEstimate"
    );

    let paidNotInIntimatedBeginingOsRevivedSummary = getSummary(
      paidNotInIntimatedBeginingOsRevived,
      "paidAmount"
    );

    console.log("-----------------------hiut 6 ---------------");

    let EndOSNotInBeginIntimatedRevivedSummary = getSummary(
      EndOSNotInBeginIntimatedRevived,
      "osEndMonthEstimate"
    );

    let totalClosedAsNoClaimSummary = intimatedClosedAsNoClaimSummary.map(
      intClosedClaim => {
        let newObj = {};

        let totalCount = intClosedClaim.COUNT;
        let total = toFloat(intClosedClaim.TOTAL);

        beginingClosedAsNoClaimSummary.map(function(beginingClosedClaim) {
          if (intClosedClaim.CLASS === beginingClosedClaim.CLASS) {
            totalCount = totalCount + beginingClosedClaim.COUNT;

            total = total + toFloat(beginingClosedClaim.TOTAL);

            newObj = {
              CLASS: beginingClosedClaim.CLASS,
              COUNT: totalCount,
              TOTAL: total.toFixed(2)
            };
          }
        });

        return newObj;
      }
    );

    let totalRevivedSummary = paidNotInIntimatedBeginingOsRevivedSummary.map(
      function(inPaidClaim) {
        let newObj = {};

        let totalCount = inPaidClaim.COUNT;
        let total = toFloat(inPaidClaim.TOTAL);

        EndOSNotInBeginIntimatedRevivedSummary.map(function(endOsClaim) {
          if (inPaidClaim.CLASS === endOsClaim.CLASS) {
            totalCount = totalCount + endOsClaim.COUNT;

            total = total + toFloat(endOsClaim.TOTAL);

            newObj = {
              CLASS: endOsClaim.CLASS,
              COUNT: totalCount,
              TOTAL: total.toFixed(2)
            };
          }
        });

        return newObj;
      }
    );

    let paidSettledSummary = getSummary(paidSettled, "paidAmount");

    claimData.osBeginingClaims = osBeginMonth;
    claimData.osEndClaims = osEndMonth;
    claimData.intimatedClaims = intimatedWithZeros;
    claimData.paidClaims = paymentWithDuplicate;
    claimData.paidSettledClaims = paidSettled;
    claimData.closedAsNoClaims = totalClosedAsNoClaim;
    claimData.revived = revivedClaims;
    claimData.movement = totalMovement;
    claimData.allClaims = combinedAllwithLabels;
    claimData.osBeginingSummary = oSbeginMonthSummary;
    claimData.intimatedSummary = intimatedSummary;
    claimData.paidSummary = paidSummary;
    claimData.paidSettledSummary = paidSettledSummary;
    claimData.osEndsummary = oSEndMonthSummary;
    claimData.movementSummary = totalMovementSummary;
    claimData.closedAsNoClaimSummary = totalClosedAsNoClaimSummary;
    claimData.revivedSummary = totalRevivedSummary;
    claimData.outputFilePath = finalReportPath;
    claimData.outputFileName = finalReportName;

    // create work book
    let wb = XLSX.utils.book_new();

    // create sheetsNames
    wb.SheetNames.push(
      "OS BEGIN SUMMARY",
      "INTIMATED SUMMARY",
      "PAID SUMMARY",
      "PAID AND SETTLED SUMMARY",
      "OS END SUMMARY",
      "MOVEMENT SUMMARY",
      "CLOSED AS NO CLAIM SUMMARY",
      "REVIVED CLAIMS SUMMARY",
      "ALL COMBINED SORTED",
      "OS LAST MONTH CLAIMS",
      "PAID CLAIMS",
      "INTIMATED CLAIMS",
      "MOVEMENTS",
      "CLOSED AS NO CLAIM",
      "REVIVED CLAIMS"
    );

    console.log("paidSettledSummary", paidSettledSummary);

    let wsOsBeginMonthSummary = XLSX.utils.json_to_sheet(
      toExcelSheet(oSbeginMonthSummary),
      summaryHeader
    );

    let wsIntimatedSummary = XLSX.utils.json_to_sheet(
      intimatedSummary,
      summaryHeader
    );

    let wsPaidSummary = XLSX.utils.json_to_sheet(paidSummary, summaryHeader);

    let wsPaidSettledSummary = XLSX.utils.json_to_sheet(
      paidSettledSummary,
      summaryHeader
    );

    let wsOsEndMonthSummary = XLSX.utils.json_to_sheet(
      toExcelSheet(oSEndMonthSummary),
      summaryHeader
    );

    let wsMovementSummary = XLSX.utils.json_to_sheet(
      totalMovementSummary,
      summaryHeader
    );

    let wsClosedAsNoClaimSummary = XLSX.utils.json_to_sheet(
      totalClosedAsNoClaimSummary,
      summaryHeader
    );

    let wsRevivedClaimSummary = XLSX.utils.json_to_sheet(
      totalRevivedSummary,
      summaryHeader
    );

    let wsCombinedOs = XLSX.utils.json_to_sheet(
      toExcelSheet(combinedAllwithLabels),
      combinedAllHeader
    );

    let wsOsBeginMonth = XLSX.utils.json_to_sheet(toExcelSheet(osBeginMonth));

    let wsPayments = XLSX.utils.json_to_sheet(toExcelSheet(payments));

    let wsIntimatedClaims = XLSX.utils.json_to_sheet(toExcelSheet(intimated));

    let wsTotalMovement = XLSX.utils.json_to_sheet(toExcelSheet(totalMovement));

    let wsClosedASNoClaim = XLSX.utils.json_to_sheet(
      toExcelSheet(totalClosedAsNoClaim),
      closedAsnoClaimHeader
    );

    let wsRevivedOs = XLSX.utils.json_to_sheet(
      toExcelSheet(revivedClaims),
      revivedHeader
    );

    let wsheets = [
      wsOsBeginMonthSummary,
      wsIntimatedSummary,
      wsPaidSummary,
      wsPaidSettledSummary,
      wsOsEndMonthSummary,
      wsMovementSummary,
      wsClosedAsNoClaimSummary,
      wsRevivedClaimSummary,
      wsCombinedOs,
      wsOsBeginMonth,
      wsPayments,
      wsIntimatedClaims,
      wsTotalMovement,
      wsClosedASNoClaim,
      wsRevivedOs
    ];

    // formating of column widths

    wsheets.map(function(sheet, index) {
      if (index < 7) {
        sheet["!cols"] = [
          {
            wch: 20
          },
          {
            wch: 10
          },
          {
            wch: 20
          }
        ];
        sheet.A1.s = {
          patternType: "solid",
          fgColor: {
            theme: 8,
            tint: 0.3999755851924192
          },
          bgColor: {
            indexed: 64
          }
        };
      } else {
        sheet["!cols"] = [
          {
            wch: 40
          },
          {
            wch: 30
          },
          {
            wch: 30
          },
          {
            wch: 30
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          },
          {
            wch: 15
          }
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

    wb.Sheets["OS BEGIN SUMMARY"] = wsOsBeginMonthSummary;
    wb.Sheets["INTIMATED SUMMARY"] = wsIntimatedSummary;
    wb.Sheets["PAID SUMMARY"] = wsPaidSummary;
    wb.Sheets["PAID AND SETTLED SUMMARY"] = wsPaidSettledSummary;
    wb.Sheets["OS END SUMMARY"] = wsOsEndMonthSummary;
    wb.Sheets["MOVEMENT SUMMARY"] = wsMovementSummary;
    wb.Sheets["CLOSED AS NO CLAIM SUMMARY"] = wsClosedAsNoClaimSummary;
    wb.Sheets["REVIVED CLAIMS SUMMARY"] = wsRevivedClaimSummary;
    wb.Sheets["ALL COMBINED SORTED"] = wsCombinedOs;
    wb.Sheets["OS LAST MONTH CLAIMS"] = wsOsBeginMonth;
    wb.Sheets["PAID CLAIMS"] = wsPayments;
    wb.Sheets["INTIMATED CLAIMS"] = wsIntimatedClaims;
    wb.Sheets.MOVEMENTS = wsTotalMovement;
    wb.Sheets["CLOSED AS NO CLAIM"] = wsClosedASNoClaim;
    wb.Sheets["REVIVED CLAIMS"] = wsRevivedOs;

    console.log("hit before write");

    XLSX.write(wb, {
      bookType: "xlsx",
      type: "binary"
    });

    console.log("hit after write");

    del.sync(["./app/tmp/*.xlsx"]);

    XLSX.writeFile(wb, "./app/tmp/unstyledReport.xlsx");

    XlsxPopulate.fromFileAsync("./app/tmp/unstyledReport.xlsx")
      .then(function(otherWorkBook) {
        const sheets = otherWorkBook.sheets();

        sheets.map(function(sheet) {
          sheet.row(1).style({
            bold: true,
            fill: "ffff00",
            fontSize: 14
          });
        });

        otherWorkBook.toFileAsync(finalReportPath);

        console.log("data done");

        resolve(claimData);
      })
      .catch(err => {
        reject({
          error: err.message
        });
      });
  });
}

module.exports = {
  runCalculationsFromIndex: runCalculationsFromIndex
};
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


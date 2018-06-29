$(document).ready(function () {

    const storage = require("electron-json-storage");
    const {
        shell,
        clipboard
    } = require("electron");

    let claimData;

    function toClipboard(passedString) {

        return clipboard.writeText(passedString)

    }


    function titleCase(str) {
        return str.toLowerCase().split(" ").map(function (word) {
            return (word.charAt(0).toUpperCase() +
                word.slice(1));
        }).join(" ");
    }

    function notify() {
        let myNotification = new Notification("Final report", {
            body: `${claimData.outputFileName} geneted and stored in ${claimData.outputFilePath}`,
            icon: "../assets/images/excel.png",
            image: "../assets/images/excel.png",
        });
        myNotification.onclick = () => {
            shell.showItemInFolder(claimData.outputFilePath);
        };
    }


    $(".table").delegate("td", "click", function () {

        toClipboard($(this).html());
    });


    function initializeDataLoad() {

        claimData.osBeginingSummary.map((item) => {


            $(".osBegining-summary tbody").append(
                `
<tr>
<td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );

        });

        claimData.intimatedSummary.map(item => {

            $(".intimated-summary tbody").append(
                `
<tr>
<td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });

        claimData.paidSummary.map(item => {
            $(".paid-summary tbody").append(
                `
<tr>
<td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });
        claimData.paidSettledSummary.map(item => {

            $(".paid-settled-summary tbody").append(
                `
<tr>
    <td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });


        claimData.osEndsummary.map(item => {
            $(".os-End-summary tbody").append(
                `
<tr>
    <td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });


        claimData.movementSummary.map(item => {
            $(".revision-summary tbody").append(
                `
<tr>
    <td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });

        claimData.revivedSummary.map(item => {
            $(".revived-summary tbody").append(
                `
<tr>
    <td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });


        claimData.closedAsNoClaimSummary.map(item => {
            $(".closed-as-noclaim-summary tbody").append(
                `
<tr>
    <td class='class-name'>${titleCase(item.CLASS)}</td>
    <td>${item.COUNT}</td>
    <td>${item.TOTAL}</td>
</tr>`
            );
        });


    }

    storage.get("data", function (error, data) {
        if (error) {
            throw error;
        }
        claimData = data;

        // notify();         
        initializeDataLoad();

    });


});
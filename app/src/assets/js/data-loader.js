$(document).ready(function () {

    const storage = require("electron-json-storage");
    const {
        shell,
        clipboard
    } = require("electron");

    let claimData;

    function toClipboard(passedString) {

        return clipboard.writeText(passedString);

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




        /**
         *
         *
         * @param {*} dataset array containing data to be displayed
         * @param {*} className of the table name in the view to be populated
         */


        /*
        function populateTable(dataset, className) {
            var table = $(`.${className} tbody`);

            dataset.map(item => {

                var tableRow = table.append("<tr></tr>");

                for (const prop in item) {
  
                    if (item.hasOwnProperty(prop)) {

                        if (item[prop] === undefined) {

                            tableRow.append("<td></td>");


                        } else if (prop === "osBeginEstimate") {

                            //tableRow.append(`<td class="text-primary">${item[prop]}</td>`);

                        } else {

                            tableRow.append(`<td>${item[prop]}</td>`);

                        }



                    }


                }

            });


        }

        populateTable(claimData.osBeginingClaims, "os-begining-claims");

*/


    
    }

    storage.get("data", function (error, data) {
        if (error) {
            throw error;
        }
        claimData = data;

                
        initializeDataLoad();

    });


});
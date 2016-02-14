/// <reference path="../App.js" />
var GlobalVariables = {
    DataArray: [],
    UpdatedData: [],
};

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $(document).ready(function () {
                app.initialize();


                $("#readDataBtn").click(function (event) {
                    writeData();
                });

            });

        });
    };

    function writeData() {

        var root = 'https://jsonplaceholder.typicode.com';
        var myTable = new Office.TableData();

        $.support.cors = true;
        $.ajax({
            url: root + '/posts',
            type: "GET",
            async: false,
            success: function (data) {
                GlobalVariables.UpdatedData = [];
                GlobalVariables.DataArray = [];
                for (var i in data) {
                    var subarr = [];
                    subarr.push(data[i].id);
                    subarr.push(data[i].body);

                    GlobalVariables.DataArray.push(subarr);

                }


            },
            error: function (XMLHttpRequest, textStatus, errorThrown) {

            }
        });
        var myFormat = { fontStyle: "bold", width: "autoFit", borderColor: "purple" };



        myTable.headers = ["Key", "Value"];
        myTable.rows = GlobalVariables.DataArray;

        Office.context.document.setSelectedDataAsync(myTable, { coercionType: Office.CoercionType.Table, tableOptions: { alignHorizontal: "left", filterButton: false, borderStyle: "double", style: "TableStyleLight16" }, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit", height: "auto fit", wrapping: false } }] }, function (asyncResult) {
            if (asyncResult.status == "failed") {
                $("#updatediv").empty();
                $("#results").text("Status : " + asyncResult.error.message);
            } else {
                $("#updatediv").empty();
                $("#results").text("Status : Data get succesfully");
                var btn = $("<button>").appendTo($("#updatediv"));
                $(btn).attr("id", "updateBtn");
                $(btn).text("Update Selected Data");
                $(btn).on("click", readBoundData);
                $(btn).css("display", "block")

                var subdiv = $("<div>").appendTo($("#updatediv"));
                $(subdiv).attr("id", "updatecontent");
                $('<span id="updateresults style="display:block"> Status : </span>').appendTo($("#updatecontent"));

            }
        });

        //give id for the binding
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: 'myBinding' }, function (asyncResult) {
        });

    }

    function readBoundData() {
        Office.select("bindings#myBinding").getDataAsync({ coercionType: "table" }, function (asyncResult) {
            GlobalVariables.UpdatedData = [];
            if (asyncResult.status === "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            }
            else {
                $.grep(GlobalVariables.DataArray, function (elem) {
                    var x = $.grep(asyncResult.value.rows, function (item) {
                        if (elem[0] == item[0] && elem[1] != item[1]) {
                            elem[1] = item[1]
                            return true;
                        }
                    });
                    if (x.length > 0) {
                        GlobalVariables.UpdatedData.push(x);

                    }
                });
                $("#updatecontent").empty();


                if (GlobalVariables.UpdatedData.length > 0) {
                    $('<span id="updateresults" style="display:block"> Status : Update Start</span>').appendTo($("#updatecontent"));
                    GlobalVariables.UpdatedData.forEach(function (item) {
                        var p = $("<p>").appendTo($("#updatecontent"));
                        $(p).text("Key : " + item[0][0])

                        var p = $("<p>").appendTo($("#updatecontent"));
                        $(p).text("New Value : " + item[0][1])
                    });
                    $('<span id="updateresults1" style="display:block"> Status : Update End</span>').appendTo($("#updatecontent"));
                }
                else {
                    $('<span id="updateresults style="display:block"> Status : No Values Changed</span>').appendTo($("#updatecontent"));
                }

            }
        });
    }
})();
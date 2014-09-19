/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {

        var time;
        Office.context.document.bindings.addFromNamedItemAsync("Uhrzeit", "text", { id: 'timeOfDay', valueFormat: Office.ValueFormat.Formatted, filterType: "all" }, function (timeResult) {
            if (timeResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Uhrzeit> existiert nicht');
                return;
            }
        });

        var date;
        Office.context.document.bindings.addFromNamedItemAsync("Datum", "text", { id: 'date', valueFormat: Office.ValueFormat.Formatted, filterType: "all" }, function (dateResult) {
            if (dateResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Datum> existiert nicht');
                return;
            }
        });

        var name;
        Office.context.document.bindings.addFromNamedItemAsync("Name", "text", { id: 'name', valueFormat: "unformatted", filterType: "all" }, function (nameResult) {
            if (nameResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Name> existiert nicht');
                return;
            }
        });

        var schedule;
        Office.context.document.bindings.addFromNamedItemAsync("Tabelle", Office.CoercionType.Matrix, { id: 'schedule', valueFormat: "unformatted", filterType: "all" }, function (nameResult) {
            if (nameResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Tabelle> existiert nicht');
                return;
            }
        });

        var dates;
        Office.context.document.bindings.addFromNamedItemAsync("Daten", Office.CoercionType.Matrix, { id: 'dates', valueFormat: Office.ValueFormat.Formatted, filterType: "all" }, function (nameResult) {
            if (nameResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Daten> existiert nicht');
                return;
            }
        });

        var timeTable;
        Office.context.document.bindings.addFromNamedItemAsync("Uhrzeiten", Office.CoercionType.Matrix, { id: 'timeTable', valueFormat: Office.ValueFormat.Formatted, filterType: "all" }, function (nameResult) {
            if (nameResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Ein Feld mit den <Uhrzeiten> existiert nicht');
                return;

            }
        });

        Office.select("bindings#name").getDataAsync(function (result) {
            name = result.value;
        });

        Office.select("bindings#date").getDataAsync(function (result) {
            date = result.value;
        });

        Office.select("bindings#timeOfDay").getDataAsync(function (result) {
            time = result.value;
        });

        Office.select("bindings#dates").getDataAsync(function (result) {
            dates = result.value;
        });

        Office.select("bindings#timeTable").getDataAsync(function (result) {
            timeTable = result.value;
        });

        Office.select("bindings#name").getDataAsync(function (result) {
            name = result.value;
        });

        var column;
        for (column = 0; column < dates[0].length; column++) {
            if (dates[0][column] === date) {
                break;
            }
        }

        var row;
        for (row = 0; row < timeTable.length ; row++) {
            if (parseFloat(timeTable[row][0].toString()) + 0.00002 > parseFloat(time.toString()) && parseFloat(timeTable[row][0].toString()) + 0.00002 > parseFloat(time.toString())) {
                break;
            }
        }

        Office.select("bindings#schedule").getDataAsync(function (result) {
            if (result.value[row][column].toString() == "") {
                Office.select("bindings#schedule").setDataAsync([[name.toString()]], { startRow: row, startColumn: column });
            }
        });
    }
})();
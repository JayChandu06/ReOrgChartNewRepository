/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file contains all the common utility methods.
  
 * load / unload a div pop up window
 * make an ajax request - get/post
 * switch between full screen and normal mode – toggle Full
 * Open an URL
 * Show/Hide progress bar
 * Escape html codes 

These methods can be readably invoked from different places in the application with appropriate parameters to get the desired  result
  
***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

function toggleFull() {
    var elem = document.body;
    toggleFullScreen(elem);
}

function toggleFullScreen(elem) {
    // ## The below if statement seems to work better ## if ((document.fullScreenElement && document.fullScreenElement !== null) || (document.msfullscreenElement && document.msfullscreenElement !== null) || (!document.mozFullScreen && !document.webkitIsFullScreen)) {
    if ((document.fullScreenElement !== undefined && document.fullScreenElement === null) || (document.msFullscreenElement !== undefined && document.msFullscreenElement === null) || (document.mozFullScreen !== undefined && !document.mozFullScreen) || (document.webkitIsFullScreen !== undefined && !document.webkitIsFullScreen)) {
        if (elem.requestFullScreen) {
            elem.requestFullScreen();
        } else if (elem.mozRequestFullScreen) {
            elem.mozRequestFullScreen();
        } else if (elem.webkitRequestFullScreen) {
            elem.webkitRequestFullScreen(Element.ALLOW_KEYBOARD_INPUT);
        } else if (elem.msRequestFullscreen) {
            elem.msRequestFullscreen();
        }

        $('#fullscreen-false').hide();
        $('#fullscreen-true').show();

    } else {
        if (document.cancelFullScreen) {
            document.cancelFullScreen();
        } else if (document.mozCancelFullScreen) {
            document.mozCancelFullScreen();
        } else if (document.webkitCancelFullScreen) {
            document.webkitCancelFullScreen();
        } else if (document.msExitFullscreen) {
            document.msExitFullscreen();
        }

        $('#fullscreen-true').hide();
        $('#fullscreen-false').show();
    }
}

function doNothing(key, value) {
    return value;
}

function unloadPopupBox(popBoxID) {    // TO Unload the Popupbox
    $('#' + popBoxID).fadeOut("slow");
    cleanMessages();
}

function loadPopupBox(popBoxID) {    // To Load the Popupbox
    $('#' + popBoxID).fadeIn("slow");
}

function loadPage(PAGE_URL, DIV_ID) {

    $.ajax({
        url: PAGE_URL,
        async: true,
        success: function (res) {
            //alert(res);
            document.getElementById(DIV_ID).innerHTML = res;
        },
        error: function (err) {
            alert(err.message);
        }
    });

}

function postData(method, jsonData, responseDiv, responseHandler) {

    //kendo.ui.progress($(".loadingPanel"), true);

    ShowProgressBar(".loadingPanel");

    $.ajax({
        type: 'POST',
        url: method,
        data: "{'JSON':'" + jsonData + "'}",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        async: true,
        cache: false,
        success: function (msg) {
            
            HideProgressBar(".loadingPanel");

            if (responseHandler)
                responseHandler(msg.d);
            else
                $("#" + responseDiv).val(msg.d);

           kendo.ui.progress($(".loadingPanel"), false);

          
        },
        error: function (err) {
            HideProgressBar(".loadingPanel");
            alert(err.message);
        }
    });

}

function getData(method, parameters, responseDiv, responseHandler) {

    //kendo.ui.progress($(".loadingPanel"), true);
    
    ShowProgressBar(".loadingPanel");

    $.ajax({
        type: 'POST',
        url: method,
        data: parameters,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        async: true,
        cache: false,
        success: function (msg) {

            HideProgressBar(".loadingPanel");

            if (responseHandler)
                responseHandler(msg.d);
            else
                $("#" + responseDiv).text(msg.d);

            
            //kendo.ui.progress($(".loadingPanel"), false);

           

        },
        error: function (err) {
            HideProgressBar(".loadingPanel");
            alert(err.message);
        }
    });

}

function buildHtmlTableFromJson(jsonStr,tableSelector){

    var json = JSON.parse(jsonStr);

    $(tableSelector).empty();

    var columns = addAllColumnHeaders(json, tableSelector);

    for (var i = 0 ; i < json.length ; i++) {
        var row$ = $('<tr/>');
        for (var colIndex = 0 ; colIndex < columns.length ; colIndex++) {
            var cellValue = json[i][columns[colIndex]];

            if (cellValue == null) { cellValue = ""; }

            row$.append($('<td/>').html(cellValue));
        }
        $(tableSelector).append(row$);
    }
}

// Adds a header row to the table and returns the set of columns.
// Need to do union of keys from all records as some records may not contain
// all records
function addAllColumnHeaders(json, tableSelector) {
    var columnSet = [];
    var headerTr$ = $('<tr/>');

    for (var i = 0 ; i < json.length ; i++) {
        var rowHash = json[i];
        for (var key in rowHash) {
            if ($.inArray(key, columnSet) == -1) {
                columnSet.push(key);
                headerTr$.append($('<th/>').html(key));
            }
        }
    }
    $(tableSelector).append(headerTr$);
    return columnSet;
}

function openExcel(filePath) {
    window.open(filePath, "_new");
}

function ShowPopuMessage(jsonStr) {
    buildHtmlTableFromJson(jsonStr, "#actionmessagetable")
    loadPopupBox("ACTIONMESSAGE");
}

function ShowPopuMessage2(message) {
    $("#actionmessagediv").text(message);
    loadPopupBox("ACTIONMESSAGE");
}

function ShowProgressBar(selector) {
    kendo.ui.progress($(selector), true);
}

function HideProgressBar(selector) {
    kendo.ui.progress($(selector), false);
}

var entityMap1 = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': '&quot;',
    "'": '&#39;',
    "'": '&#39;',
    "(": '&#40;',
    ")": '&#41;',
    "\\": '&#92;',
    "/": '&#x2F;'
};

function escapeHtml1(string) {
    return String(string).replace(/[&<>"'\/]/g, function (s) {
        return entityMap1[s];
    });
}


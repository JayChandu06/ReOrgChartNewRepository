/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the helper methods initiative various activities upon a action is taken by the user on scree.

 * Add Initiative, Population, Version etc. 
 * Save Data
 * Load Orgchart Data
 * Load lookup Data
 * Delete 
 * Download Excel 
 * etc

 These methods inturn makes one or more asynchronous request to the backend server to get/post data and refresh the user interface 
 asynchronously.

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/


function SaveInitiative() {
    if (BlankString($("#initiativetitle").val())) {
        alert(MESSAGE.VALIDATE_BLANK_INITIATIVE_TITLE);
        return false;
    }
    
    if (OptionExists("#INITIATIVE", LOOKUPNAME.INITIATIVE + " :: " + $("#initiativetitle").val().trim())) {
        alert(MESSAGE.VALIDATE_UNIQUE_INITIATIVE_TITLE);
        return false;
    }

    var countryID = 0;

    if ($("#txtOper").val() == CHART_TYPE.OPERATIONAL)
        countryID = 0;
    else {
        countryID = $("#COUNTRY").val();
        if (countryID == "-2") {
            alert(MESSAGE.VALIDATE_BLANK_BLANK_COUNTRY);
            return false;
        }
    }
    var method = "OrgChart.aspx/AddNewInitiative";
    var data = '{"Title":"' + $("#initiativetitle").val() + '","Description":"' + $("#initiativedescription").val() + '","ParentId":"'+countryID +'"}';
    var responseDivId = "respInitiative";

    var responseHandler = function (respJSON) {

        var obj = JSON.parse(respJSON, doNothing);
        if (typeof obj[0].STATUS !== "undefined" && obj[0].STATUS == STATUS.TRUE) {
            loadLookup(LOOKUPNAME.INITIATIVE, countryID);
            //ShowPopuMessage(respJSON);
        } else {
            //ShowPopuMessage(respJSON);
        }

        buildHtmlTableFromJson(respJSON, "#respInitiativeTable");
        //ShowPopuMessage(respJSON);
    }

    postData(method, data, responseDivId, responseHandler);

}
function AddNewVersion() {

    if ($("#VERSION").val() < 0) {
        alert(MESSAGE.VALIDATE_VERSION_SELECTION);
        return false;
    }

    if (OptionExists("#VERSION", LOOKUPNAME.VERSION + " :: " + $("#versiondescription").val().trim())) {
        alert(MESSAGE.VALIDATE_UNIQUE_VERSION_TITLE);
        return false;
    }
    

    var method = "OrgChart.aspx/AddNewVersion";
    var data = '{"ParentId":"' + $("#VERSION").val() + '","Description":"' + $("#versiondescription").val() + '"}';
    var responseDivId = "responseAddVersion";

    var responseHandler = function (respJSON) {
        buildHtmlTableFromJson(respJSON,"#responseAddVersionTable");
        //do something
        //lookupChangeHandler(LOOKUPNAME.VERSION, $("#" + LOOKUPNAME.POPULATION).val(), '', $("#" + LOOKUPNAME.USER).val(), '');
        loadLookup(LOOKUPNAME.VERSION, $("#" + LOOKUPNAME.POPULATION).val(), $("#" + LOOKUPNAME.USER).val(), $("#" + LOOKUPNAME.VERSION).val());
    }

    postData(method, data, responseDivId, responseHandler);

}
function FinalizeVersion() {

    if ($("#VERSION").val() < 0) {
        alert(MESSAGE.VALIDATE_VERSION_SELECTION);
        return false;
    }

    /*
    if (OptionExists("#VERSION", $("#versiondescription").val())) {
        alert(MESSAGE.VALIDATE_UNIQUE_VERSION_TITLE);
        return false;
    }
    */

    var method = "OrgChart.aspx/AddFinalVersion";
    var data = '{"ParentId":"' + $("#VERSION").val() + '","Description":"' + $("#versiondescription").val() + '"}';
    var responseDivId = "responseAddVersion";

    var responseHandler = function (respJSON) {
        ShowPopuMessage(respJSON);
        loadLookup(LOOKUPNAME.VERSION, $("#" + LOOKUPNAME.POPULATION).val(), $("#" + LOOKUPNAME.USER).val(), $("#" + LOOKUPNAME.VERSION).val());
    }

    postData(method, data, responseDivId, responseHandler);

}
function RelaodPopulations() {
    try {
        loadLookup(LOOKUPNAME.POPULATION,$("#"+LOOKUPNAME.INITIATIVE).val());
    }
    catch (ex) {
        alert(ex.message);
    }
}
function SavePopulation() {
     // this function is not in use
    if (BlankString($("#initiativetitle").val())) {
        alert(MESSAGE.VALIDATE_BLANK_POPULATION_TITLE);
        return false;
    }

    if (OptionExists("#POPULATION", LOOKUPNAME.POPULATION + " :: " + $("#populationtitle").val().trim())) {
        alert(MESSAGE.VALIDATE_BLANK_POPULATION_TITLE);
        return false;
    }

    var method = "OrgChart.aspx/AddNewPopulation";
    var data = '{"InitiativeTitle":"' + $("#initiativetitle").val() + '","InitiativeDescription":"' + $("#initiativedescription").val() + '"}';
    var responseDivId = "respInitiative";

    var responseHandler = function (respJSON) {
        //do something

    }

    postData(method, data, responseDivId, responseHandler);

}
function lookupChangeHandler(name, filter1, filter1Val, filter2, filter2Val) {

    if (filter1 == '-1') {
        loadPopupBox(filter1Val);
        return true;
    }

    var selectedVal = filter2Val;
    

    loadLookup(name, filter1, filter2, selectedVal);
}
function initLookups() {
    //alert($("#txtInitiative").val());

    var chart_type = $("#txtOper").val();
    if (BlankString(chart_type)) { chart_type = CHART_TYPE.OPERATIONAL; }

    var countryVal = $("#txtCountry").val();
    if (BlankString(countryVal)) { countryVal = '-2'; }

    var initiativeVal = $("#txtInitiative").val();
    if (BlankString(initiativeVal)) { initiativeVal = '-2'; }

    var populationVal = $("#txtPopulation").val();
    if (BlankString(initiativeVal)) { populationVal = '-2'; }

    var userVal = $("#txtUser").val();
    if (BlankString(userVal)) { userVal = '-2'; }

    var versionVal = $("#txtVersion").val();
    if (BlankString(versionVal)) { versionVal = '-2'; }


    if (chart_type != CHART_TYPE.OPERATIONAL)
        loadLookup(LOOKUPNAME.COUNTRY, "", "", countryVal);
    else
        hideCountryDropDwon();

    if (chart_type == CHART_TYPE.LEGAL && countryVal == '-2') {
        cleanLookup(LOOKUPNAME.INITIATIVE);
    } else {

        if (chart_type == CHART_TYPE.OPERATIONAL) {
            countryVal = "0";
        }

        loadLookup(LOOKUPNAME.INITIATIVE, countryVal, "", initiativeVal);
    }

    if (initiativeVal != '-2')
        loadLookup(LOOKUPNAME.POPULATION, initiativeVal, "", populationVal);
    else
        cleanLookup(LOOKUPNAME.POPULATION);


    if (populationVal != '-2')
        loadLookup(LOOKUPNAME.USER, populationVal, "", userVal);
    else
        cleanLookup(LOOKUPNAME.USER);

    if (userVal != '-2' && userVal != 'SYS')
        loadLookup(LOOKUPNAME.VERSION, populationVal, userVal, versionVal);
    else
        cleanLookup(LOOKUPNAME.VERSION);
}
function loadLookup(name, filter1, filter2, selectedVal) {

    // alert(name);
    var method = "OrgChart.aspx/GetLookupValue";
    var data = '{"LookupName":"' + name + '","Filter1":"' + filter1 + '","Filter2":"' + filter2 + '"}';
    var responseDivId = "respLookup";

    var responseHandler = function (jsonStr) {
        // populate the look up drop down
        var json = JSON.parse(jsonStr);
        /*
        try{
            if ($("#txtOper").val() == CHART_TYPE.OPERATIONAL) unLoad("emptyModelOperational");
            else unLoad("emptyModelLegal");
        }catch(e){
            //alert(e);
        }
        */

        var currentSelectedVal = $('#' + name).val();

        $('#' + name).empty();

        if (LOOKUPNAME.VERSION != name)
            $('#' + name).append($('<option >').text("SELECT " + name).attr('value', "-2"));

        if (LOOKUPNAME.VERSION == name) // && json.length == 0
            $('#' + name).append($('<option >').text(name + " :: " + "HR Core Version").attr('value', "-2"));

        $.each(json, function (i, obj) {
           $('#'+name).append($('<option>').text(name + " :: "+obj.TITLE).attr('value', obj.ID));
        });

        if(LOOKUPNAME.INITIATIVE == name || LOOKUPNAME.POPULATION == name) 
            $('#' + name).append($('<option>').text("ADD" + name).attr('value', "-1"));

        
        if (LOOKUPNAME.USER == name) {
            if (!OptionExists("#" + name, name + " :: " + LoggedInUser))
                $('#' + name).append($('<option>').text(name + " :: " + LoggedInUser).attr('value', LoggedInUser));
        }
        
        if (selectedVal && selectedVal != "" && selectedVal != currentSelectedVal) {
            $("#" + name).val(selectedVal).trigger("change");
        }

        /*
        if (LOOKUPNAME.VERSION == name)
            onVersionChange();

        if (LOOKUPNAME.INITIATIVE == name && $("#" + LOOKUPNAME.COUNTRY).val() != "-2")
            onCountryChange();
        */

        setSelectionPath();
    }
        
    getData(method, data, responseDivId, responseHandler);

}
function cleanLookup(name) {

    $('#' + name).empty();

    
    if (LOOKUPNAME.VERSION != name && LOOKUPNAME.USER != name) {
        $('#' + name).append($('<option >').text("SELECT " + name).attr('value', "-2").attr('selected', 'selected'));
       
        $("#" + name).val("-2").trigger("change");
        
    }

    if (LOOKUPNAME.USER == name) {
        if (!OptionExists("#" + name, name + " :: " + LoggedInUser)) {
            $('#' + name).append($('<option>').text(name + " :: " + LoggedInUser).attr('value', LoggedInUser).attr('selected', 'selected'));

            $("#" + name).val(LoggedInUser).trigger("change");
        }
    }

    if (LOOKUPNAME.VERSION == name)
    {
        if ($("#txtOper").val() == CHART_TYPE.OPERATIONAL) {
            $('#' + name).append($('<option >').text(name + " :: " + "HR Core Version").attr('value', "-2").attr('selected', 'selected'));
            $("#" + name).val("-2").trigger("change");
            // onVersionChange();
        }
        else
        {
            $('#' + name).append($('<option >').text("SELECT " + name).attr('value', "-2").attr('selected', 'selected'));
            $("#" + name).val("-2").trigger("change");

            //unLoad("emptyModelLegal");
            //onCountryChange();
        }
    }
}
function onCountryChange() {
    $("#txtCountry").val($("#COUNTRY").val());
    $("#txtShowLevel").val("");

    if ($("#COUNTRY").val() == '-2') {
        showMap(true);
    }
    else {
        showMap(false);
        reloadTreeOnDemand();
        initOrghartData();
    }
    //loadOrgChartData();
}
function onVersionChange() {
    $("#txtVersion").val($("#VERSION").val());
    setSelectionPath();

    // SET the start position to the first position in the version by setting the showlevel as blank
    $("#txtShowLevel").val("");
    loadOrgChartData();
}
function onInitiativeChage() {
    $("#selectedInitiative").text($("#INITIATIVE option:selected").text());
}
function setSelectionPath() {

    // for copy version popup
    $("#selectedVersion").text(
              $("#INITIATIVE option:selected").text()
            + " > "
            + $("#POPULATION option:selected").text()
            + " > "
            + $("#USER option:selected").text()
            + " > "
            + $("#VERSION option:selected").text()
        );
    // for download popup
    $("#dwonloadpath").text(
              $("#INITIATIVE option:selected").text()
            + " > "
            + $("#POPULATION option:selected").text()
            + " > "
            + $("#USER option:selected").text()
            + " > "
            + $("#VERSION option:selected").text()
        );
    
}
function openPopulationWindow() {
    try {
        loadPopupBox(LOOKUPNAME.POPULATION);
        document.getElementById('uploadexcel').contentWindow.setInitiative($("#" + LOOKUPNAME.INITIATIVE).val());
    } catch (e) {
    }
}
function BreadcrumbClicked(positionid) {
    if (positionid) {
        if (positionid == '10000000')
            $("#COUNTRY").val("-2").trigger("change");

        $("#txtShowLevel").val(positionid);
        loadOrgChartData();
    }
}
function showPreview(val) {
    $(".showPreview").each(function () { $(this).removeClass("showPreview"); $(this).addClass("hidePreview"); })
    $("#" + val).removeClass("hidePreview"); $("#" + val).addClass("showPreview");
}
function DownloadVersionData() {

    if ($("#VERSION").val() < 0) {
        alert(MESSAGE.VALIDATE_VERSION_SELECTION_DOWNLOAD);
        return false;
    }

    var method = "Excel.aspx/DownloadVersion";
    var data = '{"Oper":"' + $("#txtOper").val() + '","Id":"' + $("#VERSION").val() + '","Description":"' + $("#VERSION option:selected").text().split(" :: ")[1] + '"}';
    var responseDivId = "downloadMessageDiv";

    var responseHandler = function (jsonStr) {
        // populate the look up drop down
        var json = JSON.parse(jsonStr);
        HideProgressBar("#downloadMessageDiv");
        
        if (typeof json[0].STATUS !== "undefined" && json[0].STATUS == STATUS.TRUE) {
            openExcel(json[0].MESSAGE);
            buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
        } else {
            //ShowPopuMessage(jsonStr);
            buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
        }

    }
    ShowProgressBar("#downloadMessageDiv");
    getData(method, data, responseDivId, responseHandler);
}

//NAGARVI5--
function DownloadPDFVersionData() {


    var method = "Pdf.aspx/DownloadPDFVersion";



    // var method = "OrgChart.aspx/DownloadPDFVersion";
    var data = "{'Country':'" + $("#txtCountry").val() + "','ShowLevel':'" + $("#txtShowLevel").val() + "','Levels':'" + $("#txtLevelFlag").val() + "','Oper':'" + $("#txtOper").val() + "','Version':'" + $("#txtVersion").val() + "'}";
    var responseDivId = "downloadMessageDiv";

    // populate the look up drop down
    // var json = JSON.parse(jsonStr);
    //  HideProgressBar("#downloadMessageDiv");
    $.ajax({
        type: 'POST',
        url: method,
        data: data,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        async: true,
        cache: false,
        success: function (msg) {

            var win = window.open(msg.d, '_blank');
            win.focus();

            // window.location = msg.d;

        }
    });



    //if (typeof json[0].STATUS !== "undefined" && json[0].STATUS == STATUS.TRUE) {
    //    openExcel(json[0].MESSAGE);
    //    buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
    //} else {
    //    //ShowPopuMessage(jsonStr);
    //    buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
    //}
}




function DownloadInitiativeData() {

    if ($("#INITIATIVE").val() < 0) {
        alert(MESSAGE.VALIDATE_INITIATIVE_SELECTION_DOWNLOAD);
        return false;
    }

    var method = "Excel.aspx/DownloadInitiative";
    var data = '{"Oper":"' + $("#txtOper").val() + '","Id":"' + $("#INITIATIVE").val() + '","Description":"' + $("#INITIATIVE option:selected").text().split(" :: ")[1] + '"}';
    var responseDivId = "downloadMessageDiv";

    var responseHandler = function (jsonStr) {
        // populate the look up drop down
        var json = JSON.parse(jsonStr);
        HideProgressBar("#downloadMessageDiv");

        if (typeof json[0].STATUS !== "undefined" && json[0].STATUS == STATUS.TRUE) {
            openExcel(json[0].MESSAGE);
            buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
        } else {
            //ShowPopuMessage(jsonStr);
            
            buildHtmlTableFromJson(jsonStr, "#downloadMessageTable");
        }

    }
    ShowProgressBar("#downloadMessageDiv");
    getData(method, data, responseDivId, responseHandler);
}
function DeactivatePopulation() {

    if ($("#POPULATION").val() < 0) {
        alert(MESSAGE.VALIDATE_POPULATION_SELECTION_DELETE);
        return false;
    }

    if (!confirm(MESSAGE.CONFIRM_POPULATION_DELETION))
        return false;

    var method = "Orgchart.aspx/DeactivatePopulation";
    var data = '{"Id":"' + $("#POPULATION").val() + '"}';
    var responseDivId = "respDownloadPopup";

    var responseHandler = function (jsonStr) {
        // populate the look up drop down
        var json = JSON.parse(jsonStr);


        if (typeof json[0].STATUS !== "undefined" && json[0].STATUS == STATUS.TRUE) {
            RelaodPopulations();
            loadLookup(LOOKUPNAME.USER, "-2", "", "");
            loadLookup(LOOKUPNAME.VERSION, "-2", "-2", "");
        } 

        ShowPopuMessage(jsonStr);
    }

    getData(method, data, responseDivId, responseHandler);
}
function DeactivateVersion() {

    if ($("#VERSION").val() < 0) {
        alert(MESSAGE.VALIDATE_VERSION_SELECTION);
        return false;
    }

    if (!confirm(MESSAGE.CONFIRM_VERSION_DELETION))
        return false;

    var method = "Orgchart.aspx/DeactivateVersion";
    var data = '{"Id":"' + $("#VERSION").val() + '"}';
    var responseDivId = "respDownloadPopup";

    var responseHandler = function (jsonStr) {
        // populate the look up drop down
        
        loadLookup(LOOKUPNAME.VERSION, $("#" + LOOKUPNAME.POPULATION).val(), $("#" + LOOKUPNAME.USER).val());
        
        ShowPopuMessage(jsonStr);
    }

    getData(method, data, responseDivId, responseHandler);
}
function initOrghartData() {

    if (!ConfirmDataLoad()) return false;

    var method = "OrgChart.aspx/GetOrgChartData";
    var data = "{'Country':'" + $("#txtCountry").val() + "','ShowLevel':'" + $("#txtShowLevel").val() + "','Levels':'" + $("#txtLevelFlag").val() + "','Oper':'" + $("#txtOper").val() + "','Version':'" + $("#txtVersion").val() + "'}";
    var responseDivId = "respLookup";

    var responseHandler = "";

    responseHandler = function (jsonStr) {
        GenerateBreadCrumb($("#txtShowLevel").val(), "#workspaceBreadCrumb");
        GenerateBreadCrumb($("#txtShowLevel").val(), "#palleteBreadCrumb");
        reloadRemote("{ \"class\": \"go.TreeModel\",  \"nodeDataArray\":" + jsonStr + " }");
    }

    if ($("#txtOper").val() == CHART_TYPE.LEGAL && $("#txtCountry").val() == '-2') {
        responseHandler($("#emptyModelLegal").val());
        //showmap();
    }else{
        getData(method, data, responseDivId, responseHandler);
    }
}
function loadOrgChartData() {

    if (!ConfirmDataLoad()) return false;

    if ($("#txtShowLevel").val() == '10000000')
        $("#COUNTRY").val("-2").trigger("change");

    var method = "home/GetOrgChartData";
    var data = "{'Country':'" + $("#txtCountry").val() + "','ShowLevel':'" + $("#txtShowLevel").val() + "','Levels':'" + $("#txtLevelFlag").val() + "','Oper':'" + $("#txtOper").val() + "','Version':'" + $("#txtVersion").val() + "'}";
    var responseDivId = "respLookup";

    var responseHandler = "";
    if (splitScreen == SPLIT_SCREEN.TRUE && lefttreerefreshed == 0) {

        
        responseHandler = function (jsonStr) {
            
                                                    GenerateBreadCrumb($("#txtShowLevel").val(), "#palleteBreadCrumb");
                                                    loadPallete("{ \"class\": \"go.TreeModel\",  \"nodeDataArray\":" + jsonStr + " }");
                                              }
    } else {

        
        responseHandler = function (jsonStr) {
            
                                                    GenerateBreadCrumb($("#txtShowLevel").val(), "#workspaceBreadCrumb");
                                                    load("{ \"class\": \"go.TreeModel\",  \"nodeDataArray\":" + jsonStr + " }");
                                              }
    }

    getData(method, data, responseDivId, responseHandler);
}

function saveOrgChartData(DataElementSelector) {

    if (!DataElementSelector) return false;

    var method = "OrgChart.aspx/UpdateOrgChartData";
    var data = "{'Oper':'" + $("#txtOper").val() + "','" + LOOKUPNAME.VERSION + "':'" + $("#" + LOOKUPNAME.VERSION).val() + "','JSON':'" + escapeHtml1($(DataElementSelector).val()) + "'}";
    var responseDivId = "actionmessagetableholder";

    var responseHandler = function (respJSON) {

        var obj = JSON.parse(respJSON, doNothing);

        if (typeof obj[0].STATUS !== "undefined"  && obj[0].STATUS == STATUS.TRUE) {
            modified = MODIFIED_FLEG.FALSE;
            loadOrgChartData();
            //ShowPopuMessage(respJSON);
        } else {
            ShowPopuMessage(respJSON);
        }

    }

    getData(method, data, responseDivId, responseHandler);
    
}
function GenerateBreadCrumb(position,containerSelector) {

    //if (!position) return false;
    if (!containerSelector) return false;

    var method = "OrgChart.aspx/GetBreadcrumb";
    var data = "{'Oper':'" + $("#txtOper").val() + "','" + LOOKUPNAME.VERSION + "':'" + $("#" + LOOKUPNAME.VERSION).val() + "','" + FIELD.POSITIONID + "':'" + position + "'}";
    var responseDivId = containerSelector;

    var responseHandler = function (respJSON) {

        var obj = JSON.parse(respJSON, doNothing);

        if (obj[0] && typeof obj[0].STATUS !== "undefined") {
            // Some error
            if( obj[0].STATUS == STATUS.FALSE )
                ShowPopuMessage(respJSON);
        } else {
            BuildBreadCrumbFromJSON(respJSON, responseDivId);
        }

    }

    getData(method, data, responseDivId, responseHandler);

}
function BuildBreadCrumbFromJSON(jsonStr,containerSelector) {

    var array = JSON.parse(jsonStr, doNothing);
    var displayDiv = $(containerSelector);
    // slider div
    var lasElm = $(containerSelector + " div:last-child").detach();

    displayDiv.empty();

    var hrefEle = document.createElement("a");
    hrefEle.setAttribute("href", "#");
    hrefEle.innerHTML = "<b>Location:</b>";
    displayDiv.append(hrefEle);
    //alert(array.length);
    for (var i = array.length - 1; i > -1; i--) {
        var hrefEle = document.createElement("a");
        hrefEle.setAttribute("href", "javascript:BreadcrumbClicked('" + array[i].ID+"')");
            hrefEle.setAttribute("class", "breadcrumbs-item");
            hrefEle.innerHTML = array[i].TITLE;
          //  hrefEle.onclick = function () { BreadcrumbClicked(array[i].ID); };
                
            displayDiv.append(hrefEle);
    }
    //displayDiv.append(lasElm);
}
function BlankString(theString) {

    if (!theString) {
        return true;
    }

    if (typeof theString !== "undefined" && theString !== null) {
        return false;
    } else {
        return true;
    }


}
function OptionExists(dropdownselector, optionText) {
    var exists = false;
    var selText = $.trim(optionText);
    $(dropdownselector + '  option').each(function () {
        if ($(this).text() == selText) {
            exists = true;
        }
    });
    return exists;
}
function SetSelectOption(dropdownselector, optionText) {
    var exists = false;
    var selText = $.trim(optionText);
    $(dropdownselector + '  option').each(function () {
        if ($(this).text() == selText) {
            exists = true;
            $(dropdownselector).val($(this).val()).trigger("change");
            return true;
        }
    });
    return exists;
}
function ConfirmDataLoad() {
    if (modified == MODIFIED_FLEG.TRUE) {
        if (confirm(MESSAGE.CONFIRM_MOVE_OUT))
            return true;
        else
            return false;
    }
    return true;
}
function UserHasAccess() {

    var selectedVersionUser = $("#USER").val();
    var selectedVersionText = $("#VERSION option:selected").text().split(" :: ")[1]; 
    if (selectedVersionUser == LoggedInUser && selectedVersionText != KEY.INITIALVERSIONTEXT && selectedVersionText !=KEY.HRVERSIONTEXT)
        return true;

    return false;

}
function SharePageURL() {

    var URL = window.location.href;
    // remove any parameter
    URL = URL.split("?")[0];

    // Add required parameters to form the bookmarkable URL

    // VersionID 
    // PopulationID
    // UserID
    // PositionID
    // Depth

    URL = URL + "?INITIATIVE=" + $("#INITIATIVE").val() +
                "&POPULATION=" + $("#POPULATION").val() +
                "&VERSION=" + $("#VERSION").val() +
                "&USER=" + $("#USER").val() +
                "&POSITIONID=" + $("#txtShowLevel").val() +
                "&LEVELS=" + $("#slider1").val();

    $("#pagelink").val(URL);

    loadPopupBox('SHARELINK');

    }
function ShareByEmail() {
    window.location = "mailto:?body=" + $("#pagelink").val();
}
function CopyToClipboard() {
    var text = document.getElementById("pagelink").value;

    copyToClipboard(text);
    /*
    Copied = text.createTextRange();
    Copied.execCommand("Copy");
    */
}
function toggleSplitScreen() {

    var currentVal = $('input[name=splitscreen]:checked', '#settingsform').val();
    
    if (currentVal == 0) {
        $("[name=splitscreen]").val(["1"]);

    } else {
        $("[name=splitscreen]").val(["0"]);
    }
         
    reload();
}
function copyToClipboard(s)
{
    if( window.clipboardData && clipboardData.setData )
    {
        clipboardData.setData("Text", s);
    }
    else
    {
      	   
    }
}
function cleanMessages() {

    if ($("#respInitiativeTable")) $("#respInitiativeTable").empty();
    if ($("#responseTable")) $("#responseTable").empty();
    if ($("#responseAddVersionTable")) $("#responseAddVersionTable").empty();
    if ($("#downloadMessageTable")) $("#responseAddVersionTable").empty();
    

    if ($("#initiativetitle")) $("#initiativetitle").val(" ");
    if ($("#initiativedescription")) $("#initiativedescription").val(" ");
    if ($("#versiondescription")) $("#versiondescription").val(" ");

    if ($("#actionmessagediv")) $("#actionmessagediv").text(" ");
    

}
function hideCountryDropDwon() {
    $("#select-div-contry").hide();
}
function showCountryDropDown() {
    $("#select-div-contry").show();
}
function toggleChartType(chart_type) {

    if (chart_type == CHART_TYPE.LEGAL) {
        showCountryDropDown();
        showMap(true);
    } else {
        hideCountryDropDwon();
        showMap(false);
      //  loadLookup(LOOKUPNAME.COUNTRY);
    }

    $("#txtOper").val(chart_type);
    
    initLookups();

    //reload the left tree based on the chart type
    reloadTreeOnDemand();

    initOrghartData();
}
function showMap(flag) {

    if (flag == true) {
        $("#canvas-section").hide();
        $("#map-section").show();
    } else {
        $("#canvas-section").show();
        $("#map-section").hide();
    }


}
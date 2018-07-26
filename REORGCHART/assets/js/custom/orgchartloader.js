/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the supporting attributes and functions required for funtioning or orgchart.

* Inititializes parameters
* Defines supporting functions e.g. OnSelectionChange, SetModifiedStatus etc

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

var splitScreen = SPLIT_SCREEN.FALSE;
var splitScreenDirection = SPLIT_SCREEN_DIRECTION.VERTICAL;

var modified = MODIFIED_FLEG.FALSE;

var lefttreerefreshed = 1;
var duplicatedrop = 0;

var charttype = CHART_TYPE.OPERATIONAL;

var sizingscale = SIZE_FULLSCREEN.NODE_SCALE;
var nodespacing = SIZE_FULLSCREEN.NODE_SPACING;

function initOrgChart() {
    
    var userSettings =  JSON.parse($("#userSettings").val(), doNothing);
    var admimSettings = JSON.parse($("#adminSettings").val(), doNothing);
    
    charttype = $("#txtOper").val();

  loadorgchart(userSettings, admimSettings);

    $("#myDiagram").kendoDropTarget({
        filter: ".k-in",
        drop: onDropMyDiagram
    });

    $("#myDiagramPalleteContainer").hide();

    if (splitScreen == SPLIT_SCREEN.TRUE) {
        $("#myDiagramPalleteContainer").show();
        loadorgchartPallete(userSettings, admimSettings);
        
        $("#myDiagramPallete").kendoDropTarget({
            filter: ".k-in",
            drop: onDropMyDiagramPallete
        });
       
    }
    

}
// Allow the user to edit text when a single node is selected
function onSelectionChanged(e) {
    var node = e.diagram.selection.first();
    if (node instanceof go.Node) {
        updateProperties(node.data);
    } else {
        updateProperties(null);
    }
}

// Update the HTML elements for editing the properties of the currently selected node, if any
function updateProperties(data) {
    if (data === null) {
        
        // For Editing position details
        document.getElementById(FIELD.POSITIONID).value = "";
        document.getElementById(FIELD.NAME).value = "";
        document.getElementById(FIELD.TITLE).value = "";
        document.getElementById(FIELD.TITLE_PERSON).value = "";
        document.getElementById(FIELD.PERSON).value = "";
        document.getElementById(FIELD.HDN_KEY).value = "";
        document.getElementById(FIELD.HDN_PERSON_GDDBID).value =  "";


        // for editing CPM Details
        document.getElementById(FIELD.ORGUNIT_CPM).value = "";
        document.getElementById(FIELD.POSITIONID_CPM).value = "";
        document.getElementById(FIELD.TITLE_CPM).value = "";
        document.getElementById(FIELD.PERSON_CPM).value = "";

        document.getElementById(FIELD.HDN_ORGUNIT_CPM).value = "";
        document.getElementById(FIELD.HDN_POSITIONID_CPM).value = "";

        // for editing orgunit
        document.getElementById(FIELD.ORGUNIT).value = "";
        document.getElementById(FIELD.ORGUNITSHORTTEXT).value = "";
        document.getElementById(FIELD.ORGUNITTEXT).value = "";
        document.getElementById(FIELD.FDC).value = "";
        document.getElementById(FIELD.OPL3).value = "";
        
    } else {

        // For Editing position details
        document.getElementById(FIELD.POSITIONID).value = data.key || "";
        document.getElementById(FIELD.NAME).value = data.NAME || "";
        document.getElementById(FIELD.TITLE).value = data.TITLE || "";
        document.getElementById(FIELD.TITLE_PERSON).value = data.TITLE || "";

        
        var dropdownlist = $("#PERSON").data("kendoDropDownList");
        dropdownlist.dataSource.add({ ID: data.GDDBID, Title: data.NAME });
        dropdownlist.value(data.GDDBID);

        document.getElementById(FIELD.HDN_KEY).value = data.key || "";
        document.getElementById(FIELD.HDN_PERSON_GDDBID).value = data.GDDBID || "";
        

        // for editing CPM Details
        document.getElementById(FIELD.ORGUNIT_CPM).value = data.ORGUNIT || "";
        document.getElementById(FIELD.POSITIONID_CPM).value = data.POSITIONID || "";
        document.getElementById(FIELD.TITLE_CPM).value = data.TITLE || "";
        
        document.getElementById(FIELD.HDN_ORGUNIT_CPM).value = data.ORGUNIT || "";
        document.getElementById(FIELD.HDN_POSITIONID_CPM).value = data.POSITIONID || "";

        dropdownlist = $("#PERSON_CPM").data("kendoDropDownList");
        dropdownlist.dataSource.add({ ID: data.GDDBID, Title: data.NAME });
        dropdownlist.value(data.GDDBID);
        
        // for editing orgunit
        document.getElementById(FIELD.ORGUNIT).value = data.ORGUNIT || "";
        document.getElementById(FIELD.ORGUNITSHORTTEXT).value = data.ORGUNITSHORTTEXT || "";
        document.getElementById(FIELD.ORGUNITTEXT).value = data.ORGUNITTEXT || "";

        dropdownlist = $("#FDC").data("kendoDropDownList");
        dropdownlist.dataSource.add({ ID: data.FDC, Title: data.FDCTEXT });
        dropdownlist.value(data.FDC);

        dropdownlist = $("#OPL3").data("kendoDropDownList");
        dropdownlist.dataSource.add({ ID: data.OPL3, Title: data.OPL3TEXT });
        dropdownlist.value(data.OPL3);

    }
}

// This is called when the user has finished inline text-editing
function onTextEdited(e) {
    var tb = e.subject;
    if (tb === null || !tb.name) return;
    var node = tb.part;
    if (node instanceof go.Node) {
        updateProperties(node.data);
    }
}

function removePersonFromPosition(ID) {
    var example = { GDDBID: ID };
    try{
        var node = myDiagram.findNodesByExample(example).first();
        // maxSelectionCount = 1, so there can only be one Part in this collection
        var data = node.data;

        if (node instanceof go.Node && data !== null) {
            var model = myDiagram.model;
            model.setDataProperty(data, FIELD.GDDBID, "");
            model.setDataProperty(data, FIELD.NAME, "");
            node.findObject("Picture").source = CONST.IMAGE_PATH + "NOIMAGE" + CONST.IMAGE_EXTENSION;

            setModifiedStatusForNode(node.data.key, CONST.MODIFIED_STATUS_MODIFIED);
        }
    } catch (e) {
        //alert(e);
    }
}
// Update the data fields when the text is changed

function updateData(text, field) {
    var node = myDiagram.selection.first();
    // maxSelectionCount = 1, so there can only be one Part in this collection
    var data = null;
    if (node) data = node.data;
            
    if (node instanceof go.Node && data !== null) {
        var model = myDiagram.model;
        model.startTransaction("modified " + field);

        // alert(node.findObject("Picture").source);

        model.setDataProperty(data, field, text);

        /*
        if (field === FIELD.PERSON) {

            var dropdownlist = $("#"+FIELD.PERSON).data("kendoDropDownList");
            var dataItem = dropdownlist.dataItem();

            removePersonFromPosition(dropdownlist.value());
            
            model.setDataProperty(data, FIELD.GDDBID, dropdownlist.value());
            model.setDataProperty(data, FIELD.NAME, dropdownlist.text());
            node.findObject("Picture").source = CONST.IMAGE_PATH + dropdownlist.value() + CONST.IMAGE_EXTENSION;

        } else {
            model.setDataProperty(data, field, text);
        }
       */
        setModifiedStatusForNode(node.data.key, CONST.MODIFIED_STATUS_MODIFIED);
        model.commitTransaction("modified " + field);
    }
}

function setModifiedStatusForNode(key, status) {
    var node = myDiagram.model.findNodeDataForKey(key);
    if (node) {
        modified = MODIFIED_FLEG.TRUE;
        myDiagram.isModified = true;
        modified = MODIFIED_FLEG.TRUE
        if (node.CLINETMODIFIEDSTATUS == CONST.MODIFIED_STATUS_CREATED) {
            // do nothing - Node has been created in this editing session previously
        } else {
            node.CLINETMODIFIEDSTATUS = status;
        }
    }
}

// Show the diagram's model in JSON format
function load(strJson) {
   
        if (strJson) {
            document.getElementById("mySavedModel").value = strJson;
        }
        myDiagram.model = go.Model.fromJson(document.getElementById("mySavedModel").value);
        document.getElementById("myDiagram").style.zIndex = 0;

        myDiagram.isModified = false;
        modified = MODIFIED_FLEG.FALSE;

}

function loadPallete(strJson) {
    
    if (strJson) {
        document.getElementById("mySavedModelPallete").value = strJson;
    }
    myDiagramPallete.model = go.Model.fromJson(document.getElementById("mySavedModelPallete").value);
    document.getElementById("myDiagramPallete").style.zIndex = 0;
}

function unLoad(emptyModel) {
    myDiagram.model = go.Model.fromJson(document.getElementById(emptyModel).value);
}

function SaveWorkspaceData() {
    saveLocal();
    if (!UserHasAccess()) { alert(MESSAGE.SAVE_NOT_ALLOWED); return false; }
    if (confirm(MESSAGE.BEFORE_SAVE)) {
        //if (!myDiagram.isModified) { alert("No changes has been made to the diagram."); return false; }
        saveOrgChartData("#mySavedModel");
        myDiagram.isModified = false;
    } else {

    }
}

function saveLocal() {
    document.getElementById("mySavedModel").value = myDiagram.model.toJson();

    /*
    if (myDiagramPallete !== undefined)
            if(myDiagramPallete.model !== undefined && myDiagramPallete.model !== null)
                        document.getElementById("mySavedModelPallete").value = myDiagramPallete.model.toJson();
                    */
}
function loadLocal(strJson) {
    document.getElementById("mySavedModel").value = strJson;
    document.getElementById("mySavedModelPallete").value = strJson;
}

function doNothing(key, value) {
    return value;
}

function calculateShapeXDimention(shape, viewType, noOfFields) {

    switch (shape) {
        case "RoundedRectangle":
            switch (viewType) {
                case "SIMPLE":
                    return 240;
                    break;
                case "DETAILED":
                    return 240;
                    break;
                case "CUSTOM":
                    return 200;
                    break;
            }
            break;

        case "Circle":
            switch (viewType) {
                case "SIMPLE":
                    return 100;
                    break;
                case "DETAILED":
                    return 120;
                    break;
                case "CUSTOM":
                    return 150;
                    break;
            }
            break;
        case "Ellipse":
            switch (viewType) {
                case "SIMPLE":
                    return 140;
                    break;
                case "DETAILED":
                    return 160;
                    break;
                case "CUSTOM":
                    return 150;
                    break;
            }
            break;
    }
    return 160;
}
function calculateShapeYDimention(shape, viewType, noOfFields) {
    switch (shape) {
        case "RoundedRectangle":
            switch (viewType) {
                case "SIMPLE":
                    return 70;
                    break;
                case "DETAILED":
                    return 110;
                    break;
                case "CUSTOM":
                    return 200;
                    break;
            }
            break;

        case "Circle":
            switch (viewType) {
                case "SIMPLE":
                    return 100;
                    break;
                case "DETAILED":
                    return 120;
                    break;
                case "CUSTOM":
                    return 100;
                    break;
            }
            break;
        case "Ellipse":
            switch (viewType) {
                case "SIMPLE":
                    return 60;
                    break;
                case "DETAILED":
                    return 120;
                    break;
                case "CUSTOM":
                    return 100;
                    break;
            }
            break;
    }
    return 80;
}

function showMessage(s) {
    document.getElementById("diagramEventsMsg").textContent = s;
}

function reload() {
   
    unloadPopupBox('settingsBox');
    saveLocal();

    //myDiagram.clear();
    var radArray = document.getElementsByName('splitscreen');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            splitScreen = radArray[i].value;
        }
    }

    $("#myDiagram").detach();
    $("#myDiagramPallete").detach();

    $("#myDiagramPalleteContainer").hide();

    var iDiv = document.createElement('div');
    iDiv.id = 'myDiagram';
    iDiv.className = "workspace";
    document.getElementById('myDiagramContainer').appendChild(iDiv);

    if (splitScreen == SPLIT_SCREEN.TRUE) {
        iDiv = document.createElement('div');
        iDiv.id = 'myDiagramPallete';
        iDiv.className = "pallete";
        document.getElementById('myDiagramPalleteContainer').appendChild(iDiv);
        $("#myDiagramPalleteContainer").show();
    }

    initOrgChart();

}

function reloadRemote(remoteData) {

    //unloadPopupBox('settingsBox');
    loadLocal(remoteData);

    //myDiagram.clear();
    var radArray = document.getElementsByName('splitscreen');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            splitScreen = radArray[i].value;
        }
    }

    $("#myDiagram").detach();
    $("#myDiagramPallete").detach();

    $("#myDiagramPalleteContainer").hide();

    var iDiv = document.createElement('div');

    iDiv.id = 'myDiagram';
    iDiv.className = "workspace";
    document.getElementById('myDiagramContainer').appendChild(iDiv);

    if (splitScreen == SPLIT_SCREEN.TRUE) {
        iDiv = document.createElement('div');
        iDiv.id = 'myDiagramPallete';
        iDiv.className = "pallete";
        document.getElementById('myDiagramPalleteContainer').appendChild(iDiv);
        $("#myDiagramPalleteContainer").show();
    }

    initOrgChart();

}

function onDropMyDiagram(e) {
    var node = tv.dataItem(e.draggable.currentTarget);
    if(!node)
        node = tvtemp.dataItem(e.draggable.currentTarget);
    $("#txtShowLevel").val(node.ID);
    lefttreerefreshed = 1;
    loadOrgChartData();
}

function onDropMyDiagramPallete(e) {
    var node = tv.dataItem(e.draggable.currentTarget);
    if (!node)
        node = tvtemp.dataItem(e.draggable.currentTarget);
    $("#txtShowLevel").val(node.ID);
    lefttreerefreshed = 0;
    loadOrgChartData();
}

function getSelectedText(elementId) {
    var elt = document.getElementById(elementId);

    if (elt.selectedIndex == -1)
        return null;

    return elt.options[elt.selectedIndex].text;
}

// if the shape is editable
function isTrue(key) {
    if (key > 0)
        return true;
    else
        return false;
};

function makePattern(fillcolor, linecolor) {
    var patternCanvas = document.createElement('canvas');
    patternCanvas.width = 5;
    patternCanvas.height = 5;

    var pctx = patternCanvas.getContext('2d');

    // This creates small squares in background, 
    // which will create a look of grpah paper look
    pctx.beginPath();

    pctx.moveTo(0.0, 0.0);
    pctx.lineTo(5, 0.0);
    pctx.lineTo(5, 5);
    pctx.lineTo(0.0, 5);
    pctx.lineTo(0.0, 0.0);

    pctx.closePath();
    pctx.fillStyle = fillcolor;
    pctx.fill();

    pctx.lineWidth = 0.3;
    pctx.strokeStyle = linecolor;
    pctx.lineJoin = "miter";
    pctx.miterLimit = 4.0;

    pctx.stroke();

    return patternCanvas;
}

function makeGraph() {
    var chart = "4:10:40:100:80:50:60:80:20:30:60:80:50:60#4:10:40:100:80:50:60:80:20:30:60:80:50:60"

    var chartData1 = chart.split("#")[0];
    var chartData2 = chart.split("#")[1];

    var dataArry = chartData1.split(":");

    var width = 340;
    var height = 110;
    var startPosition_X = 50;
    var startPosition_Y = 110;

    //scale
    var maxHeight = 200;

    
    // range of incoming value
    var maxVal = 100;
    var minVal = 0;

    if (maxVal > 0.0099)
        maxHeight = 0.01
    if (maxVal > 0.099)
        maxHeight = 0.1
    if (maxVal > 0.99)
        maxHeight = 1
    if (maxVal > 9.9)
        maxHeight = 50
    if (maxVal > 49)
        maxHeight = 100
    if (maxVal > 99)
        maxHeight = 200


    
    var gap = (width - 2 * startPosition_X) / dataArry[0];

    var patternCanvas = document.createElement('canvas');
    patternCanvas.width = width;
    patternCanvas.height = height;

    var pctx = patternCanvas.getContext('2d');

    var fillcolor = "orange";
    var linecolor = "red";

    // This creates small squares in background, 
    // which will create a look of grpah paper look
    pctx.beginPath();

    var itemheight;
    var nextitemheight;

    for (i = 0; i < (dataArry[0]+1); i++) {
        
        itemheight = (((maxVal - minVal) * dataArry[i + 1]) / maxHeight) + minVal; 

        pctx.moveTo(startPosition_X + gap * i, startPosition_Y);
        pctx.lineTo(startPosition_X + gap * i + 10, startPosition_Y);
        pctx.lineTo(startPosition_X + gap * i + 10, startPosition_Y - itemheight);
        pctx.lineTo(startPosition_X + gap * i, startPosition_Y - itemheight);
        pctx.lineTo(startPosition_X + gap * i, startPosition_Y);
    }

    var dataArry = chartData2.split(":");
    for (i = 0; i < (dataArry[0]); i++) {

       itemheight = (((maxVal - minVal) * dataArry[i + 1]) / maxHeight) + minVal;
       nextitemheight  = (((maxVal - minVal) * dataArry[i + 2]) / maxHeight) + minVal;

       pctx.moveTo(startPosition_X + gap * i , startPosition_Y - itemheight);
       pctx.lineTo(startPosition_X + gap * (i + 1), startPosition_Y - nextitemheight);
    }
    pctx.closePath();
    pctx.fillStyle = fillcolor;
    pctx.fill();

    pctx.lineWidth = 1;
    pctx.strokeStyle = linecolor;
    pctx.lineJoin = "miter";
    pctx.miterLimit = 4.0;

    pctx.stroke();

    return patternCanvas;
}
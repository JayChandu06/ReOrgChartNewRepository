/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the method which initializes and load the main orgchart screen. This methos is invoked from init.js file 
when application is loaded

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

function loadorgchart(userSettings, admimSettings) {
    //  if (window.goSamples) goSamples();  // init for these samples -- you don't need to call this
    var $ = go.GraphObject.make;  // for conciseness in defining templates

    var skin = "1";
    var shape = "RoundRectangle";
    var view = "SIMPLE";
    var radArray = document.getElementsByName('shape');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            shape = radArray[i].value;
        }
    }

    view = getSelectedText('view');
    // view = document.getElementsName('view').value;

    var shape_X_Dimention = calculateShapeXDimention(shape, view, userSettings.displayFields.length);
    var shape_Y_Dimention = calculateShapeYDimention(shape, view, userSettings.displayFields.length);

    var pictureWidth = userSettings.picture[0].width;
    var pictureHeight = userSettings.picture[0].height;

    var pictureFlag = "Y";
    radArray = document.getElementsByName('picture');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            pictureFlag = radArray[i].value;
        }
    }

    radArray = document.getElementsByName('skin');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            skin = radArray[i].value;
        }
    }

    if (pictureFlag == "N") {
        pictureWidth = 0;
        pictureHeight = 0;
    }

    radArray = document.getElementsByName('splitscreen');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            splitScreen = radArray[i].value;
        }
    }

    radArray = document.getElementsByName('splitscreendirection');
    for (var i = 0; i < radArray.length; i++) {
        if (radArray[i].checked) {
            splitScreenDirection = radArray[i].value;
        }
    }

    var h = window.innerHeight || document.documentElement.clientHeight || document.body.clientHeight;
    var w = window.innerWidth || document.documentElement.clientWidth || document.body.clientWidth;

    var treeAlignment = go.TreeLayout.ArrangementVertical;
    var parentAlignment = go.TreeLayout.AlignmentCenterChildren;
    var childAlignment = go.TreeLayout.AlignmentBus;

    var elm = document.getElementById("myDiagramContainer");

    if (splitScreen == SPLIT_SCREEN.FALSE) {

        elm.style.height = (h - elm.offsetTop - 40) + "px";
        elm.style.width = w + "px";

        elm.style.styleFloat = "left";
        elm.style.borderRight = "0px";
        elm.style.borderBottom = "0px";

        document.getElementById('myDiagramPalleteContainer').style.display = "hidden";

    } else {
        if (splitScreenDirection == SPLIT_SCREEN_DIRECTION.VERTICAL) {

            elm.style.height = (h - elm.offsetTop - 30) + "px";
            elm.style.width = ((w / 2) - 10) + "px";
            //elm.style.width = (w / 2)  + "px";

            elm.style.styleFloat = "left";
            elm.style.borderRight = "2px solid black";
            elm.style.borderBottom = "0px";

            treeAlignment = go.TreeLayout.ArrangementVertical;
            parentAlignment = go.TreeLayout.AlignmentBus;
            childAlignment = go.TreeLayout.AlignmentBottomRightBus;

        } else {

            elm.style.height = (h - elm.offsetTop - 2) / 2 + "px";
            elm.style.width = w + "px";

            elm.style.borderRight = "0px";
            elm.style.borderBottom = "2px solid black";

            treeAlignment = go.TreeLayout.ArrangementVertical;
            parentAlignment = go.TreeLayout.AlignmentCenterChildren;
            childAlignment = go.TreeLayout.AlignmentBus;
        }
    }

    myDiagram =
      $(go.Diagram, "myDiagram", // must be the ID or reference to div
        {
            initialContentAlignment: go.Spot.Center,
            //scale: 1,
            //initialAutoScale: go.Diagram.Uniform,
            // make sure users can only create trees
            validCycle: go.Diagram.CycleDestinationTree,
            // users can select only one part at a time
            maxSelectionCount: 1,
            layout:
              $(go.TreeLayout,
                {
                    treeStyle: go.TreeLayout.StyleRootOnly,
                    arrangement: treeAlignment,
                    // properties for most of the tree:
                    angle: 90,
                    layerSpacing: 20,
                    nodeSpacing: 20,
                    breadthLimit: w,
                    alignment: parentAlignment,
                    // properties for the "last parents":
                    alternateAngle: 90,
                    alternateLayerSpacing: 20,
                    alternateAlignment: childAlignment,
                    alternateNodeSpacing: 20
                }),
            // support editing the properties of the selected person in HTML
           // "ChangedSelection": onSelectionChanged,
            "TextEdited": onTextEdited,
            // enable undo & redo
            "undoManager.isEnabled": true,
            //JAY
           // allowDragOut: false,
            allowDragOut: true,
            allowDrop: true
        });
    //myDiagram.scale = .8;
    // when the document is modified, add a "*" to the title and enable the "Save" button
    myDiagram.addDiagramListener("Modified", function (e) {
        var button = document.getElementById("SaveButton");
        if (button) button.disabled = !myDiagram.isModified;
        var idx = document.title.indexOf("*");
        if (myDiagram.isModified) {
            if (idx < 0) document.title += "*";
        } else {
            if (idx >= 0) document.title = document.title.substr(0, idx);
        }
    });

    // Deleting multiple nodes is not allowed
    myDiagram.addDiagramListener("SelectionDeleting",
    function (e) {
        if (e.diagram.selection.count > 1) {
            e.cancel = true;
            showMessage("Cannot delete multiple selected parts");
        }
    });

    // pre action listners

    myDiagram.addDiagramListener("SelectionDeleting",
   function (e) {  // before deleting
       e.subject.each(function (part) { if (!(part instanceof go.Shape)) { if (!isTrue(part.data.EDITABLE)) { e.cancel = true; alert(MESSAGE.DELETION_NOT_ALLOWED); } else { var link = part.findTreeParentLink(); link.fromNode = null; part.data.parent = KEY.ORPHAN; setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_DELETED); e.cancel = true; } } });
   });

    // post action listers
    myDiagram.addDiagramListener("SelectionDeleted",
    function (e) {  // for delete and cut 
        e.subject.each(function (part) { if (!(part instanceof go.Shape)) { part.data.parent = KEY.ORPHAN; setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_DELETED); } });
    });

    myDiagram.addDiagramListener("ClipboardChanged",
    function (e) {
        e.subject.each(function (part) { if (!(part instanceof go.Shape)) { /* setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_MODIFIED);*/ } });
    });

    myDiagram.addDiagramListener("ClipboardPasted",
    function (e) {
        e.subject.each(function (part) { if (!(part instanceof go.Shape)) { setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_CREATED); } });
    });

    /*
    myDiagram.addDiagramListener("TextEdited",
    function (e) {
        e.subject.each(function (part) { if (!(part instanceof go.Shape)) { setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_MODIFIED); } });
    });
    */
    myDiagram.addDiagramListener("LinkRelinked",
    function (e) {
        e.subject.each(function (part) { if (!(part instanceof go.Shape)) { setModifiedStatusForNode(part.data.key, CONST.MODIFIED_STATUS_RELINKED); } });
    });

    myDiagram.addDiagramListener("ExternalObjectsDropped", function (e) {
        var newnode = myDiagramPallete.selection.first(); //always dropping from palette
        if (isDuplicateKey(newnode.data.key)) { //already in diagram, ignore drop
            //alert('duplicate');
            //e.cancel = true;
            //myDiagram.commandHandler.deleteSelection();
        }
    });

    // Read the SKIN Color array for selected skin
    var levelColors = SKIN_COLOR[skin];

    // override TreeLayout.commitNodes to also modify the background brush based on the tree depth level
    myDiagram.layout.commitNodes = function () {
        go.TreeLayout.prototype.commitNodes.call(myDiagram.layout);  // do the standard behavior


        // then go through all of the vertexes and set their corresponding node's Shape.fill
        // to a brush dependent on the TreeVertex.level value
        myDiagram.layout.network.vertexes.each(function (v) {
            if (v.node) {
                var level = v.level % (levelColors.length);
                var colors = levelColors[level].split("/");
                var shape = v.node.findObject("SHAPE");
                if (shape) shape.fill = $(go.Brush, "Linear", { 0: colors[0], 1: colors[1], start: go.Spot.Left, end: go.Spot.Right });
            }
        });
    }

    function onMouseDragEnter(e, node, prev) {
        var diagram = node.diagram;
        var selnode = diagram.selection.first();
        if (!mayWorkFor(selnode, node)) return;
        var shape = node.findObject("SHAPE");
        if (shape) {
            shape._prevFill = shape.fill;  // remember the original brush
            shape.fill = "darkred";
        }
    };
    function onMouseDragLeave(e, node, next) {
        var shape = node.findObject("SHAPE");
        if (shape && shape._prevFill) {
            shape.fill = shape._prevFill;  // restore the original brush
        }
    };

    function onMouseDrop(e, node) {
        //alert("Hello");
        var diagram = node.diagram;
        var selnode = diagram.selection.first();  // assume just one Node in selection

        // if (isDuplicateDrop()) { alert("duplicate"); e.cancel = true; return false; }
        //if (isDuplicateNode(selnode.part.data)) { alert("duplicate"); e.cancel = true; }

        //if (isTrue(selnode.part.data.EDITABLE)) {
            if (mayWorkFor(selnode, node)) {
                //if (charttype == CHART_TYPE.LEGAL) {
                //    if (!isOrgunit(node.part.data.key)) {
                //        // not allowed to drop on an position
                //        alert(MESSAGE.REPORTING_NOT_ALLOWED_TO_POSITION);
                //        e.cancel = true;
                //        return true;
                //    } else {
                //        // Copy the new parent orgunit properties to new child node 
                //        // inheritParentsProperty(key,parent)
                //        inheritParentsProperty(selnode.part.data.key, node.part.data.key);
                //    }
                //} else {

                //    if (!isCPMOPMRuleFollowed(selnode, node)) {
                //        if (confirm(MESSAGE.CPM_OPM_RULE_VIOLATED)) {
                //            // set that node has reporting issues
                //            selnode.part.data.REPORTINGERROR = "1";
                //        }
                //        else {
                //            e.cancel = true;
                //            return true;
                //        }
                //    } else {
                //        selnode.part.data.REPORTINGERROR = "0";
                //    }
                //}

                // find any existing link into the selected node
                var link = selnode.findTreeParentLink();
                //alert(link);
                if (link !== null) {  // reconnect any existing link
                    link.fromNode = node;
                } else {  // else create a new link
                    diagram.toolManager.linkingTool.insertLink(node, node.port, selnode, selnode.port);
                }

                setModifiedStatusForNode(selnode.part.data.key, CONST.MODIFIED_STATUS_RELINKED);
            } else {
                alert(MESSAGE.RECURSIVE_LOOP);
                e.cancel = true;
            }
        //} else {
        //    alert(MESSAGE.NONEDITABLE);
        //    e.cancel = true;
        //}
    };

    // when a node is clicked,grill down its child
    function drillDown(e, obj) {
        var clicked = obj.part;
        if (clicked !== null) {
            var thisemp = clicked.data;

            if (isFirstElement(thisemp.key))
                document.getElementById("txtShowLevel").value = thisemp.parent;
            else
                document.getElementById("txtShowLevel").value = thisemp.key;
            lefttreerefreshed = 1;
            loadOrgChartData();
        }
    }

    // Adding a new node to the diagram
    function addOrgunit(e, obj) {
        var clicked = obj.part;
        if (clicked !== null) {
            var thisemp = clicked.data;
            myDiagram.startTransaction("add orgunit");
            // var nextkey = "TEMP_" + (myDiagram.model.nodeDataArray.length + 1).toString();
            var nextkey = "ORGUNIT" + Math.floor((Math.random() * 100000) + 1).toString();
            var newemp = { key: nextkey, NAME: thisemp.NAME, parent: thisemp.key, ORGUNIT: nextkey, ORGUNITTEXT: "(new orgunit)", ORGUNITSHORTTEXT: "(new orgunit)", TITLE: thisemp.TITLE, POSITIONID: thisemp.POSITIONID, GDDBID: thisemp.GDDBID, FDC: thisemp.FDC, FDCTEXT: thisemp.FDCTEXT, OPL3: thisemp.OPL3, OPL3TEXT: thisemp.OPL3TEXT, LEGALENTITY: thisemp.LEGALENTITY, CPM: thisemp.CPM, CPMLEGALENTITY: thisemp.CPMLEGALENTITY, CLINETMODIFIEDSTATUS: "2", VERSIONSTATUS: "2", EDITABLE: "1", };
            myDiagram.model.addNodeData(newemp);
            //setModifiedStatusForNode(nextkey, CONST.MODIFIED_STATUS_CREATED);
            myDiagram.commitTransaction("add orgunit");
        }
    }
    function addPositionOperational(e, obj) {
        var clicked = obj.part;
        if (clicked !== null) {
            var thisemp = clicked.data;
            myDiagram.startTransaction("add employee");
            // var nextkey = "TEMP_" + (myDiagram.model.nodeDataArray.length + 1).toString();
            var nextkey = "POSITION" + Math.floor((Math.random() * 100000) + 1).toString();
            var newemp = { key: nextkey, NAME: "(new position)", parent: thisemp.key, TITLE: "(title)", POSITIONID: nextkey, GDDBID: "GDDBID", ORGUNIT: thisemp.ORGUNIT, ORGUNITTEXT: thisemp.ORGUNITTEXT, LEGALENTITY: "", CPM: "", CPMLEGALENTITY: "", CLINETMODIFIEDSTATUS: "2", VERSIONSTATUS: "2", EDITABLE: "1", };
            myDiagram.model.addNodeData(newemp);
            //setModifiedStatusForNode(nextkey, CONST.MODIFIED_STATUS_CREATED);
            myDiagram.commitTransaction("add employee");
        }
    }
    function addPositionLegal(e, obj) {
        var clicked = obj.part;
        if (clicked !== null) {
            var thisemp = clicked.data;
            myDiagram.startTransaction("add employee");
            // var nextkey = "TEMP_" + (myDiagram.model.nodeDataArray.length + 1).toString();
            var nextkey = "POSITION" + Math.floor((Math.random() * 100000) + 1).toString();
            var newemp = { key: nextkey, NAME: "(new position)", parent: thisemp.key, TITLE: "(title)", POSITIONID: nextkey, GDDBID: "", ORGUNIT: thisemp.ORGUNIT, ORGUNITTEXT: thisemp.ORGUNITTEXT, ORGUNITSHORTTEXT: thisemp.ORGUNITSHORTTEXT, FDC: thisemp.FDC, FDCTEXT: thisemp.FDCTEXT, OPL3: thisemp.OPL3, OPL3TEXT: thisemp.OPL3TEXT, LEGALENTITY: "", CPM: "", CPMLEGALENTITY: "", CLINETMODIFIEDSTATUS: "2", VERSIONSTATUS: "2", EDITABLE: "1", };
            myDiagram.model.addNodeData(newemp);
            //setModifiedStatusForNode(nextkey, CONST.MODIFIED_STATUS_CREATED);
            myDiagram.commitTransaction("add employee");
        }
    }

    // this is used to determine feedback during drags
    function mayWorkFor(node1, node2) {
        if (!(node1 instanceof go.Node)) return false;  // must be a Node
        if (node1 === node2) return false;  // cannot work for yourself
        if (node2.isInTreeOf(node1)) return false;  // cannot work for someone who works for you
        //if (node1.part.data.CPM != node2.part.data.GDDBID && node1.part.data.CPMLEGALENTITY == node2.part.data.LEGALENTITY) return false;
        return true;
    }

    function isCPMOPMRuleFollowed(node1, node2) {
        if (node1.part.data.CPM != node2.part.data.GDDBID && node1.part.data.CPMLEGALENTITY == node2.part.data.LEGALENTITY) return false;
        return true;
    }

    function isPosition(node) {
        return isOrgunit(node.part.data.key);
    }
    // Find if the drop from palette is duplicate
    function isDuplicateDrop() {
        // alert("trying to find dupe key " + key);

        if (!myDiagramPallete) return false;
        if (!myDiagramPallete.selection) return false;

        var newnode = myDiagramPallete.selection.first();
        if (!newnode) return false;

        var key = newnode.data.key;
        var nodes = myDiagram.nodes;
        while (nodes.next()) {
            var n = nodes.value;
            var nd = n.data;
            var id = nd.key;
            if (id === key)
                return true;
        }
        return false;
    }
    // Find duplicate
    function isDuplicateKey(key) {
        // alert("trying to find dupe key " + key);
        var nodes = myDiagram.nodes;
        while (nodes.next()) {
            var n = nodes.value;
            var nd = n.data;
            var id = nd.key;
            if (id === key)
                return true;
        }
        return false;
    }

    function isDuplicateNode(data) {
        // alert("trying to find dupe key " + data.id);
        var contains = myDiagram.model.containsNodeData(data);
        return contains;
    }
    // This converter is used by the Picture.
    function findHeadShot(key) {
        //  if (key > 16) return ""; // There are only 16 images on the server
        //alert(key);
        return CONST.IMAGE_PATH + key + CONST.IMAGE_EXTENSION;
    };


    // Id first element but NOT CEO
    function isFirstElement(key) {
        // alert(key);
        if (key == document.getElementById("txtShowLevel").value && myDiagram.model.findNodeDataForKey(key).parent != -1)
            return true;
        else
            return false;
    };

    // if the shape is editable
    function isTrue(key) {
        if (key > 0)
            return true;
        else
            return false;
    };

    function isOrgunit(key) {
        try {
            if (key.charAt(0) == '1' || key.charAt(0) == 'O') return true;
            else return false;
        } catch (e) {
            return false;
        }
    };


    function isEditableOrgunit(key) {
        try {
            var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
            if (editable > 0 && (key.charAt(0) == '1' || key.charAt(0) == 'O')) return true;
            else return false;
        } catch (e) {

            return false;
        }
    };




    function findPattern(key) {
        var versionstatus = myDiagram.model.findNodeDataForKey(key).VERSIONSTATUS;
        var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
        if (versionstatus == 2) // new node
        {
            if (key.charAt(0) == '1' || key.charAt(0) == 'O') // Orgunit
            {
                var pattern = $(go.Brush, "Pattern", { pattern: makePattern("#fed300", findBorderColorForPopulation(editable)) });
                return pattern;
            }
            else {
                return "#fed300"
            }
        }
        else // existing node
        {

            if (key.charAt(0) == '1' || key.charAt(0) == 'O') // Orgunit
            {
                var pattern = $(go.Brush, "Pattern", { pattern: makePattern("#FFFFFF", findBorderColorForPopulation(editable)) });

                return pattern;
            }
            else {
                return "#FFFFFF"
            }
        }
    }
    function isEditablePositionOperational(key) {
        try {
            var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
            if (charttype == CHART_TYPE.OPERATIONAL && editable > 0 && (key.charAt(0) == '2' || key.charAt(0) == 'P')) return true;
            else return false;
        } catch (e) {

            return false;
        }
    };
    function isEditablePositionLegal(key) {
        try {
            var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
            if (charttype == CHART_TYPE.LEGAL && editable > 0 && (key.charAt(0) == '2' || key.charAt(0) == 'P')) return true;
            else return false;
        } catch (e) {

            return false;
        }
    };
    function isEditableVersion(key) {
        var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
       // if (document.getElementById("USER").value == LoggedInUser && document.getElementById("VERSION").value > 0) return true;
        if (editable > 0) return true;
        else return false;
    };

    function isOrgunitInEditableVersion(key) {
        var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
        if (editable > 0 && (key.charAt(0) == '1' || key.charAt(0) == 'O')) return true;
        else return false;
    };
    function isPositionInEditableVersion(key) {
        var editable = myDiagram.model.findNodeDataForKey(key).EDITABLE;
        if (editable > 0 && (key.charAt(0) == '2' || key.charAt(0) == 'P')) return true;
        else return false;
    };
    

    function isReportingError(key) {
        if (key == 1)
            return true;
        else
            return false;
    }
    // if not first element and has children then return true.
    function hasChildren(key) {

        // if first element 
        if (key == document.getElementById("txtShowLevel").value) return false;

        //if has children
        return myDiagram.model.findNodeDataForKey(key).SOC_COUNT > 0 ? true : false;
    };

    // This converts dotted line to a panel border width
    function findStrokeWidthForDottedLine(key) {
        if (key == "Y")
            return 4;
        else
            return 0;
    };

    function findStrokeWidth(key) {

        if (key > -1)
            return 6;
        else
            return 2;
    }

    function findBorderColorForPopulation(key) {
        if (key > -1)
            return "#e44c16";
        else
            return "#00a4a4";

    }

    function findNewItemColor(key) {
        //if (key == -1) return "black";
        if (key == 2)
            return "#fed300";
        else
            return "white";
    }

    function isDeleted(key) {
        if (key == -1)
            return true;
        else
            return false;
    }

    function findDottedLineReport(key) {
        if (key == 'N')
            return null;
        else
            return [2, 6];

    }

    function drawGraph(key) {
        //return makeGraph();
        var pattern = $(go.Brush, "Pattern", { pattern: makeGraph() });
       // var pattern = $(go.Brush, "Pattern", { pattern: makePattern("#008899","#994433") });
        
        return pattern;
    }

    //Node Context Menu
    var nodeMenu =  // context menu for each Node
     $(go.Adornment, "Vertical",
        { alignment: go.Spot.Left },
        $("ContextMenuButton",
          $(go.TextBlock, " Add New Orgunit"),
          new go.Binding("visible", FIELD.KEY, isOrgunitInEditableVersion),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { addOrgunit(e, obj); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Add New Position"),
          new go.Binding("visible", FIELD.KEY, isPositionInEditableVersion),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { addPositionOperational(e, obj); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Add New Position"),
          new go.Binding("visible", FIELD.KEY, isOrgunitInEditableVersion),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { addPositionLegal(e, obj); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Edit Position Details"),
          new go.Binding("visible", FIELD.KEY, isEditablePositionOperational),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { onSelectionChanged(e); loadPopupBox('propertiesPanel'); } }
          ),
          $("ContextMenuButton",
          $(go.TextBlock, " Edit Position Details"),
          new go.Binding("visible", FIELD.KEY, isEditablePositionLegal),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { onSelectionChanged(e); loadPopupBox('propertiesPanel_position'); } }
          ),
          $("ContextMenuButton",
          $(go.TextBlock, " Edit OrgUnit Details"),
          new go.Binding("visible", FIELD.KEY, isEditableOrgunit),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { onSelectionChanged(e); loadPopupBox('propertiesPanel_orgunit_details'); } }
          ),
          $("ContextMenuButton",
          $(go.TextBlock, " Assign/Change CPM"),
          new go.Binding("visible", FIELD.KEY, isEditableOrgunit),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { onSelectionChanged(e); loadPopupBox('propertiesPanel_orgunit_cpm'); } }
          ),
         $("ContextMenuButton",
          $(go.TextBlock, " Cut Selected (CNTRL X)"),
          new go.Binding("visible", FIELD.EDITABLE, isTrue),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { e.diagram.commandHandler.cutSelection(); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Paste Selected (CNTRL V) "),
          new go.Binding("visible", FIELD.EDITABLE, isTrue),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { e.diagram.commandHandler.pasteFromClipboard(); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Delete Selected (Delete) "),
          new go.Binding("visible", FIELD.EDITABLE, isTrue),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { e.diagram.commandHandler.deleteSelection(); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Undo (CNTRL Z) "),
          new go.Binding("visible", FIELD.EDITABLE, isTrue),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { e.diagram.commandHandler.undo(); } }
         ),
         $("ContextMenuButton",
          $(go.TextBlock, " Redo (CNTRL Y) "),
          new go.Binding("visible", FIELD.EDITABLE, isTrue),
          { width: 200, height: 25, alignment: go.Spot.Left },
          { click: function (e, obj) { e.diagram.commandHandler.redo(); } }
         )
       );

    // define simple the Node template
    // var simpleNodeTemplate =
    NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.SIMPLE] =
    $(go.Node, "Auto",
        {
            contextMenu: nodeMenu
        },
        {
            doubleClick: drillDown
        },
        {
            // handle dragging a Node onto a Node to (maybe) change the reporting relationship
            mouseDragEnter: onMouseDragEnter,
            mouseDragLeave: onMouseDragLeave,
            mouseDrop: onMouseDrop
        },
        // for sorting, have the Node.text be the data.name
        new go.Binding("text", FIELD.NAME),
        // bind the Part.layerName to control the Node's layer depending on whether it isSelected
        new go.Binding("layerName", "isSelected", function (sel) { return sel ? "Foreground" : ""; }).ofObject(),

        // define the node's outer shape
        $(go.Shape, shape,

         new go.Binding("strokeWidth", FIELD.DOTTED_LINE_FLAG, findStrokeWidthForDottedLine),
          {
              name: "SHAPE", fill: "white", stroke: "white",
              // set the port properties:
              portId: "", fromLinkable: true, toLinkable: true, cursor: "pointer"
          }),

       $(go.Panel, "Vertical",
         $(go.Picture,
            {
                name: 'UpArrow',
                desiredSize: new go.Size(15, 6),
                source: "images/arrow_white_up.png",
                visible: false
            },
            new go.Binding("visible", FIELD.KEY, isFirstElement)),

        $(go.Panel, "Horizontal",
        $(go.Picture,
            {
                name: 'Person',
                desiredSize: new go.Size(pictureWidth, pictureHeight),
                margin: new go.Margin(4, 4, 4, 4),
                visible: true
            },
            new go.Binding("source", FIELD.UNIQUEID, findHeadShot)),

        $(go.Panel, "Table",
            {
                desiredSize: new go.Size(shape_X_Dimention, shape_Y_Dimention),
                margin: new go.Margin(4, 4, 4, 4),
                defaultAlignment: go.Spot.Center
            },

          // define the panel where the text will appear

            $(go.TextBlock, textStyleTitle(skin),  // the name
              {
                  row: 1, column: 1, columnSpan: 4,
                  editable: false, isMultiline: true,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.NAME).makeTwoWay()),

           $(go.TextBlock, textStyle(skin),
               {
                   row: 2, column: 1, columnSpan: 4,
                   editable: false, isMultiline: false,
                   minSize: new go.Size(10, 16)
               },
              new go.Binding("text", FIELD.DIVISION).makeTwoWay()),

            $(go.TextBlock, textStyle(skin),
              {
                  row: 3, column: 1, columnSpan: 4,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.TITLE).makeTwoWay())

          )  // end Table Panel

        ), // end Horizontal Panel

        $(go.Picture,
            {
                name: 'DownArrow',

                desiredSize: new go.Size(15, 6),
                source: "images/arrow_white_down.png",
                alignment: go.Spot.Bottom,
                visible: false

            },
            new go.Binding("visible", FIELD.KEY, hasChildren))
       )

      );  // end Node


    // var detailedNodeTemplate =
    NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.DETAILED] =
    $(go.Node, "Auto",
        {
            contextMenu: nodeMenu
        },
        {
            doubleClick: drillDown
        },
        {
            // handle dragging a Node onto a Node to (maybe) change the reporting relationship
            mouseDragEnter: onMouseDragEnter,
            mouseDragLeave: onMouseDragLeave,
            mouseDrop: onMouseDrop
        },
        // for sorting, have the Node.text be the data.name
        new go.Binding("text", FIELD.NAME),
        // bind the Part.layerName to control the Node's layer depending on whether it isSelected
        new go.Binding("layerName", "isSelected", function (sel) { return sel ? "Foreground" : ""; }).ofObject(),

        // define the node's outer shape
        $(go.Shape, shape,
         new go.Binding("strokeWidth", FIELD.DOTTED_LINE_FLAG, findStrokeWidthForDottedLine),
          {
              name: "SHAPE", fill: "white", stroke: "white",
              // set the port properties:
              portId: "", fromLinkable: true, toLinkable: true, cursor: "pointer"
          }),

      $(go.Panel, "Vertical",
         $(go.Picture,
            {
                name: 'UpArrow',
                desiredSize: new go.Size(15, 6),
                position: new go.Point(100, 0),
                source: "images/arrow_white_up.png",

                visible: false

            },
        new go.Binding("visible", FIELD.KEY, isFirstElement)),
        $(go.Panel, "Horizontal",
          $(go.Picture,
            {
                name: 'Picture',
                alignment: go.Spot.Top,
                desiredSize: new go.Size(pictureWidth, pictureHeight),
                margin: new go.Margin(4, 4, 4, 4),

            },
            new go.Binding("source", FIELD.UNIQUEID, findHeadShot)),

          // define the panel where the text will appear
          $(go.Panel, "Table",
            {
                desiredSize: new go.Size(shape_X_Dimention, shape_Y_Dimention),
                margin: new go.Margin(4, 4, 4, 4),
                defaultAlignment: go.Spot.Left
            },

            $(go.RowColumnDefinition,
                    { row: 4, separatorStrokeWidth: 1, separatorStroke: "orange", separatorPadding: 6 }),

            $(go.TextBlock, textStyle(skin),  // the name
              {
                  row: 0, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.ORGUNITTEXT).makeTwoWay()),

           $(go.TextBlock, textStyle(skin),
               {
                   row: 1, column: 0, columnSpan: 7,
                   editable: false, isMultiline: false,
                   minSize: new go.Size(10, 16)
               },
              new go.Binding("text", FIELD.LEGALENTITY).makeTwoWay()),

            $(go.TextBlock, textStyle(skin),
              {
                  row: 2, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.TITLE).makeTwoWay()),

           $(go.TextBlock, textStyleTitle(skin),
              {
                  row: 3, column: 0, columnSpan: 6,
                  editable: false, isMultiline: true,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.NAME).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 2, columnSpan: 3,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.DIVISION).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 5, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right

              },
             "      " + FIELD.NOR_COUNT + " : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 6, columnSpan: 1,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.NOR_COUNT).makeTwoWay()),
       $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 2, columnSpan: 3,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.SITECODE).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 5, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right
              },
             "      " + FIELD.SOC_COUNT + " : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 6, columnSpan: 1,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.SOC_COUNT).makeTwoWay())


          )  // end Table Panel
        ), // end Horizontal Panel
          $(go.Picture,
            {
                name: 'DownArrow',

                desiredSize: new go.Size(15, 6),
                source: "images/arrow_white_down.png",
                alignment: go.Spot.Bottom,
                visible: false

            },
            new go.Binding("visible", FIELD.KEY, hasChildren))
       )
      );  // end Node;


    //var customNodeTemplate =
    NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.CUSTOM] =
        $(go.Node, "Auto",
        {
            contextMenu: nodeMenu
        },
        {
            doubleClick: drillDown
        },
        {
            // handle dragging a Node onto a Node to (maybe) change the reporting relationship
            mouseDragEnter: onMouseDragEnter,
            mouseDragLeave: onMouseDragLeave,
            mouseDrop: onMouseDrop
        },
        // for sorting, have the Node.text be the data.name
        new go.Binding("text", FIELD.NAME),
        // bind the Part.layerName to control the Node's layer depending on whether it isSelected
        new go.Binding("layerName", "isSelected", function (sel) { return sel ? "Foreground" : ""; }).ofObject(),

        // define the node's outer shape
        $(go.Shape, shape,
        new go.Binding("stroke", FIELD.EDITABLE, findBorderColorForPopulation),
        new go.Binding("strokeDashArray", FIELD.DOTTED_LINE_FLAG, findDottedLineReport),
        new go.Binding("fill", FIELD.VERSIONSTATUS, findNewItemColor),
        //new go.Binding("fill", FIELD.KEY, drawGraph),
        //new go.Binding("strokeWidth",FIELD.EDITABLE, findStrokeWidth),
          {
              name: "SHAPE_SKIN2", strokeWidth: 4,
              // set the port properties:
              portId: "", fromLinkable: true, toLinkable: true, cursor: "pointer"
          }),

      $(go.Panel, "Vertical",
         $(go.Picture,
            {
                name: 'UpArrow',
                desiredSize: new go.Size(25, 15),
                position: new go.Point(100, 0),
                source: "images/uparroww.png",
                visible: false

            },
              new go.Binding("visible", FIELD.KEY, isFirstElement)),
        $(go.Panel, "Horizontal",
          $(go.Picture,
            {
                name: 'Picture',
                alignment: go.Spot.Top,
                desiredSize: new go.Size(pictureWidth, pictureHeight),
                margin: new go.Margin(4, 4, 4, 4),

            },
            new go.Binding("source", FIELD.UNIQUEID, findHeadShot)),

          // define the panel where the text will appear
          $(go.Panel, "Table",
            {
                desiredSize: new go.Size(shape_X_Dimention, shape_Y_Dimention),
                margin: new go.Margin(4, 4, 4, 4),
                defaultAlignment: go.Spot.Left
            },

            $(go.RowColumnDefinition,
                    { row: 4, separatorStrokeWidth: 1, separatorStroke: "orange", separatorPadding: 6 }),

            $(go.TextBlock, textStyleTitle(skin),  // the name
              {
                  row: 0, column: 0, columnSpan: 4,
                  editable: true, isMultiline: false,
                  desiredSize: new go.Size(145, 16),
                  minSize: new go.Size(10, 16)
              },
          new go.Binding("text", FIELD.ORGUNITTEXT).makeTwoWay()),
           $(go.Picture,
            {
                name: 'ReportingError',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 5, columnSpan: 1,
                source: "images/error.jpg",
                visible: false
            },
            new go.Binding("visible", FIELD.REPORTINGERROR, isReportingError)),
            $(go.Picture,
            {
                name: 'Modified',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 6, columnSpan: 1,
                source: "images/user.ico",
                visible: false

            },
              new go.Binding("visible", FIELD.VERSIONSTATUS, isTrue)),
             $(go.Picture,
            {
                name: 'Deleted',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 6, columnSpan: 1,
                source: "images/deleted.png",
                visible: false

            },
              new go.Binding("visible", FIELD.VERSIONSTATUS, isDeleted)),
           $(go.Picture,
            {
                name: 'Editable',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 7, columnSpan: 1,
                source: "images/edit.png",
                visible: false

            },
              new go.Binding("visible", FIELD.EDITABLE, isTrue)),
           $(go.TextBlock, textStyle(skin),
               {
                   row: 1, column: 0, columnSpan: 7,
                   editable: false, isMultiline: false,
                   minSize: new go.Size(10, 16)
               },
              new go.Binding("text", FIELD.LEGALENTITY).makeTwoWay()),

            $(go.TextBlock, textStyle(skin),
              {
                  row: 2, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.TITLE).makeTwoWay()),

           $(go.TextBlock, textStyleTitle(skin),
              {
                  row: 3, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  width: 180, height: 16,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.NAME).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 2, columnSpan: 3,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.DIVISION).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 4, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right

              },
             "  NOR : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 5, columnSpan: 2,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.NOR_COUNT).makeTwoWay()),
       $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 2, columnSpan: 3,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.SITECODE).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 4, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right
              },
             "  SOC : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 5, columnSpan: 2,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.SOC_COUNT).makeTwoWay())


          )  // end Table Panel
        ), // end Horizontal Panel
          $(go.Picture,
            {
                name: 'DownArrow',

                desiredSize: new go.Size(25, 15),
                source: "images/downarrow.ico",
                alignment: go.Spot.Bottom,
                visible: false

            },
            new go.Binding("visible", FIELD.KEY, hasChildren))
       )
      );


    //var Legal =
    NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.LEGAL] =
        $(go.Node, "Auto",
        {
            contextMenu: nodeMenu
        },
        {
            doubleClick: drillDown
        },
        {
            // handle dragging a Node onto a Node to (maybe) change the reporting relationship
            mouseDragEnter: onMouseDragEnter,
            mouseDragLeave: onMouseDragLeave,
            mouseDrop: onMouseDrop
        },
        // for sorting, have the Node.text be the data.name
        new go.Binding("text", FIELD.NAME),
        // bind the Part.layerName to control the Node's layer depending on whether it isSelected
        new go.Binding("layerName", "isSelected", function (sel) { return sel ? "Foreground" : ""; }).ofObject(),

        // define the node's outer shape
        $(go.Shape, shape,
        new go.Binding("stroke", FIELD.EDITABLE, findBorderColorForPopulation),
        new go.Binding("strokeDashArray", FIELD.DOTTED_LINE_FLAG, findDottedLineReport),
       // new go.Binding("fill", FIELD.VERSIONSTATUS, findNewItemColor),
        new go.Binding("fill", FIELD.KEY, findPattern),
        //new go.Binding("strokeWidth",FIELD.EDITABLE, findStrokeWidth),
          {
              name: "SHAPE_SKIN2", strokeWidth: 4,
              // set the port properties:
              portId: "", fromLinkable: true, toLinkable: true, cursor: "pointer"
          }),
      $(go.Panel, "Vertical",
         $(go.Picture,
            {
                name: 'UpArrow',
                desiredSize: new go.Size(25, 15),
                position: new go.Point(100, 0),
                source: "images/uparroww.png",
                visible: false

            },
              new go.Binding("visible", FIELD.KEY, isFirstElement)),
        $(go.Panel, "Horizontal",
          $(go.Picture,
            {
                name: 'Picture',
                alignment: go.Spot.Top,
                desiredSize: new go.Size(pictureWidth, pictureHeight),
                margin: new go.Margin(4, 4, 4, 4),

            },
            new go.Binding("source", FIELD.UNIQUEID, findHeadShot)),
          // define the panel where the text will appear
          $(go.Panel, "Table",
            {
                desiredSize: new go.Size(shape_X_Dimention, shape_Y_Dimention),
                margin: new go.Margin(4, 4, 4, 4),
                defaultAlignment: go.Spot.Left
            },
            $(go.RowColumnDefinition,
                    { row: 4, separatorStrokeWidth: 1, separatorStroke: "orange", separatorPadding: 6 }),
            $(go.TextBlock, textStyleTitle(skin),  // the name
              {
                  row: 0, column: 0, columnSpan: 4,
                  editable: true, isMultiline: false,
                  desiredSize: new go.Size(145, 16),
                  minSize: new go.Size(10, 16)
              },
          new go.Binding("text", FIELD.ORGUNITTEXT).makeTwoWay()),
          $(go.Picture,
            {
                name: 'ReportingError',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 5, columnSpan: 1,
                source: "images/error.jpg",
                visible: false
            },
            new go.Binding("visible", FIELD.REPORTINGERROR, isReportingError)),
         $(go.Picture,
            {
                name: 'Modified',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 6, columnSpan: 1,
                source: "images/user.ico",
                visible: false
            },
            new go.Binding("visible", FIELD.VERSIONSTATUS, isTrue)),
          $(go.Picture,
            {
                name: 'Deleted',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 6, columnSpan: 1,
                source: "images/deleted.png",
                visible: false
            },
              new go.Binding("visible", FIELD.VERSIONSTATUS, isDeleted)),
           $(go.Picture,
            {
                name: 'Editable',
                desiredSize: new go.Size(15, 15),
                row: 0, column: 7, columnSpan: 1,
                source: "images/edit.png",
                visible: false
            },
            new go.Binding("visible", FIELD.EDITABLE, isTrue)),
           $(go.TextBlock, textStyle(skin),
               {
                   row: 1, column: 0, columnSpan: 7,
                   editable: false, isMultiline: false,
                   minSize: new go.Size(10, 16)
               },
              new go.Binding("text", FIELD.LEGALENTITY).makeTwoWay()),
            $(go.TextBlock, textStyle(skin),
              {
                  row: 2, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.TITLE).makeTwoWay()),
           $(go.TextBlock, textStyleTitle(skin),
              {
                  row: 3, column: 0, columnSpan: 7,
                  editable: false, isMultiline: false,
                  width: 180, height: 16,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.NAME).makeTwoWay()),
          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 2, columnSpan: 2,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.DIVISION).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 4, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right

              },
             " NOR : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 4, column: 5, columnSpan: 2,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.NOR_COUNT).makeTwoWay()),
       $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 2, columnSpan: 2,
                  editable: false, isMultiline: false,
                  minSize: new go.Size(10, 16)
              },
              new go.Binding("text", FIELD.OPL3TEXT).makeTwoWay()),

          $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 4, columnSpan: 1,
                  editable: false, isMultiline: false,
                  desiredSize: new go.Size(50, 16),
                  alignment: go.Spot.Right
              },
             " SOC : "),

         $(go.TextBlock, textStyle(skin),
              {
                  row: 5, column: 5, columnSpan: 2,
                  editable: false, isMultiline: false,
                  alignment: go.Spot.Left,
                  minSize: new go.Size(10, 16)
              },
             new go.Binding("text", FIELD.SOC_COUNT).makeTwoWay())


          )  // end Table Panel
        ), // end Horizontal Panel
          $(go.Picture,
            {
                name: 'DownArrow',

                desiredSize: new go.Size(25, 15),
                source: "images/downarrow.ico",
                alignment: go.Spot.Bottom,
                visible: false

            },
            new go.Binding("visible", FIELD.KEY, hasChildren))
       )
      );


    NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.DEFAULT] = NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.SIMPLE];

    //Assign Node Template
    //switch (userSettings.view)
    switch (skin) {
        case "1":
            switch (charttype) {
                case "LV":
                    myDiagram.nodeTemplate = NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.LEGAL];
                    break;
                case "OV":
                    myDiagram.nodeTemplate = NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.DETAILED];
                    break;
            }
            break;
        case "2":
            switch (charttype) {
                case "LV":
                    myDiagram.nodeTemplate = NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.LEGAL];
                    break;
                case "OV":
                    myDiagram.nodeTemplate = NODE_TEMPLATE_DEFINITION[NODE_TEMPLATE.CUSTOM];
                    break;
            }
            break;
    }

    // define the Link template
    myDiagram.linkTemplate =
      $(go.Link, go.Link.AvoidsNodes,  // may be either Orthogonal or AvoidsNodes
        { corner: 5, relinkableFrom: false, relinkableTo: false },
        $(go.Shape,
                //{ stroke: "#00a4a4" },
                new go.Binding("stroke", FIELD.OLEVEL, findLinkColorForLevel),
                 { strokeWidth: 2 }
         ));  // the link shape


    // read in the JSON-format data from the "mySavedModel" element
    load(null);
}



function inheritParentsProperty(childKey, parentKey) {
    var example = { key: parentKey };
    var parentNode = myDiagram.findNodesByExample(example).first();
    var parentData = parentNode.data
    if (parentNode instanceof go.Node && parentData !== null) {

        // Search for children's
        example = { key: childKey };
        var nodeIterator = myDiagram.findNodesByExample(example);
        var node = nodeIterator.first();
        if (node == null) return true;
        var data = node.data;
        if (node instanceof go.Node && data !== null) {
            var model = myDiagram.model;

            if (isOrgunit(data.key)) {
                // Do nothing
            } else {
                //Update the child position as they should inherit tha changes in the reporting orgunit
                if (isTrue(data.EDITABLE)) {
                    model.setDataProperty(data, FIELD.ORGUNITSHORTTEXT, parentData.ORGUNITSHORTTEXT);
                    model.setDataProperty(data, FIELD.ORGUNITTEXT, parentData.ORGUNITTEXT);
                    model.setDataProperty(data, FIELD.FDC, parentData.FDC);
                    model.setDataProperty(data, FIELD.FDCTEXT, parentData.FDCTEXT);
                    model.setDataProperty(data, FIELD.OPL3, parentData.OPL3);
                    model.setDataProperty(data, FIELD.OPL3TEXT, parentData.OPL3TEXT);

                    setModifiedStatusForNode(data.key, CONST.MODIFIED_STATUS_MODIFIED);
                }
            }
        }
    }
}
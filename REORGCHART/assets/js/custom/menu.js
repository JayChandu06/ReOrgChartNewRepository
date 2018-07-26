/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the client side configuration related to functioning of the Tree menu and Menu Search

 
***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

var tv = null;
var tvtemp = null;
var clipBoard = null;
var selectedNode = null;

var hash = {};
function processTable(data, idField, foreignKey, rootLevel) {
    for (var i = 0; i < data.length; i++) {
        var item = data[i];
        var id = item[idField];
        var parentId = item[foreignKey];

        hash[id] = hash[id] || [];
        hash[parentId] = hash[parentId] || [];

        item.items = hash[id];
        hash[parentId].push(item);
    }
    //alert(1);
    return hash[rootLevel];
}

function loadTreeOnDemand() {

    var homogeneous = new kendo.data.HierarchicalDataSource({
        transport: {
            read: {
                url: "RequestHandlers/TreeMenu.ashx",
                dataType: "jsonp"
                
            },
            parameterMap: function (data, type) {
                try {
                    if (data.ID) {
                        return {
                            "Method": "GetOrgChartChildList",
                            "Oper": $("#txtOper").val(),
                            "ShowLevel": $("#txtShowLevel").val(),
                            "KeyDate": $("#txtKeyDate").val(),
                            "SelLang": $("#txtSelLang").val(),
                            "LevelFlag": $("#txtLevelFlag").val(),
                            "VersionId": $("#txtVersion").val(),
                            "Country": $("#txtCountry").val(),
                            "ID": data.ID
                        };
                    } else {
                        return {
                            "Method": "GetOrgChartChildList",
                            "Oper": $("#txtOper").val(),
                            "ShowLevel": $("#txtShowLevel").val(),
                            "KeyDate": $("#txtKeyDate").val(),
                            "SelLang": $("#txtSelLang").val(),
                            "LevelFlag": $("#txtLevelFlag").val(),
                            "VersionId": $("#txtVersion").val(),
                            "Country": $("#txtCountry").val()
                        };
                    }

                } catch (e) {
                    return {
                        "Method": "GetOrgChartChildList",
                        "Oper": $("#txtOper").val(),
                        "ShowLevel": $("#txtShowLevel").val(),
                        "KeyDate": $("#txtKeyDate").val(),
                        "SelLang": $("#txtSelLang").val(),
                        "LevelFlag": $("#txtLevelFlag").val(),
                        "VersionId": $("#txtVersion").val(),
                        "Country": $("#txtCountry").val()
                    };
                }
            }
        },
        schema: {
            model: {
                id: "ID",
                hasChildren: "NOR"
            }
        }
    });


    $("#searchable-tree").kendoTreeView({
        dataSource: homogeneous,
        dataTextField: "text",
        dragAndDrop: true,
        loadOnDemand: true,
        select: treeviewSelect,
        drag: function (e) {
            var dataItem = this.dataItem(e.dropTarget);
            if (dataItem) {
                if (dataItem.expanded == true)
                    dataItem.set("expanded", false);
                else
                    dataItem.set("expanded", true);
            }
        }
    });

    tv = $("#searchable-tree").data('kendoTreeView');

}

function reloadTreeOnDemand() {
    //var optionalData = { "Oper": $("#txtOper").val(), "ShowLevel": $("#txtShowLevel").val(), "Country": $("#txtCountry").val(), "LevelFlag2": "2" };
    tv.dataSource.read();
}


function initSearchTree() {

    $("#tempsearch-div").kendoTreeView({
        dragAndDrop: true,
        loadOnDemand: true,
        select: treeviewSelect
    });

    tvtemp = $('#tempsearch-div').data('kendoTreeView');
}

function treeviewSelect(e) {
    var node = this.dataItem(e.node);
    if (!node) node =  this.dataItem(treeview.select());
    node.set("expanded", true);

    /*
    if (node) {
        $("#txtShowLevel").val(node.ID);
        sliderOnChange(e);
    }
    */
}


function initTreeContextMenu(Target) {
    var menu = $("#menu").kendoContextMenu({
        orientation: "vertical",
        target: Target,
        filter: ".k-in",
        animation: {
            open: { effects: "fadeIn" },
            duration: 500
        },
        select: function (e) {
            var button = $(e.item);
            var node = $(e.target);
            //alert(node.text());
            if (button.text() == "Load Workspace Chart") {
                if (node) {
                    $("#txtShowLevel").val(tv.dataItem(node).ID);
                    lefttreerefreshed = 1;
                    sliderOnChange(e);
                }
            }
            if (button.text() == "Load Pallete Chart") {
                if (node) {
                    $("#txtShowLevel").val(tv.dataItem(node).ID);
                    lefttreerefreshed = 0;
                    sliderOnChange(e);
                }
            }

            // you can get the node data (e.g. id) via the TreeView dataItem method:
            // $("#treeview").data("kendoTreeView").dataItem(node);
        }
    });
}


function CuteTreeNode() {
    var treeview = $('#searchable-tree').data('kendoTreeView');
    var selectedItem = $(".k-in.k-state-focused");

    if(selectedItem){
        clipBoard = selectedItem.closest(".k-item");
    } else {
            
            alert("Node not selected. Please try again!");
    }

    
    /*
    if (selectedNode) {
        clipBoard = $(this).closest(".k-item");
    } else {
        clipBoard = selectedNode;
    }
    */

    //clipBoard = treeview.select();
}


function PasteTreeNode() {
    var treeview = $('#searchable-tree').data('kendoTreeView');
    var selNode = $(".k-in.k-state-focused");

    if (!selNode)
        selNode = treeview.select();
    else
        selNode = selNode.closest(".k-item")

    /*
    var selNode = null;
    if (selectedNode) {
        selNode = $(this).closest(".k-item");
    } else {
        selNode = selectedNode;
    }
    */
    if (selNode && clipBoard) {
        treeview.append({
            text: clipBoard.text()
        }, selNode);
        treeview.remove(clipBoard);
        clipBoard = null;
    } else {
        alert("Node not selected. Please try again!");
    }
}


function initMenuPanel() {


    $('.vertical-menu > *').addClass('animated');

    $('.showMenu').on('click', function () {
        if ($(this).attr('title') == 'Search') {
            toogle_menu();

            if ($(this).hasClass('fadeInLeft')) {
                $('.vertical-menu > *').removeClass('fadeInLeft');
            } else {
                setTimeout(function () {
                    animate_menu();
                }, 70);
            }

            if ($(this).hasClass('search')) {
                $('.vertical-menu .form-item-search-block-form input').focus();

            }
        }
    });

    $('.vertical-menu').hover(
      function () {
          open_menu();
      },
      function () {
          // close_menu();
      });
}


function animate_menu() {
    $('.vertical-menu > *').addClass('fadeInLeft');
}

function toogle_menu() {
    if ($('.vertical-menu').hasClass('close')) {
        $('.vertical-menu').removeClass('close');
    } else {
        $('.vertical-menu').addClass('close');
    }
}
function open_menu() {
    if ($('.vertical-menu').hasClass('close')) {
        $('.vertical-menu').removeClass('close');
    }
}

function close_menu() {
    $('.vertical-menu').addClass('close');
}


function SearchTreeview() {

    if ($("#search-term").val().length == 0) {
        $("#searchable-tree").show();
        $("#tempsearch-div").hide();
        $("#searchable-message").hide();
    }


    if ($("#search-term").val().length > 3) {
        $("#searchable-tree").hide();
        $("#searchable-message").hide();
        $("#tempsearch-div").show();

        ShowProgressBar("#search-progress");

        $.ajax({
            type: "POST",
            url: "RequestHandlers/TreeMenu.ashx",
            data: JSON.stringify({ "Method": "SearchOrgList", "Oper": $("#txtOper").val(), "SearchString": $("#search-term").val(), "VersionId": $("#txtVersion").val() }),
            async: true,
            success: function (res) {
                hash = {};
                HideProgressBar("#search-progress");
                responseData = res;
                
                if(BlankString(responseData)){
                    HideProgressBar("#search-progress");
                    $("#tempsearch-div").hide();
                    $("#searchable-message").show();
                }
                processTable(responseData, "ID", "ParentId", -1);

                if (tvtemp)
                {
                    tvtemp.dataSource.data(hash[-1]);
                    //tvtemp.dataSource.read();
                }
                else
                {
                    $("#tempsearch-div").kendoTreeView({
                        dragAndDrop: true,
                        dataSource: hash[-1],
                        loadOnDemand: true,
                        select: treeviewSelect
                    });
                }

                tvtemp = $('#tempsearch-div').data('kendoTreeView');

                // highLightSearchResults("#search-term", "#tempsearch-div");

                expandNextLevel('#tempsearch-div');

                //tvtemp.collapse(".k-item");
                //tvtemp.expand(".k-item");
            },
            error: function (err) {
                HideProgressBar("#search-progress");
                $("#tempsearch-div").hide();
                $("#searchable-message").show();
            }
        });
    }
    else {
        alert("Please enter more than 3 charater!");
    }
    return false;
}


function highLightSearchResults(searchTextSelector, treeSelector) {
    
   // Tree menu Search functionality
    
        $(treeSelector+' li.k-item').show();
  

       $('span.k-in > span.highlight').each(function () {
           $(this).parent().text($(this).parent().text());
       });

       $('a.k-in > span.highlight').each(function () {
           $(this).parent().text($(this).parent().text());
       });

    // ignore if no search term
     /*
       if ($.trim($(searchTextSelector).val()) === '') {
           tvtemp.collapse(".k-item");
           return;
       }
       */
       var term = $(searchTextSelector).val().toUpperCase();
       var tlen = term.length;

       alert(term);
        /*
       if (tlen < 4) {
           return;
       }
       */
       $(treeSelector+' span.k-in').each(function (index) {
           var text = $(this).text();
           var html = '';
           var q = 0;
           var p;

           while ((p = text.toUpperCase().indexOf(term, q)) >= 0) {
               html += text.substring(q, p) + '<span class="highlight">' + text.substr(p, tlen) + '</span>';
               q = p + tlen;
           }

           if (q > 0) {
               html += text.substring(q);
               $(this).html(html);

               $(this).parentsUntil('.k-treeview').filter('.k-item').each(function (index, element) {
                   tvtemp.expand($(this));
                   $(this).data('SearchTerm', term);
               });
           }
       });

       $(treeSelector+' a.k-in').each(function (index) {
           var text = $(this).text();
           var html = '';
           var q = 0;
           var p;

           while ((p = text.toUpperCase().indexOf(term, q)) >= 0) {
               html += text.substring(q, p) + '<span class="highlight">' + text.substr(p, tlen) + '</span>';
               q = p + tlen;
           }

           if (q > 0) {
               html += text.substring(q);
               $(this).html(html);

               $(this).parentsUntil('.k-treeview').filter('.k-item').each(function (index, element) {
                   tvtemp.expand($(this));
                   $(this).data('SearchTerm', term);
               });
           }
       });

       $(treeSelector+' li.k-item:not(:has(".highlight"))').hide();

       tvtemp.expand(".k-item");

}


function expandNextLevel(treeViewSelector) {
    setTimeout(function () {
        var treeview = $(treeViewSelector).data("kendoTreeView");
        var b = $('.k-item .k-plus').length;
        treeview.expand(".k-item");
        treeview.trigger('dataBound');
        if (b > 0) {
            expandNextLevel(treeViewSelector);
        }

    }, 200);
}


function tohgle_h_menu() {
   
    if ($('#horizonal-menu').hasClass('close')) {
        $('#horizonal-menu').removeClass('close');
    } else {
        $('#horizonal-menu').addClass('close');
    }
    reload();
}
function open_h_menu() {
    if ($('#horizonal-menu').hasClass('close')) {
        $('#horizonal-menu').removeClass('close');
    }
}

function close_h_menu() {
    $('#horizonal-menu').addClass('close');
    
}

function initChartSelector() {
    $("#select-chart").kendoMobileButtonGroup({
        select: function (e) {
            //kendoConsole.log("selected index:" + e.index);

            if (e.index == 0) {
                toggleChartType(CHART_TYPE.OPERATIONAL);
            } else {
                toggleChartType(CHART_TYPE.LEGAL);
            }

        },
        index: 0
    });
}
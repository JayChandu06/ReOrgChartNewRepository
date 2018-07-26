/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file invokes all the initialization scripts together in proper sequence 

 * Init Chart selector - to toggle between Legal or Operational chart
 * Initialize the slider - to select orgchart depth : default 1
 * Load Tree menu
 * Load Left Tree Context Menu
 * Initialize OrgChart
 * Initialize Lookup Menus
 * Bind window Resize scripts to enable dynamic sizing based on window resize 

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

$(document).ready(function () {

    //// Use Select2
    //$('select').select2();

    ////init Chart selection switch to select operational or legal chart
    //initChartSelector();

  //  // Depth Selector 
  //  initSlider("slider1", defaultval);
  ////  initSlider("slider2", defaultval);

  //  // Initialize menu panel
  //  initMenuPanel();

  //  // Load lefft tree
  //  loadTreeOnDemand();
  //  initSearchTree();

  //  // initialize context menu
  //  var conextMenu = $("#menu"),
  //  original = conextMenu.clone(true);
  //  original.find(".k-state-active").removeClass("k-state-active");

  //  initTreeContextMenu("#searchable-tree");
  //  initTreeContextMenu("#tempsearch-div");

  //  $('#search-term').on('keyup', function () {

  //      var term = this.value.toUpperCase();
  //      var tlen = term.length;

  //      if (tlen == 0) {
  //          $("#searchable-tree").show();
  //          $("#tempsearch-div").hide();
  //          $("#searchable-message").hide();
  //      }
  //  });
    
    
    $(window).bind('beforeunload', function () {
        return MESSAGE.CONFIRM_MOVE_OUT;
    });
    
    $(window).resize(onresize);
    onresize(1);

    //initialize orchart
    initOrgChart();
    
    //// Handle URl parameter on load
    //initLookups();


});

function onresize(init) {

    var h = window.innerHeight || document.documentElement.clientHeight || document.body.clientHeight;
    var w = window.innerWidth || document.documentElement.clientWidth || document.body.clientWidth;

    //if (w < 768) close_h_menu();
    //else open_h_menu();

    if (init && init == 1)
        return true;
    else {
        //$('select').select2();
        reload();
    }
}
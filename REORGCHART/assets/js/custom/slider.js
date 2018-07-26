/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the client side configuration related to functioning of the Slider menu

 
***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

var defaultval = 1;

function initSlider(sliderDiv,defaultval) {

    $("#" + sliderDiv).kendoSlider({
        change: sliderOnChange,
        slide: sliderOnSlide,
        min: 1,
        max: 4,
        smallStep: 1,
        largeStep: 1,
        value: defaultval,
        tickPlacement: "both",
        showButtons: false
    });
}

// Slider event handlers

function sliderOnSlide(e) {
    //kendoConsole.log("Slide :: new slide value is: " + e.value);
}

function sliderOnChange(e) {

    if (e.value) {
        $("#txtLevelFlag").val(e.value);
    }

    loadOrgChartData();
}



/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the helper methods to initialize/bind the Dropdown menus with appropriate methods in HRMasterData request handler

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

function loadHRDataSelector(elmSelecter, method, defaultVal, itemTemplateSelector,valueTemplateSelector) {

    var version =  $("#txtVersion").val();

    $(elmSelecter).kendoDropDownList({
        template: $(itemTemplateSelector).html(),
        valueTemplate: $(valueTemplateSelector).html(),
        placeholder: "Select value...",
        optionLabel: {
            Title: "Select value...",
            ID: "0"
        },
        filter: "contains",
        dataTextField: "Title",
        dataValueField: "ID",
        minLength: 3,
        height: 220,
        highlightFirst: true,
        autoBind: false,
        text: defaultVal,
        suggest: true,
        dataSource: {
            pageSize: 7,
            serverPaging: false,
            serverFiltering: true,
            transport: {
                read: {
                    dataType: "json",
                    url: "RequestHandlers/HRMasterData.ashx"
                },
                parameterMap: function (data, type) {
                    try{
                        return {
                            "Method": method,
                            "VersionId": $("#txtVersion").val(),
                            "SearchString": data.filter.filters[0].value
                        }
                    } catch (e) {
                        return {
                            "Method": method,
                            "VersionId": $("#txtVersion").val(),
                            "SearchString": ""
                        }
                    }
                }
            }
        }
    });

}
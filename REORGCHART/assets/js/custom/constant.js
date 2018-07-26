/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This script file defines the application contants used at cleint side and is complimentary to the AppConstant.cs class.
  

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/

var SIZE_FULLSCREEN = {
    "NODE_SPACING": "30",
    "NODE_SCALE": "1"
}
var SIZE_TAB = {
    "NODE_SPACING": "20",
    "NODE_SCALE": "0.8"
}
var SIZE_MOBILE = {
    "NODE_SPACING": "15",
    "NODE_SCALE": "0.8"
}

var CHART_TYPE = { "OPERATIONAL": "OV", "LEGAL": "LV" }


var CONST = {
    "MODIFIED_STATUS": "CLINETMODIFIEDSTATUS",
    "MODIFIED_STATUS_NEWPERSONASSIGNEDTOPOSITION": "6",
    "MODIFIED_STATUS_PERSONUNASSIGNEDFROMPOSITION": "5",
    "MODIFIED_STATUS_HASNEWCPM": "4",
    "MODIFIED_STATUS_WASOLDCPM": "3",
    "MODIFIED_STATUS_CREATED": "2",
    "MODIFIED_STATUS_MODIFIED": "1",
    "MODIFIED_STATUS_RELINKED": "1",
    "MODIFIED_STATUS_DELETED": "-1",
    "MODIFIED_STATUS_NOCHANGE": "0",
    "IMAGE_PATH": "images/Photos/",
    "IMAGE_EXTENSION": ".jpg",
    "NOIMAGE": "NOIMAGE"
};




var SKIN = ["DEFAULT","NOVARTIS", "WHITE"];

var SKIN_COLOR = [
                    ["#634329/#634329", "#923222/#923222", "#e44c16/#e44c16", "#ec8026/#ec8026", "#fcaf17/#fcaf17", "#fed300/#fed300", "#f5eb07/#f5eb07", "#e44c16/#e44c16" ],
                    ["#634329/#634329", "#923222/#923222", "#e44c16/#e44c16", "#ec8026/#ec8026", "#fcaf17/#fcaf17", "#fed300/#fed300", "#f5eb07/#f5eb07", "#e44c16/#e44c16" ],
                    ["#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF", "#FFFFFF/#FFFFFF"]
];



var MODIFIED_FLEG = { "TRUE": "1", "FALSE": "0" };
var SPLIT_SCREEN = { "TRUE": "1", "FALSE": "0" };
var SPLIT_SCREEN_DIRECTION = { "VERTICAL": "0", "HORIZONTAL": "1" }


var VIEW = ["DEFAULT", "SIMPLE", "DETAILED", "CUSTOM"]; // Default is Simple
var SHAPE = ["DEFAULT", "RoundRectanle", "Circle", "Ellipse"]; // Default is RoundRectangle

var FIELD = {
    "KEY": "key",
    "PARENT": "parent",
    "POSITIONID": "POSITIONID",
    "NEXTPOSITIONID": "NEXTPOSITIONID",

    "UNIQUEID": "GDDBID",
    "GDDBID": "GDDBID",
    "ORGUNIT": "ORGUNIT",
    "ORGUNITTEXT": "ORGUNITTEXT",
    "LEGALENTITY": "LEGALENTITY",
    "TITLE": "TITLE",
    "SITECODE": "SITECODE",
    "DIVISION": "DIVISION",

    "NAME": "NAME",
    "DOTTED_LINE_FLAG": "DOTTED_LINE_FLAG",
    "NOR_COUNT": "NOR_COUNT",
    "SOC_COUNT": "SOC_COUNT",

    "ORGUNITSHORTTEXT":"ORGUNITSHORTTEXT",
    "ORGUNITTEXT": "ORGUNITTEXT",
    "FDC": "FDC",
    "FDCTEXT": "FDCTEXT",
    "OPL3": "OPL3",
    "OPL3TEXT":"OPL3TEXT",

    "CPM": "CPM",
    "CPMLEGALENTITY": "CPMLEGALENTITY",
    "OLEVEL":"OLEVEL",
    "OPM": "OPM",
    "OPMLEGALENTITY": "OPMLEGALENTITY",

    "EDITABLE": "EDITABLE",

    "MODIFIEDSTATUS": "CLINETMODIFIEDSTATUS",
    "HRCORESTATUS": "HRCORESTATUS",
    "VERSIONSTATUS": "VERSIONSTATUS",
    "REPORTINGERROR": "REPORTINGERROR",

    "PERSON": "PERSON",

    "PERSON_CPM": "PERSON_CPM",
    "ORGUNIT_CPM": "ORGUNIT_CPM",
    "POSITIONID_CPM": "POSITIONID_CPM",
    "TITLE_CPM": "TITLE_CPM",
    "TITLE_PERSON":"TITLE_PERSON",
    "HDN_ORGUNIT_CPM": "HDN_ORGUNIT_CPM",
    "HDN_POSITIONID_CPM":"HDN_POSITIONID_CPM",
    "HDN_PERSON_GDDBID": "HDN_PERSON_GDDBID",
    "HDN_KEY":"HDN_KEY"
};

var KEY = {
    "ROOT": "-1",
    "ORPHAN": "-2",
    "INITIALVERSIONTEXT": "Initial Population",
    "HRVERSIONTEXT": "HR CORE VERSION",
    "BLANK_POSITIONHOLDERSID": "GDDBID",
    "BLANK_POSITIONHOLDERSNAME": "Unassigned position",
    "BLANK_ORGUNITPOSITIONID": "POSITIONID",
    "UNASSIGNED_POSITIONID": "00000000",
    "NO_POSITION":"Person without position"
}

var LOOKUPNAME = {
    "COUNTRY":"COUNTRY",
    "INITIATIVE": "INITIATIVE",
    "POPULATION": "POPULATION",
    "USER": "USER",
    "VERSION": "VERSION"
}
    

var MESSAGE = {
    "DELETION_NOT_ALLOWED": "CUT/DELETE on selected position is not allowed.",
    "NONEDITABLE": "Selected position is not editable.",
    "REPORTING_NOT_ALLOWED_OPERATIONAL": "Position cannot report to the selected operational manager because they form a recursive loop.",
    "REPORTING_NOT_ALLOWED_LEGAL": "Position/Orgunit cannot be moved under selected orguint because either they form a recursive loop or the target selection is a position.",
    "RECURSIVE_LOOP": "Position/Orgunit cannot report to the selected position/orgunit because they form a recursive loop.",
    "REPORTING_NOT_ALLOWED_TO_POSITION": "Position/Orgunit cannot be moved under a position.",
    "SAVED": "The changes are saved successfully.",
    "BEFORE_SAVE": "You are about to save the changes!",
    "SAVE_NOT_ALLOWED": "You are allowed to save only your versions!",
    "CONFIRM_MOVE_OUT": "By moving out of the current dataset you will lose all the changes! Please cancel and save the changes before navigating out.",
    "VALIDATE_BLANK_INITIATIVE_TITLE": "Please provide an Initiative title.",
    "VALIDATE_UNIQUE_INITIATIVE_TITLE": "An Initiative with same title already exists.",
    "VALIDATE_BLANK_POPULATION_TITLE": "Please provide population Title.",
    "VALIDATE_UNIQUE_POPULATION_TITLE": "A population with the same title already exists under the current initiative.",
    "VALIDATE_VERSION_SELECTION": "Please select a population & version to be copied.",
    "VALIDATE_UNIQUE_VERSION_TITLE": "A version with the same title already exists under the current population for you.",
    "VALIDATE_VERSION_SELECTION_DOWNLOAD": "Please select a population & version to be downloaded.",
    "VALIDATE_INITIATIVE_SELECTION_DOWNLOAD": "Please select a initiative to be downloaded.",
    "VALIDATE_POPULATION_SELECTION_DELETE": "Please select a population & version to be deleted.",
    "CONFIRM_POPULATION_DELETION": "Deleting the population will also delete all the associated versions! Are you sure you want to delete the selected population?",
    "CONFIRM_VERSION_DELETION": "Deleting the Version will also delete all the changes made in the versions! Are you sure you want to delete the selected version?",
    "CPM_OPM_RULE_VIOLATED": "Warning: OPM should equal CPM when they are located in the same Legal Entity as the associate.  Note the CPM will need to be re-assigned to equal OPM prior to upload to HR Core . Please click  ok to continue",
    "VALIDATE_BLANK_BLANK_COUNTRY":"Please select a valid country to create an initiative under it."
}

var STATUS =
    {
        "TRUE": "SUCCESSFUL",
        "FALSE": "FAILED"
    }


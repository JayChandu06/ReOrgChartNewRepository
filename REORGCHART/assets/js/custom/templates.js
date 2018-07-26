var FIELDS = {
    "KEY": "key",
    "PARENT": "parent",
    "NAME": "FullName"
}

// This function provides a common style for most of the TextBlocks.
// Some of these values may be overridden in a particular TextBlock.
function textStyle(skin) {
    switch (skin) {
        case "1":
            return { font: "8pt  Segoe UI,sans-serif", stroke: "white" };
            break;
        case "2":
            return { font: "8pt  Segoe UI,sans-serif", stroke: "black" };
            break;
    }
    return { font: "8pt  Segoe UI,sans-serif", stroke: "white" };
}
function textStyleTitle(skin) {
    switch (skin) {
        case "1":
            return { font: "11pt  Segoe UI,sans-serif", stroke: "white" };
            break;
        case "2":
            return { font: "11pt  Segoe UI,sans-serif", stroke: "black" };
            break;
    }
    return { font: "11pt  Segoe UI,sans-serif", stroke: "white" };
}

// This converts dotted line to a different shape
function findLinkColorForLevel(key) {
    //alert(key);
    switch (key) {
        case "0":
            return "#00a4a4";
            break;
        case "1":
            return "#e44c16";
            break;
        case "2":
            return "#f5eb07";
            break;
        case "3":
            return "#fed300";
            break;
        case "4":
            return "#fcaf17";
            break;
        case "5":
            return "#ec8026";
            break;
        case "6":
            return "#e44c16";
            break;
        case "7":
            return "#923222";
            break;
        case "8":
            return "#634329";
            break;
    }

    return "#00a4a4";
};

function findBorderColorForLevel(key) {
    switch (key) {
        case "0":
            return "#00a4a4";
            break;
        case "1":
            return "#a4a400";
            break;
        case "2":
            return "#a400a4";
            break;
        case "3":
            return "#fed300";
            break;
        case "4":
            return "#fcaf17";
            break;
        case "5":
            return "#ec8026";
            break;
        case "6":
            return "#e44c16";
            break;
        case "7":
            return "#923222";
            break;
        case "8":
            return "#634329";
            break;
    }

    return "#00a4a4";
};

var NODE_TEMPLATE = {
    "DEFAULT": "0",
    "SIMPLE": "1",
    "DETAILED": "2",
    "CUSTOM": "3",
    "LEGAL":"4"
};

var NODE_TEMPLATE_DEFINITION = [];


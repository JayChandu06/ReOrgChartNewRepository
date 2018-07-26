
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

    return hash[rootLevel];
}

$(document).ready(function () {
    var tv;
    $.ajax({
        type: "POST",
        url: "RequestHandlers/TreeMenu.ashx",
        data: JSON.stringify({ "Method": "GetMenuList" }),
        async: true,
        success: function (res) {
            responseData = res;
            processTable(responseData, "ID", "ParentId", 0);

            $("#searchable-tree").kendoTreeView({
                dataSource: hash[0],
                dataUrlField: "LinksTo",
                loadOnDemand: false,
                select: treeviewSelect
            });

            tv = $('#searchable-tree').data('kendoTreeView');
        },
        error: function (err) {
            //alert(err);
        }
    });

    function treeviewSelect(e) {
        var node = this.dataItem(e.node);
        if (node.IsNewWindow) {
            window.open(node.LinksTo, '_blank');

        }
        else {
            window.open(node.LinksTo, '_self');
        }
    }


    var menu = $("#menu"),
    original = menu.clone(true);

    original.find(".k-state-active").removeClass("k-state-active");

    var initTreeContextMenu = function () {

        menu = $("#menu").kendoContextMenu({
            orientation: "vertical",
            target: "#bodyContent",
            filter: ".columns",

            animation: {
                open: { effects: "fadeIn" },
                duration: 500
            },
            select: function (e) {
                // Do something on select
            }
        }).data("kendoContextMenu");
    };

    initTreeContextMenu();

    //Search functionality
    $('#search-term').on('keyup', function () {

        $('#searchable-tree li.k-item').show();

        $('span.k-in > span.highlight').each(function () {
            $(this).parent().text($(this).parent().text());
        });

        $('a.k-in > span.highlight').each(function () {
            $(this).parent().text($(this).parent().text());
        });

        // ignore if no search term
        if ($.trim($(this).val()) === '') {
            tv.collapse(".k-item");
            return;
        }

        var term = this.value.toUpperCase();
        var tlen = term.length;

        $('#searchable-tree span.k-in').each(function (index) {
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
                    tv.expand($(this));
                    $(this).data('SearchTerm', term);
                });
            }
        });

        $('#searchable-tree a.k-in').each(function (index) {
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
                    tv.expand($(this));
                    $(this).data('SearchTerm', term);
                });
            }
        });

        $('#searchable-tree li.k-item:not(:has(".highlight"))').hide();

        tv.expand(".k-item");
    });

});



function OpenContextMenu(obj) { var position = $(obj).offset(); $("#menu").data("kendoContextMenu").open(position.left + 30, position.top + 15); }



function close_menu() {
    $('.vertical-menu').addClass('close');
}


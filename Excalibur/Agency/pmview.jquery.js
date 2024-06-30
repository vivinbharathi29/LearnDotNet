(function ($) {
    $.fn.isNullOrEmpty = function () {
        return (this === null || this === undefined || this === '');
    };
})(jQuery);

$(function () {

    //Setup Table Format
    $(".StatusTable th").addClass("ui-state-default");
    $(".StatusTable td").addClass("ui-widget-content");

    $(".StatusTable thead th").hover(function () {
        $(this).addClass("hover");
    }, function () {
        $(this).removeClass("hover");
    });

    $(".StatusTable tr").hover(function () {
        $(this).children("td").addClass("ui-state-hover2");
        $(this).addClass("hover");
    }, function () {
        $(this).children("td").removeClass("ui-state-hover2");
        $(this).removeClass("hover");
    });

    $(".StatusTable tbody tr").click(function () {
        $('#txtProjectedDate').attr("disabled", true);
        var row = $(this);
        var statusId = $(".statusId", row).val();
        if (!$(statusId).isNullOrEmpty())
            DetailsDialog.Show(statusId);
    });

    var showPastDueWarning = false;
    var showInitializationWarning = false;

    if ($(".StatusTable tbody.Uninitialized").length > 0) {
        showInitializationWarning = true;

        $(".StatusTable tbody.Uninitialized tr td").addClass("ui-state-highlight");
        $(".StatusTable tbody.Uninitialized tr td:nth-child(4)").text("");

    }

    $(".StatusTable tbody.In_Progress tr").each(function () {
        var daysToTargetCell = $("td:nth-child(4)", $(this));
        var days = parseInt($(daysToTargetCell).text());
        if (days < 7) {
            $(daysToTargetCell).parent().children().addClass("ui-state-highlight");
        }
        if (days < 0) {
            showPastDueWarning = true;
            $(daysToTargetCell).parent().children().addClass("ui-state-error");
            $(daysToTargetCell).text("Past Due");
        }
    });

    $(".StatusTable tbody.Complete tr").each(function () {
        if (parseInt($(".nextStepId", $(this)).val()) == 0) {
            daysToTargetCell = $("td:nth-child(4)", $(this)).text("Complete");
        }
    });

    $(".StatusTable thead").click(function () {
        $(this).next().toggle();
    });

    if ($(".StatusTable").length > 1) {
        $(".StatusTable tbody").hide();
    }

    if (showPastDueWarning) {
        $("#errorText").append("You have workflow steps that are past due!");
    }

    if (showInitializationWarning) {
        $("#warningText").append("You have certifications that need to be initialized.");
    }

    if ($("#errorText").text() == '') {
        $("#error").hide();
    }

    if ($("#warningText").text() == '') {
        $("#warning").hide();
    }

    $(".Link").hover(function () {
        $(this).addClass("hover");
    }, function () {
        $(this).removeClass("hover");
    });

    $("#initWorkflow").click(function () {
        InitDialog.Show();
    });

    $("#showFrame").click(function () {
        ShowIframeDialog();
    });

    $("#iframeDialog").dialog({
        modal: true,
        autoOpen: false,
        width: 800,
        height: 800
    });

    InitDialog.Setup();

    //Hide Loading panel and show the rest of the body.
    $("#loading").hide();
    $("#body").fadeIn('slow');
});

function ShowIframeDialog() {
    $("#iframeDialog iframe").attr("width", "95%");
    $("#iframeDialog iframe").attr("height", "95%");
    $("#iframeDialog iframe").attr("src", "Agency.asp?ID=5000");
    $("#iframeDialog").dialog("open");
}

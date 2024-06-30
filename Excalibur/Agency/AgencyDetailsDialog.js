
$(function () {
    DetailsDialog.Setup();
    ChangeLeadTimeDialog.Setup();
});

var DetailsDialog = function () { }
DetailsDialog.Setup = function () {
    // Dialog
    $("#AgencyDetailsDialog").dialog({
        modal: true,
        autoOpen: false,
        width: 350,
        height: 450,
        show: "clip",
        buttons: {
            "Ok": function () {
                var response = DetailsDialog.Save();
                if (response) {
                    $(this).dialog("close");
                    location.reload();
                } else {
                    alert("Something went wrong.");
                }
            },
            "Cancel": function () {
                $(this).dialog("close");
            }
        }
    });

    $("#wfStatus th").each(function () {
        $(this).addClass("ui-state-default");
    });

    $("#wfStatus td").each(function () {
        $(this).addClass("ui-widget-content");
    });

    // Datepicker
    $('#txtProjectedDate').datepicker();
}

DetailsDialog.Show = function (AgencyStatusId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/GetAgencyDetails",
        data: "{ AgencyStatusId: " + AgencyStatusId + " }",
        dataType: "json",
        async: false,
        success: function (msg) {
            var table = ("#AgencyDetailsDialog");
            var r = $.parseJSON(msg.d);
            $("#AgencyStatusId", table).val(r.agencystatusid);
            $("#lblDocument", table).text(r.agencyname);
            $("#lblWorkflow", table).text(r.workflowname);
            $("#lblCurrentStep", table).text(r.stepname);
            $("#lblTargetDt", table).text(r.targetdate);
            $("#lblDaysToTarget", table).text(r.daystotarget);
            $("#txtNotes", table).val(r.notes);
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Initializing Status\n" + json.Message);
            initDialog.saveErrorThrown = true;
        }

    });

    DetailsDialog.PopulateWorkflowTable(AgencyStatusId);

    $("#txtProjectedDate").attr("disabled", true);
    $("#AgencyDetailsDialog").dialog("open");
    $("#txtProjectedDate").removeAttr("disabled");
}

DetailsDialog.PopulateWorkflowTable = function (AgencyStatusId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/GetAgencyWorkflowStatusByAgencyStatusId",
        data: "{ AgencyStatusId: " + AgencyStatusId + " }",
        dataType: "json",
        async: false,
        success: function (msg) {
            var table = ("#AgencyDetailsDialog #wfStatus");
            $("tbody", table).find("tr").remove()
            var dt = $.parseJSON(msg.d);
            var bEditModeOn = ($("#isAgencyEditModeOn").val() == "True");
            var statusValue;
            var targetDate;
            for (var i = 0; i < dt.length; i++) {
                var dr = dt[i];

                statusValue = '';
                if (dr.CompletionDate == '') {
                    if (dr.CurrentWorkflowStepId == dr.Id && bEditModeOn && (dr.LeveragedStatusId == 0 || dr.LeveragedStatusId == '')) {
                        statusValue = '<span id="workflowStatusComplete" class="Link"><input type="hidden" id="workflowStepId" value="' + dr.Id + '" />Complete</span>';
                    }
                } else if ((dr.CurrentWorkflowStepId - 1) == dr.Id && dr.LeveragedStatusId == '') {
                    statusValue = '<span id="workflowRollback" class="Link"><input type="hidden" id="rollbackStepId" value="' + dr.Id + '" />Rollback</span>'
                } else {
                    statusValue = $.datepicker.formatDate("mm/dd/yy", new Date(dr.CompletionDate));
                }

                targetDate = '';
                targetDate = '<span class="Link TargetLink"><input type="hidden" class="workflowStepId" value="' + dr.Id + '" />' + $.datepicker.formatDate("mm/dd/yy", new Date(dr.TargetDate)) + '</span>';

                $("tbody", table).append('<tr id="' + dr.Id + '"><td>' + dr.StepName + '</td>' +
                            '<td>' + targetDate + '</td>' +
                            '<td style="text-align:right;">' + dr.DaysToTarget + '&nbsp;&nbsp;&nbsp;</td>' +
                            '<td>' + statusValue + '</td></tr>');
            }

        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Initializing Status\n" + json.Message);
            initDialog.saveErrorThrown = true;
        }

    });

    var dialog = $("#AgencyDetailsDialog");

    $(".Link", dialog).hover(function () {
        $(this).addClass("hover");
    }, function () {
        $(this).removeClass("hover");
    });

    $("#workflowStatusComplete", dialog).click(function () {
        var span = $(this);
        var workflowId = $("#workflowStepId", span).val();
        DetailsDialog.MarkWorkflowComplete(workflowId);
    });

    $("#workflowRollback", dialog).click(function () {
        var span = $(this);
        var workflowId = $("#rollbackStepId", span).val();
        DetailsDialog.RollbackWorkflowStep(workflowId);
    });

    $(".TargetLink", dialog).click(function () {
        var workflowId = $(".workflowStepId", this).val();
        ChangeLeadTimeDialog.Show(workflowId);
    });
}

DetailsDialog.RollbackWorkflowStep = function (workflowId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/RollbackAgencyWorkflowStatus",
        data: "{ AgencyWorkflowStatusId: " + workflowId + " }",
        dataType: "json",
        async: false,
        success: function (msg) {
            var r = $.parseJSON(msg.d);
            var dialog = $("#AgencyDetailsDialog");
            var agencyStatusId = $("#AgencyStatusId", dialog).val();
            DetailsDialog.PopulateWorkflowTable(agencyStatusId);
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Updating Workflow Status\n" + json.Message);
        }
    });
}

DetailsDialog.MarkWorkflowComplete = function (workflowId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/SetAgencyWorkflowStatusComplete",
        data: "{ AgencyWorkflowStatusId: " + workflowId + " }",
        dataType: "json",
        async: false,
        success: function (msg) {
            var r = $.parseJSON(msg.d);
            var dialog = $("#AgencyDetailsDialog");
            var agencyStatusId = $("#AgencyStatusId", dialog).val();
            DetailsDialog.PopulateWorkflowTable(agencyStatusId);
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Updating Workflow Status\n" + json.Message);
        }
    });
}

DetailsDialog.Save = function () {
    return $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/SaveAgencyStatus",
        data: "{AgencyStatusId: '" + $("#AgencyStatusId").val() + "'" +
            ",Notes: '" + $("#txtNotes").val() + "'" +
            "}",
        dataType: "json",
        success: function (msg) {
            return true;
        },
        error: function (msg) {
            return false;
        }
    });
}

var ChangeLeadTimeDialog = function () { };

ChangeLeadTimeDialog.Setup = function () {
    $("#ChangeLeadTime").dialog({
        modal: true,
        autoOpen: false,
        width: 350,
        height: 200,
        show: 'clip',
        buttons: {
            "Save": function () {
                if (ChangeLeadTimeDialog.Save())
                    $(this).dialog("close");
            },
            "Cancel": function () {
                $(this).dialog("close");
            }
        },
        open: function (event, ui) {
            $('#txtNewDueDt').removeAttr('disabled')
        },
        close: function (event, ui) {
            if ($.datepicker._datepickerShowing) {
                $('#txtNewDueDt').datepicker('hide');
            }
            $('#txtNewDueDt').attr("disabled", true);
        }
    });

    $("#ChangeLeadTime #txtNewDueDt").datepicker({
        showOn: 'both',
        buttonImage: '/Excalibur/images/calendar.gif',
        buttonImageOnly: true
    });
}

ChangeLeadTimeDialog.Show = function (AgencyStatusWorkflowId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/AgencyStatusWorkflowGetById",
        data: "{ AgencyStatusWorkflowId: " + AgencyStatusWorkflowId + " }",
        dataType: "json",
        async: false,
        success: function (msg) {
            var r = $.parseJSON(msg.d);
            $("#ChangeLeadTime #CurrentTargetDt").text($.datepicker.formatDate("mm/dd/yy", new Date(r.targetdt)));
            $("#ChangeLeadTime #CurrentDueDt").text($.datepicker.formatDate("mm/dd/yy", new Date(r.duedt)));
            $("#ChangeLeadTime #txtNewDueDt").val('');
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Initializing Status\n" + json.Message);
            initDialog.saveErrorThrown = true;
        }
    });
    $("#ChangeLeadTime #hidWorkflowId").val(AgencyStatusWorkflowId);
    $("#ChangeLeadTime #txtNewDueDt").attr("disabled", true);
    $("#ChangeLeadTime").dialog('open');
}

ChangeLeadTimeDialog.Save = function () {
    return $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/SetAgencyStatusWorkflowLeadTime",
        data: "{AgencyStatusWorkflowId: '" + $("#ChangeLeadTime #hidWorkflowId").val() + "'" +
            ",TargetDt: '" + $("#ChangeLeadTime #CurrentTargetDt").text() + "'" +
            ",NewDt: '" + $("#ChangeLeadTime #txtNewDueDt").val() + "'}",
        dataType: "json",
        success: function (msg) {
            DetailsDialog.PopulateWorkflowTable($("#AgencyDetailsDialog #AgencyStatusId").val());
            return true;
        },
        error: function (msg) {
            alert(msg.responseText);
            return false;
        }
    });
}


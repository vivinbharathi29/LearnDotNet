$(function () {
    InitDialog.Setup();
    Leverage.Setup();
    LeverageVersion.Setup();
});

var InitDialog = function () {
    this.workflowJson = null;
    this.saveErrorThrown = false;
};
var initDialog = new InitDialog();

InitDialog.Show = function (DeliverableVersionId, AgencyTypeId) {
    //$("#AgencyInitWorkflowDialog tbody").find("tr").remove();
    $("#AgencyInitWorkflowDialog").dialog("open");
    InitDialog.Load(DeliverableVersionId, AgencyTypeId);
}

InitDialog.Setup = function () {
    $("#AgencyInitWorkflowDialog").dialog({
        modal: true,
        autoOpen: false,
        width: 900,
        height: 600,
        show: 'clip',
        buttons: {
            "Initialize": function () {
                $("#AgencyInitWorkflowDialog").block();
                
                InitDialog.Save();
                //InitDialog.Load();

                $("#AgencyInitWorkflowDialog").unblock();
            },
            "Close": function () {
                $(this).dialog("close");
            }
        }
    });

    //    var wfDialog = $("#AgencyInitWorkflowDialog");
    //    $("#leverageSearch", wfDialog).click(function() {
    //        InitDialog.ShowLeverage();
    //    });
};

InitDialog.Save = function () {
    var table = $('#AgencyInitWorkflowDialog table');
    $("tr", table).each(function () {
        var id = parseInt($(this).attr("id"));
        if (id > 0 && !initDialog.saveErrorThrown) {
            var selectedWorkflow = $("#selMilestone" + id, $(this)).val();
            var selectedMilestone = 93;
            var selectedDate = $("#dp" + id, $(this)).val();
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: "/ExcaliburApi/AgencyCertification.asmx/InitializeAgencyStatusWorkflow",
                data: "{ AgencyStatusId: '" + id + "', " +
                    "AgencyWorkflowId: '" + selectedWorkflow + "', " +
                    "TargetMilestoneId: '" + selectedMilestone + "', " +
                    "TargetDt: '" + selectedDate + "'}",
                dataType: "json",
                async: false,
                success: function (msg) {
                    $("#loading").toggle();
                    $("#body").toggle();
                    location.reload();
                },
                error: function (msg, status, error) {
                    var json = $.parseJSON(msg.responseText);
                    alert("Error Initializing Records\n" + json.Message);
                    initDialog.saveErrorThrown = true;
                    $("#AgencyInitWorkflowDialog").unblock();
                }
            });
        }
    });
};

InitDialog.Format = function () {
    var table = $('#AgencyInitWorkflowDialog table');
    $('caption', table).addClass('ui-state-default');
    $('th', table).addClass('ui-state-default');
    $('td', table).addClass('ui-widget-content');
    $(table).delegate('tr', 'hover', function () {
        $('td', $(this)).toggleClass('ui-state-highlight');
    });
    $(".datePicker", table).datepicker();
};

InitDialog.SetupNewRows = function (rows) {
    $(rows).find('td').addClass('ui-widget-content');
};

InitDialog.Load = function (DeliverableVersionId, AgencyTypeId) {
    InitDialog.DataBindTable(DeliverableVersionId, AgencyTypeId);
    InitDialog.Format();
}

InitDialog.PopulateDeliverableVersionList = function (DeliverableRootId) {
    $("#ddlDeliverableVersion > option").remove();

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi//AgencyCertification.asmx/SelectUninitializedVersions",
        data: "{ DeliverableRootId: '' }",
        dataType: "json",
        async: false,
        success: function (msg) {
            if (msg.d == "[]") {
                $("#AgencyInitWorkflowDialog").block({ message: "<h1>No Records to Initialize</h1>" });
                setTimeout($(this).dialog("close"), 3000);
            } else {
                var dt = $.parseJSON(msg.d);
                var myCombo = $("#ddlDeliverableVersion");
                for (var i = 0; i < dt.length; i++) {
                    var dr = dt[i];
                    var optionId = dr.DeliverableVersionId;
                    var optionText = dr.ProductVersionName + ' - ' + dr.Name + ' ' + dr.Version + ' - ' + dr.AgencyTypeName;
                    myCombo.get(0).options[i] = new Option(optionText, optionId);
                }
            }
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Retrieving Uninitialized Deliverables\n" + json.Message);
        }
    });
}



InitDialog.DataBindTable = function (DeliverableVersionId, AgencyTypeId) {
    var table = $('#AgencyInitWorkflowDialog table');

    $("tbody", table).empty();

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi//AgencyCertification.asmx/ListWorkflows",
        data: "{}",
        dataType: "json",
        async: false,
        success: function (msg) {
            initDialog.workflowJson = $.parseJSON(msg.d);
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Retrieving Agency Workflow List\n" + json.Message);
        }
    });

    var workflowJson = initDialog.workflowJson;

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/SelectUninitialized",
        data: "{ ProductVersionId: '', DeliverableRootId: '', DeliverableVersionId: '" + DeliverableVersionId + "' }",
        dataType: "json",
        async: false,
        success: function (msg) {
            //alert(msg.d);
            var dt = $.parseJSON(msg.d);
            for (var i = 0; i < dt.length; i++) {
                var dr = dt[i];
                $("tbody", table).append('<tr id="' + dr.ID + '"><td>' + dr.AgencyName + '</td>' +
                            '<td>' + InitDialog.GetSelectList(dr.ID, dr.WorkflowId, dr.AgencyTypeId, workflowJson) + '</td>' +
                            '<td>' + dr.MilestoneName + '</td>' +
                            '<td><input class="datePicker" id="dp' + dr.ID + '" type="text" /></td></tr>');
            }
        },
        error: function (msg, status, error) {
            try {
                var json = $.parseJSON(msg.responseText);
                alert("Error Retrieving Uninitialized Records\n" + json.Message);
            } catch (e) {
                alert("Error Retrieving Uninitialized Records\n" + msg.responseText);
            }
        }
        //,complete: function (msg, status) {
        //    alert(msg.responseText);
        //}
    });
}

InitDialog.GetSelectList = function (rowId, defaultWorkflowId, agencyTypeId, json) {
    var selectList = '<select id="selMilestone' + rowId + '">';
    var selected = '';
    if (json != null) {
        for (var i = 0; i < json.length; i++) {
            if (json[i].AgencyTypeId == agencyTypeId) {
                selected = '';
                if (json[i].WorkflowId == defaultWorkflowId) selected = 'selected';
                selectList += '<option value="' + json[i].ID + '">' + json[i].Name + '</option>';
            }
        }
    }
    selectList += '</select>';
    return selectList;
}

/*
 * Leverage Code
 */

var Leverage = function () {
    this.saveErrorThrown = false;
}
var leverage = new Leverage();

Leverage.Setup = function () {
    $("#leverageDialog").dialog({
        modal: true,
        autoOpen: false,
        width: 400,
        height: 500,
        show: 'clip',
        stack: true,
        buttons: {
            "Save": function () {
                Leverage.Save()
                if (!leverage.saveErrorThrown) {
                    $(this).dialog("close");
                } else {
                    alert("Error Saving Changes");
                }
            },
            "Cancel": function () {
                $(this).dialog("close");
            }
        }
    });
}

Leverage.Show = function (DeliverableVersionId, AgencyTypeId) {
    var table = $('#leverageDialog table');
    $('caption', table).addClass('ui-state-default');
    $('th', table).addClass('ui-state-default');
    $('td', table).addClass('ui-widget-content');
    $('tr td', table).first().width('25px');
    $(table).delegate('tr', 'hover', function () {
        $('td', $(this)).toggleClass('ui-state-highlight');
    });

    $("#leverageDialog #deliverableVersionId").val(DeliverableVersionId);
    $("#leverageDialog #agencyTypeId").val(AgencyTypeId);

    Leverage.LoadDeliverableTable();

    $("#leverageDialog").dialog("open");
    //LeverageVersion.Show();
}

Leverage.LoadDeliverableTable = function () {
    var table = $("#leverageDialog table");
    $("tbody tr", table).remove();

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "/ExcaliburApi/AgencyCertification.asmx/SelectUninitialisedDeliverablesByCategory",
        data: "{ DeliverableCategoryId: '" + $("#deliverableCategoryId").val() + "' }",
        dataType: "json",
        async: false,
        success: function (msg) {
            //alert(msg.d);
            var dt = $.parseJSON(msg.d);
            for (var i = 0; i < dt.length; i++) {
                var dr = dt[i];
                $("tbody", table).append('<tr id="' + dr.DeliverableVersionId + '">' +
                    '<td><input id="chkDeliverable" value="' + dr.DeliverableVersionId + '" type="checkbox"></td>' +
                    '<td>' + dr.Name + ' ' + dr.DeliverableVersion + '</td>' +
                    '<td>' + dr.DotsName + '</td></tr>');
            }
        },
        error: function (msg, status, error) {
            var json = $.parseJSON(msg.responseText);
            alert("Error Retrieving Uninitialized Deliverables\n" + json.Message);
        }
    });

}

Leverage.Save = function () {
    var table = $("#leverageDialog table");
    $("input:checked", table).each(function () {
        //alert($(this).val());
        //alert($("#deliverableRootId").val());
        //alert($("#agencyTypeId").val());

        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/ExcaliburApi/AgencyCertification.asmx/LeverageAgencyStatusWorkflow",
            data: "{ SourceDeliverableId: '" + $("#leverageDialog #deliverableVersionId").val() + "', " +
                "TargetDeliverableId: '" + $(this).val() + "', " +
                "AgencyTypeId: '" + $("#leverageDialog #agencyTypeId").val() + "' }",
            dataType: "json",
            async: false,
            success: function(msg){},
            error: function (msg, status, error) {
                //var json = $.parseJSON(msg.responseText);
                alert("Error Leveraging Deliverables\n" + msg.responseText);
                leverage.saveErrorThrown = true;
            }
        });

    });
}

/*
* Leverage Code
*/

var LeverageVersion = function () {
    this.saveErrorThrown = false;
}
var leverageVersion = new Leverage();

LeverageVersion.Setup = function () {
    $("#leverageVersionDialog").dialog({
        modal: true,
        autoOpen: false,
        width: 400,
        height: 500,
        show: 'clip',
        stack: true,
        buttons: {
            "Save": function () {
                Leverage.Save()
                if (!leverageVersion.saveErrorThrown) {
                    $(this).dialog("close");
                } else {
                    alert("Error Saving Changes");
                }
            },
            "Cancel": function () {
                $(this).dialog("close");
            }
        }
    });
}

LeverageVersion.Show = function () {
    var table = $('#leverageVersionDialog table');
    $('caption', table).addClass('ui-state-default');
    $('th', table).addClass('ui-state-default');
    $('td', table).addClass('ui-widget-content');
    $('tr td', table).first().width('25px');
    $(table).delegate('tr', 'hover', function () {
        $('td', $(this)).toggleClass('ui-state-highlight');
    });

    //$("#leverageVersionDialog #sourceVersionId").val(DeliverableVersionId);
    //$("#leverageVersionDialog #targetVersionId").val(AgencyTypeId);

    //LeverageVersion.LoadDeliverableTable();

    $("#leverageVersionDialog").dialog("open");
}

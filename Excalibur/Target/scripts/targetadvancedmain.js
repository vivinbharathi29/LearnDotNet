//******************************************************************
//File Description:     USER INTERFACE MANIPULATION
//Details:              Functions that manipulate UI 
//Created:              01/22/2016 Harris, Valerie - PBI 15594 / Task 15982
//Modified:             01/28/2016 - Harris, Valerie - add page wait functionality; apply multi target to active versions; disabled inactive versions' fields
//******************************************************************
var $browser = null;
var $browserversion = null;
var $totalrows = null;

//*****************************************************************
//Description:  OnLoad, on page load instantiate functions
//*****************************************************************
$(window).load(function () {
    //*********APPLY JS FUNCTIONS PAGE ELEMENTS*******//
    //Browser Detection
    BrowserDetection();

    //Instantiate Elements
    LoadUIElement(true);

});


//*****************************************************************
//Description:  Load Jquery UI Elements
//Function:     LoadUIElement();
//*****************************************************************
function LoadUIElement(bPageLoad) {
    if (bPageLoad === true) {
        //Get Browser Name and Version: ---
        $browser = $("html").attr("data-browser");
        $browserversion = $("html").attr("data-version");
        $totalrows = (parseFloat($('#inpCount').val()) + 1);

        // If Global, Hide the multi selection column: 
        HideMultiSelection('Target');

        //Add click event to checkbox all: ---
        $('#chkSelectAll').click(function (e) {
            var bChecked = $('#chkSelectAll').is(":checked");
            var sStatus = null;

            if (bChecked === true) {
                sStatus = "Selecting";
            } else {
                sStatus = "Un-selecting"
            }

            PageWait("" + sStatus + " " + $totalrows + " rows...please wait.", "load", true);
            window.setTimeout(function () {
                var bChecked = $('#chkSelectAll').is(":checked");
                SelectMultipleVersions('Target', bChecked);
            }, 500);
        });

        //Add click event to Apply button: ---
        $('#btnApply').click(function (e) {
            PageWait("Applying target to " + $totalrows + " rows...please wait.", "load", true);
            window.setTimeout(function () {
                ApplyTargetVersion();

                //Save last applied multi select option: ---
                if ($('#selMultiTarget').val() != "") {
                    $('#inpMultiTarget').val($('#selMultiTarget').val())
                    $('#inpMultiUsed').val("True");
                }
            }, 500);
        });

        //Add click event to Undo button: ---
        $('#btnUndo').click(function (e) {
            PageWait("Changing target back to original value for " + $totalrows + " rows...please wait.", "load", true);
            window.setTimeout(function () {
                UndoTargetVersion();
            }, 500);
        });

        //Add click event to all checkboxes: ---
        $('table#tblTarget input.chkbox-option[type="checkbox"][data-status="active"]').click(function (e) {
            var bChecked = $(this).is(":checked");
            var iRowCount = $(this).attr("data-checkbox");
            SelectSingleVersion('Target', iRowCount, bChecked);
        });

        //Add change event to all drop-downs in target version table that are active: ---
        $('table#tblTarget .select-option[data-status="active"]').change(function (e) {
            var iRowCount = $(this).attr("data-select");
            var sPreviousValue = $('#select_' + iRowCount + '').val();
            var sSelectedValue = $('table#tblTarget .select-option[data-select="' + iRowCount + '"] option:selected').val()
            var sMultiTarget = $('#inpMultiTarget').val();
            var bMultiUsed = $('#inpMultiUsed').val();

            //If disabled, prevent user from changing target: ---
            if ($(this).hasClass("disabled") === true) {
                $('table#tblTarget .select-option[data-select="' + iRowCount + '"] option:selected').prop("selected", false);
                if (bMultiUsed === 'True' && sMultiTarget != "") {
                    //In case User selects empty, use the last multi target applied: ---
                    $('table#tblTarget .select-option[data-select="' + iRowCount + '"]').val(sMultiTarget);
                } else {
                    $('table#tblTarget .select-option[data-select="' + iRowCount + '"]').val(sPreviousValue);
                }
            } else {
                //If not disabled, save the change value: ---
                $('#select_' + iRowCount + '').val(sSelectedValue);
            }
        });

        //Make inactive drop-downs and input fields disabled, remove unselected options: --
        //Prevent User from changing targets for inactive versions: --
        $('table#tblTarget [data-status="inactive"]').attr("readonly", true);
        var $select = $('table#tblTarget .select-option[data-status="inactive"]');
        $.each($select, function (index) {
            if ($(this).val() === 'Available') {
                $(this).addClass("disabled");
                $("option:not(:selected)", this).remove();
            }
        });

        $('table#tblTarget .select-option[data-status="inactive"]').change(function (e) {
            $('table#tblTarget .select-option[data-status="inactive"] option:contains("Targeted")').attr('disabled', true);
            var iRowCount = $(this).attr("data-select");
            var sSelectedValue = $('table#tblTarget .select-option[data-select="' + iRowCount + '"] option:selected').val()
            $('#select_' + iRowCount + '').val(sSelectedValue);

        });

        var $selectDisabled = $('.select-option.disabled');
        $.each($selectDisabled, function (index) {
            $("option:not(:selected)", this).remove();
        });

    }
}


//*****************************************************************
//Description:  Hide/Show Select All Column
//Function:     DisableRelatives();
//*****************************************************************
function HideMultiSelection(sTableID) {
    var bGlobal = $('#inpGlobal').val();
    var iIndex = 0;

    //If, Root Target Delieverable is Global; hide multi select column: ---
    if (bGlobal == 'True') {
        //$('tr td:nth-child(1)').hide();
        $('table#tbl' + sTableID + ' tr').each(function () { $(this).children(":eq(" + iIndex + ")").addClass("hide"); });
    }
}

//*****************************************************************
//Description:  OnClick, select/deselect all checkboxes
//Function:     MultipleSelectCheckbox();
//*****************************************************************
function SelectMultipleVersions(sTableID, bChecked) {
    var oCheckBox = null;
    var oSelect = null;
    var oText = null;
    var oSection = null;
    var bMultiUsed = null;

    oCheckBox = $('table#tbl' + sTableID + ' input.chkbox-option[type="checkbox"]');
    oSelect = $('table#tbl' + sTableID + ' .select-option[data-status="active"]');
    oText = $('table#tbl' + sTableID + ' input.text-option[type="text"][data-status="active"]');
    oSection = $('#trApplyTargetSection');
    bMultiUsed = $('#inpMultiUsed').val();

    if (bChecked === true) {
        //Check All Checkboxes : ---
        oCheckBox.attr("checked", true);
        oCheckBox.prop('checked', true);

        //Make all drop-downs readonly: --- 
        oSelect.attr("readonly", true);
        oSelect.addClass("disabled");

        //Make all text fields readonly: --- 
        oText.attr("readonly", true);
        oText.addClass("disabled");

        //If checkboxes are selected and Apply Target Section is hidden, show Apply Target Section: ---
        if (isCheckbox() == true && oSection.hasClass("hide") == true) {
            oSection.removeClass("hide").addClass(ShowRowClass());
        }

        //Hide Page Load
        PageWait("", "load", false);
    } else {
        //Deselect All Checkboxes: --- 
        oCheckBox.attr("checked", false);
        oCheckBox.prop('checked', false);

        //Make all drop-downs enabled: --- 
        oSelect.attr("readonly", false);
        oSelect.removeClass("disabled");

        //Make all text fields enabled: --- 
        oText.attr("readonly", false);
        oText.removeClass("disabled");

        //If no checkboxes are selected and Apply Target Section isn't hidden, hide Apply Target Section: ---
        if (isCheckbox() == false && oSection.hasClass("hide") == false) {
            oSection.removeClass(ShowRowClass()).addClass("hide");

            //Reset multiple select fields: ---
            $('#selMultiTarget').val("");
            $('#txtMultiNotes').val("");
            $('#inpMultiTarget').val("");
            $('#inpMultiUsed').val("False");
        }


        if (bMultiUsed === 'True') {
            PageWait("Changing target back to original value " + $totalrows + " rows...please wait.", "load", true);
            window.setTimeout(function () {
                UndoTargetVersion();
            }, 1000);
        } else {
            //Done, Hide Page Load
            PageWait("", "load", false);
        }
    }
}

//*****************************************************************
//Description:  OnClick, select/deselect single checkbox
//Function:     SelectSingleVersion();
//*****************************************************************
function SelectSingleVersion(sTableID, iRowCount, bChecked) {
    var oSelect = null;
    var oText = null;
    var bIsSelected = isCheckbox();
    var bIsNotSelected = isNotCheckbox();

    oSelect = $('table#tbl' + sTableID + ' .select-option[data-select="' + iRowCount + '"]');
    oText = $('table#tbl' + sTableID + ' input[type="text"][data-text="' + iRowCount + '"]');
    sDBTarget = $('table#tbl' + sTableID + ' input[type="hidden"][data-dbtarget="' + iRowCount + '"]').val();
    sDBNote = $('table#tbl' + sTableID + ' input[type="hidden"][data-dbnote="' + iRowCount + '"]').val();
    oSection = $('#trApplyTargetSection');

    if (bChecked === true) {
        //Make all drop-downs readonly: --- 
        oSelect.attr("readonly", true);
        oSelect.addClass("disabled");


        //Make all text fields readonly: --- 
        oText.attr("readonly", true);
        oText.addClass("disabled");

        //If checkboxes are selected and Apply Target Section is hidden, show Apply Target Section: ---
        if (bIsSelected == true && oSection.hasClass("hide") == true) {
            oSection.removeClass("hide").addClass(ShowRowClass());
        }
    } else {
        //Make all drop-downs enabled: --- 
        oSelect.attr("readonly", false);
        oSelect.removeClass("disabled");

        //Change select to previous selected option from database: ---
        if (sDBTarget != "") {
            $('table#tbl' + sTableID + ' .select-option[data-select="' + iRowCount + '"] option:selected').prop("selected", false);
            $('table#tbl' + sTableID + ' .select-option[data-select="' + iRowCount + '"]').val(sDBTarget);
        }

        //Make all text fields enabled: --- 
        oText.attr("readonly", false);
        oText.removeClass("disabled");

        //If database value isn't empty, change select to original, saved default value : ---
        oText.val(sDBNote);

        //If no checkboxes are selected and Apply Target Section isn't hidden, hide Apply Target Section: ---
        if (bIsSelected == false && oSection.hasClass("hide") == false) {
            oSection.removeClass(ShowRowClass()).addClass("hide");

            //Reset multiple select fields: ---
            $('#selMultiTarget').val("");
            $('#txtMultiNotes').val("");
            $('#inpMultiTarget').val("");
            $('#inpMultiUsed').val("False");
        }
    }

    //Reset check all field when deselected and selected checkboxes exist: ---
    if (bIsSelected == true && bIsNotSelected == true) {
        //Reset select all checkbox
        $('#chkSelectAll').attr("checked", false);
        $('#chkSelectAll').prop('checked', false);
    } else if ((bIsSelected == true && bIsNotSelected == false)) {
        //Reset select all checkbox
        $('#chkSelectAll').attr("checked", true);
        $('#chkSelectAll').prop('checked', true);
    }
}


//*****************************************************************
//Description:  OnClick, apply target to multiple drop-downs and text input fields
//Function:     ApplyTargetVersion();
//*****************************************************************
function ApplyTargetVersion() {
    var sTableID = "Target";
    var sMultiTarget = $('#selMultiTarget').val();
    var sMultiNotes = $('#txtMultiNotes').val();


    //If Multi Target empty, show alert message else apply target and target notes to multiple records
    if (sMultiTarget === "") {
        alert("Please select a Target.");
    } else {
        //For each selected checkbox, change drop-down and notes to what's been entered in the Apply Targeted Version Section
        $('table#tbl' + sTableID + ' input.chkbox-option[type="checkbox"][data-status="active"]:checked').each(function (index, row) {
            //Get the row count value: ---
            //iRowCount = index //$(this).attr("data-checkbox");
            iRowCount = $(this).attr("data-checkbox");

            //If not empty, apply multi target selection to all drop-downs in the target version table: ---
            $('table#tbl' + sTableID + ' .select-option[data-select="' + iRowCount + '"]').val(sMultiTarget);
            //update the status dropdown with the multiTarget value
            $('#select_' + iRowCount + '').val(sMultiTarget);

            //Apply multi target note to all text fields in the target version table: ---
            $('table#tbl' + sTableID + ' input.text-option[type="text"][data-text="' + iRowCount + '"]').val(sMultiNotes);

        });
    }

    //Done, Hide Page Load
    PageWait("", "load", false);
}

//*****************************************************************
//Description:  OnClick, refresh the page to revert changes backed to what's currently saved in the database
//Function:     UndoTargetVersion();
//*****************************************************************
function UndoTargetVersion() {
    var sTableID = "Target";
    var iRowCount = null;
    var sDBTarget = null;
    var sDBNote = null;
    var sPreviousValue = null;
    var sMultiTarget;

    $('table#tbl' + sTableID + ' input.chkbox-option[type="checkbox"][data-status="active"]:checked').each(function (index, row) {
        //Get the row count value: ---
        //iRowCount = index //$(this).attr("data-checkbox");
        iRowCount = $(this).attr("data-checkbox");
        //Get the target and target notes that are saved in the database for this record: ---
        sDBTarget = $('input[type="hidden"][data-dbtarget="' + iRowCount + '"]').val();
        sDBNote = $('input[type="hidden"][data-dbnote="' + iRowCount + '"]').val();
        sPreviousValue = $('#select_' + iRowCount + '').val();

        //Apply target's database selection to drop-down: ---
        if (sDBTarget != "") {
            $('.select-option[data-status="active"][data-select="' + iRowCount + '"]').val(sDBTarget);
            sMultiTarget = sDBTarget;
        } else {
            //If DB value is empty, used the previously selected value...if any: ---
            $('.select-option[data-status="active"][data-select="' + iRowCount + '"]').val(sPreviousValue);
            sMultiTarget = sPreviousValue;
        }

        //apply multi target selection to all drop-downs in the target version table: ---
        $('table#tbl' + sTableID + ' .select-option[data-select="' + iRowCount + '"]').val(sMultiTarget);
        //update the status dropdown with the multiTarget value
        $('#select_' + iRowCount + '').val(sMultiTarget);

        //Apply target note to text field: ---
        $('input[type="text"][data-status="active"][data-text="' + iRowCount + '"]').val(sDBNote);
    });

    //Reset multiple select fields: ---
    $('#selMultiTarget').val("");
    $('#txtMultiNotes').val("");
    $('#inpMultiTarget').val("");
    $('#inpMultiUsed').val("False");

    //Done, Hide Page Load
    PageWait("", "load", false);
}


//*****************************************************************
//Description:  Determine if one or more of select all checkboxes are checked
//Function:     isCheckbox();
//*****************************************************************
function isCheckbox() {
    var sTableID = "Target";
    var bStatus = false;

    if ($('table#tbl' + sTableID + ' input.chkbox-option[type="checkbox"][data-status="active"]:checked').length > 0) {
        bStatus = true;
    }
    return bStatus;
}

//*****************************************************************
//Description:  Determine if one or more of select all checkboxes not checked
//Function:     isNotCheckbox();
//*****************************************************************
function isNotCheckbox() {
    var sTableID = "Target";
    var bStatus = false;

    if ($('table#tbl' + sTableID + ' input.chkbox-option[type="checkbox"][data-status="active"]:not(:checked)').length > 0) {
        bStatus = true;
    }
    return bStatus;
}


//*****************************************************************
//Description:  Based on browser, use class name to show table row
//Function:     ShowRowClass();
//*****************************************************************
function ShowRowClass() {
    var sClassShowRowName = null;

    if ($browser === 'Chrome' || $browser === 'Opera' || $browser === 'Firefox' || $browser === 'Netscape' || $browser == 'MSIE' && $browserversion > 10) {
        sClassShowRowName = 'show-row';
    } else {
        sClassShowRowName = 'show';
    }

    return sClassShowRowName;
}

//*****************************************************************
//Description:  Detect Browser Type
//Function:     BrowserDetection();
//*****************************************************************
function BrowserDetection() {
    //Browser Detection
    $("html").attr("data-browser", get_browser());
    $("html").attr("data-version", parseInt(get_browser_version()));

    function get_browser() {
        var N = navigator.appName, ua = navigator.userAgent, tem;
        var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
        if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
        M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
        return M[0];
    }
    function get_browser_version() {
        var N = navigator.appName, ua = navigator.userAgent, tem;
        var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
        if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
        M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
        return M[1];
    }
}

//*****************************************************************
//Description:  Show Loading Icon and Text for long running scripts
//Function:     PageWait();
//*****************************************************************
function PageWait(sMsg, sType, bShow) {
    switch (sType) {
        case "load":
            if (bShow == true) {
                //Hide Page
                $("#page").removeClass("show").addClass("hide");

                //Change cursor to default 
                $("body").css("cursor", "wait");

                //Update Page Load Message
                $("#msg-page-load").text(sMsg);

                //show page loading
                $("#loading-dialog").css('display', 'inline');
            } else {
                //Change cursor to default 
                $("body").css("cursor", "default");

                //Update Page Load Message
                $("#msg-page-load").text("");

                //hide page loading
                $("#loading-dialog").css('display', 'none');

                //Show Page
                $("div#page").removeClass("hide").addClass("show");
            }
            break;
        default:
            break;
    }
}
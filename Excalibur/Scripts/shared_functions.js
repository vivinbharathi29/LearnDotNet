//******************************************************************
//File Description:     USER INTERFACE MANIPULATION
//Details:              Functions that manipulate UI shared on multiple pages
//Created:              Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
//******************************************************************

//*****************************************************************
//Description:  Return the window size dimension based on desktop resolution
//Function:     GetWindowSize();
//Created:      Harris, Valerie (1/21/2016) - PBI 15594/Task 15982
//*****************************************************************
function GetWindowSize(sType) {
    var iScreenWidth = window.screen.width;
    var iScreenHeight = window.screen.height;
    var iBrowserWidth = $(window).width();
    var iBrowserHeight = $(window).height();
    var iReturnValue = null;

    switch (sType) {
        case "width":
            //Get Width of Window based on User screen Resolution
            //if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
            //    iReturnValue = (iScreenWidth - 550);
            //} else if (iScreenWidth >= 1600) {
            //    iReturnValue = (iScreenWidth - 750);
            //} else if (iScreenWidth <= 1024) {
            //    iReturnValue = (iScreenWidth - 100);
            //}
            iReturnValue = iBrowserWidth * 85 / 100;
            break;
        case "height":
            //Get Height of Window based on User screen Resolution
            //if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
            //    iReturnValue = (iScreenHeight - 450);
            //} else if (iScreenWidth >= 1600) {
            //    iReturnValue = (iScreenHeight - 650);
            //} else if (iScreenWidth <= 1024) {
            //    iReturnValue = (600);
            //}
            iReturnValue = iBrowserHeight * 90 / 100;
            break;
        case "left":
            iReturnValue = (screen.width / 2) - (iBrowserWidth / 2);
            break;
        case "top":
            iReturnValue = (screen.height / 2) - (screen.height / 2);
            break;
    }

    return iReturnValue;
}

//*****************************************************************
//Description:  Refresh the Page
//Function:     ReloadWindow();
//Modified:     Harris, Valerie (1/27/2016) - PBI 15595/Task 15983
//*****************************************************************
function ReloadWindow() {
    window.location.reload(true);
}

//*****************************************************************
//Description:  Return Current User's Permission Status
//Function:     GetUserPermission();
//Modified:     Harris, Valerie (5/10/2016) - PBI 12538/Task 20310
//*****************************************************************
function GetUserPermission(sPermission) {
    var _userPermission = null;
    var _PermissionStatus = 0;

    $.ajax({
        type: "Post",
        url: "/iPulsar/Admin/System Admin/UsersAndRoles_Main.aspx/GetUserPermission",
        contentType: "application/json; charset=utf-8",
        global: false,
        async: false,
        dataType: "json",
        success: function (data) {
            _userPermission = data;
        },
        error: function (xhr, ajaxOptions, thrownError) {
            alert(xhr.statusText);
        },
        cache: false
    });
    
    if(isEmpty(_userPermission) === false){
        for (var i = 0; i < _userPermission.d.length; i++) {
            if (_userPermission.d[i] == sPermission) {
                _PermissionStatus = 1;
                break;
            }
        }
    }

    return _PermissionStatus;
}

//*****************************************************************
//Description:  Check if Object is empty
//Function:     isEmpty();
//Modified:     Harris, Valerie (5/10/2016) - PBI 12538/Task 20310
//*****************************************************************
function isEmpty(obj) {
   for (var x in obj) { return false; }
   return true;
}

//************************************************************************
//Description:  Disable form fields and hide submit buttons 
//Function:     disableFormElements();
//Modified:     Harris, Valerie (5/10/2016) - PBI 12538/Task 20310
//              Harris, Valerie (10/20/2016) - PBI 27863/Task 24551
//************************************************************************
function disableFormElements(sFormName){
    if ($('form[name="' + sFormName + '"], form[id="' + sFormName + '"]').length > 0) {

        if ($('form[name="' + sFormName + '"] input:text, form[id="' + sFormName + '"] input:text').length > 0) {
            $('form[name="' + sFormName + '"] input:text, form[id="' + sFormName + '"] input:text').addClass("disabled").prop("disabled", true);
        }

        if ($('form[name="' + sFormName + '"] input:checkbox, form[id="' + sFormName + '"] input:checkbox').length > 0) {
            $('form[name="' + sFormName + '"] input:checkbox, form[id="' + sFormName + '"] input:checkbox').addClass("disabled").prop("disabled", true);
        }
    
        if ($('form[name="' + sFormName + '"] input:radio, form[id="' + sFormName + '"] input:radio').length > 0) {
            $('form[name="' + sFormName + '"] input:radio, form[id="' + sFormName + '"] input:radio').addClass("disabled").prop("disabled", true);
        }

        if ($('form[name="' + sFormName + '"] select, form[id="' + sFormName + '"] select').length > 0) {
            $('form[name="' + sFormName + '"] select, form[id="' + sFormName + '"] select').addClass("disabled").prop("disabled", true);
            if ($('#txtID').val() > 0){ //Remove all un-selected options
                $('form[name="' + sFormName + '"] select option:not(:selected), form[id="' + sFormName + '"] select option:not(:selected)').remove();
            }
        }

        if ($('form[name="' + sFormName + '"] textarea, form[id="' + sFormName + '"] textarea').length > 0) {
            $('form[name="' + sFormName + '"] textarea, form[id="' + sFormName + '"] textarea').addClass("disabled").prop("disabled", true);
        }

        if ($('form[name="' + sFormName + '"] input:file, form[id="' + sFormName + '"] input:file').length > 0) {
            $('form[name="' + sFormName + '"] input:file, form[id="' + sFormName + '"] input:file').addClass("disabled").prop("disabled", true);
        }

        if ($('form[name="' + sFormName + '"] input:image, form[id="' + sFormName + '"] input:image').length > 0) {
            $('form[name="' + sFormName + '"] input:image, form[id="' + sFormName + '"] input:image').addClass("disabled").prop("disabled", true);
        }

        if ($('form[name="' + sFormName + '"] input:button, form[id="' + sFormName + '"] input:button').length > 0) {
            $('form[name="' + sFormName + '"] input:button, form[id="' + sFormName + '"] input:button').addClass("disabled").prop("disabled", true);
        }

        //Disable form submission buttons in either LowerWindow frame (if it exists) or main page: ---
        if (window.parent.frames["LowerWindow"]) {
            if (window.parent.frames["LowerWindow"].cmdSubmit) {
                window.parent.frames["LowerWindow"].cmdSubmit.disabled = true;
            }
            if (window.parent.frames["LowerWindow"].cmdOK) {
                window.parent.frames["LowerWindow"].cmdOK.disabled = true;
            }
            if (window.parent.frames["LowerWindow"].cmdEditCancel) {
                window.parent.frames["LowerWindow"].cmdEditCancel.disabled = true;
            }
            if (window.parent.frames["LowerWindow"].cmdClear) {
                window.parent.frames["LowerWindow"].cmdClear.disabled = true;
            }
        } else {
            $('#cmdSubmit').prop("disabled", true);
            $('#cmdOK').prop("disabled", true);
            $('#cmdEditCancel').prop("disabled", true);
            $('#cmdClear').prop("disabled", true);
        }
    }
}

//************************************************************************
//Description:  Disable form fields and hide submit buttons;
//Function:     showPermissionMessage();
//Modified:     Harris, Valerie (5/11/2016) - PBI 12538/Task 20310
//Note:         Add shared.css to the page
//************************************************************************
function showPermissionMessage(bNoPermission, sFeatureName) {
    if (bNoPermission === true) {
        if ($('#spnErrorMessage').length > 0) {    //If errorMessage span tag on the page: ----
            if ($('#txtID').length > 0) {       //If ID field exists
                if ($('#txtID').val() > 0) {    //If the ID value is greater than 0, show msg for Edit mode else Add mode
                    $('#spnErrorMessage').text('You do not have permission to edit ' + sFeatureName + 's .');
                } else {
                    $('#spnErrorMessage').text('You do not have permission to create a new ' + sFeatureName + '.');
                }
            } else {//If ID field doesn't exist, show generic permission msg
                $('#spnErrorMessage').text('You do not have permission to access ' + sFeatureName + 's .');
            }

            $('#spnErrorMessage').removeClass("hide").addClass("show"); //make error message visable
        }
    }
}


//************************************************************************
//Description:  On page load, validate User's page permission
//Function:     ValidatePagePermission();
//Modified:     Harris, Valerie (5/11/2016) - PBI 12538/Task 20310
//              Harris, Valerie (10/20/2016) - PBI 27863/Task 24551
//************************************************************************
function ValidatePagePermission(sPageName, sFeatureName) {
    var _isSystemAdmin = null;
    var _isProductEditor = null;
    var _isScheduleEditor = null;
    var iPermissionStatus = null;
    var bIsPulsarProduct = '';
    var bNoPermission = false;

    switch(sPageName) {
        case "ProgramMain":
                _isSystemAdmin = GetUserPermission('System.Admin');
                _isProductEditor = GetUserPermission('Product.Edit');
                
                if (_isSystemAdmin !== 1 && _isProductEditor !== 1) {
                    bNoPermission = true;

                    //Disable form fields and buttons: ----
                    disableFormElements("ProgramInput");

                    //Hide additional elements specific to ProgramInput form: ---
                    $('#ReleaseLink').addClass("hide");
                    $('#SystemboardLink').addClass("hide");
                    $('#MachinePNPLink').addClass("hide");
                    $('#CycleLink').addClass("hide");
                    $('#SiteLink').addClass("hide");

                    //Hide all span tags with "Add" in text; these have onClick events : ---
                    $("span:contains('Add')").addClass("hide");

                    //Show Permission Error Message: ---
                    showPermissionMessage(bNoPermission, sFeatureName);
                }
            break;
        case "PMView":
                bIsPulsarProduct = globalVariable.get('product_type');

                _isSystemAdmin = GetUserPermission('System.Admin');
                _isProductEditor = GetUserPermission('Product.Edit');
                _isProductDelete = GetUserPermission('Schedule.Delete');

                if (GetUserPermission('Schedule.Edit') == 1 || $("#inpSEPMProductUser").val() == "1") {
                    _isScheduleEditor = 1;
                } else {
                    _isScheduleEditor = 0;
                }

                //--ALL TABS: ------
                //If Pulsar product and doesn't have Permission System.Admin and/or Product.Editor; 
                //Hide Edit and Clone Product links: ---
                if (bIsPulsarProduct === 'True') {
                    if (_isSystemAdmin !== 1 && _isProductEditor !== 1) {
                        $('td#EditLink').addClass("hide");
                        $('td#CloneLink').addClass("hide");
                    }
                }

                //--SCHEDULE TAB: ------
                //If doesn't have System.Admin, Hide Delete Current Schedule link: ---
                if (_isSystemAdmin !== 1 && _isProductDelete !== 1) {
                    $('div#AddScheduleLink span.sysadmin-scheduletab').addClass("hide");
                }

                //If doesn't have System.Admin and/or Schedule.Editor permission, 
                //Hide All Add/Edit Schedule links: ---
                if (_isSystemAdmin !== 1 && _isScheduleEditor !== 1) {
                    $('div#AddScheduleLink span.admin-scheduletab').addClass("hide");
                }
                break;
        default:
            break;
    }
}

//************************************************************************
//Description:  Return User's Permission
//Function:     UserHasPermission();
//Modified:     Harris, Valerie (10/21/2016) - PBI 27863/Task 24551
//************************************************************************
function UserHasPermission(sPermissionName) {
    var iPermissionStatus = 0;

    switch(sPermissionName) {
        case 'Schedule.Edit':
            iPermissionStatus = GetUserPermission(sPermissionName);
            if (iPermissionStatus == 1 || $("#inpSEPMProductUser").val() == "1") {
                iPermissionStatus = 1;
            } else {
                iPermissionStatus = 0;
            }
            break;
        default:
            iPermissionStatus = GetUserPermission(sPermissionName);
            break;
    }

    if (iPermissionStatus == 1) {
        return true;
    } else {
        return false;
    }
}

//************************************************************************
//Description:  Replace IE showModalDialog with compatible JQuery Dialog
//Function:     showNewModalDialog();
//Modified:     Harris, Valerie (9/6/2016) - PBI 23434/Task 24367 - Browser Compatability: Change webpage dialogs to jquery dialogs
//              Harris, Valerie (10/6/2016) - Bug 27738/Task 27767 - //10/6 - change window.reload to window.location.href so it works for all reload scenerios
//              Harris, Valerie (10/13/2016 - PBI 27769 /Task 27838 - 10/13 - add code to position dialog relative to where link is clicked on pages with alot of data 
//Note:         jquery.ui.min, juery, and shared.css files required
//************************************************************************
var modalDialog = {
    load: function () {
        //Appends the parent page with modal dialog div and iframe tags; 
        //usually the parent page and is the Main or root level page
        if ($("#modal_dialog").length == 0) {
            //append modal div to body tag
            $('<div id="modal_dialog" title="" class="hide"><iframe id="modal_iframe" width="100%" height="100%" frameborder="0" src="" class="hide"></iframe></div>').appendTo('body');
        }

        //make sure instance of dialog is closed; for pages that have multiple pop-ups
        $('#modal_dialog').bind('dialogclose', function (event) {
            $("#modal_dialog").dialog('destroy');
        });

        //Loop thru each custom content dialogs 
        if ($("div.content-dialog").length != 0) {
            $("div.content-dialog").each(function () {
                //if child doesn't already exist, wrap content within modal dialog div with modal_content div
                if ($(this).children('#modal_content').length == 0) {
                    var sHTML = $('#' + $(this).attr("id") + '').html(); //dialog's current HTML
                    $('#' + $(this).attr("id") + '').html(""); //clear dialogs HTML
                    $('<div id="modal_content">' + sHTML + '</div>').appendTo('#' + $(this).attr("id") + ''); //wrap it in modal_content div
                }
            });
        }

        //Add hidden input field for default 
        if ($("#modal-returnValue").length == 0) {
            $('<input type="hidden" id="modal-returnValue" value=""/>').appendTo('form');
        }
    },
    open: function (args) {
        //Replaces showModalDialog code that opens the showModalDialog window; 
        //usually the parent page and is the Main or root level page

        //Set default values for empty parameters
        var param = $.extend({
            'dialogTitle': null,
            'dialogURL': null,
            'dialogDivID': null,
            'dialogHeight': 0,
            'dialogWidth': 0,
            'dialogArguments': null,
            'dialogArgumentsName': null,
            'dialogResizable': null,
            'dialogDraggable': null,
            'dialogEvent': null
        }, args);

        //create dialog
        if (param.dialogURL != null) { //open URL within an iframe
            $("#modal_dialog").dialog({
                modal: true,
                autoOpen: false,
                open: function (ev, ui) {
                    //Set dialog's DIV and Iframe values: ---
                    $('#modal_iframe').attr('src', '' + param.dialogURL + '');
                    $('#modal_iframe').attr('width', '' + (param.dialogWidth * 98/100) + '');
                    $('#modal_iframe').attr('height', '' + (param.dialogHeight * 94/100) + '');
                    $("#modal_dialog").removeClass('hide').addClass('show');
                    $('#modal_iframe').removeClass('hide').addClass('show');

                    if (param.dialogEvent == null) {
                        $("#modal_dialog").dialog('option', 'position', 'center');
                    }
                    
                    //pass value 
                    if (param.dialogArguments != null && param.dialogArguments != '') {
                        modalDialog.passArgument(param.dialogArguments, param.dialogArgumentsName);
                    }
                },
                close: function () {
                    //Reset dialog's DIV and Iframe values: ---
                    $("#modal_dialog").removeClass('show').addClass('hide');
                    $('#modal_iframe').removeClass('show').addClass('hide');

                    $("#modal_dialog").attr('title', '');
                    $('#modal_iframe').attr("src", '');
                    $('#modal_iframe').attr('width', '');
                    $('#modal_iframe').attr('height', '');
                    
                    //return values
                    modalDialog.returnValue();
                },
                height: param.dialogHeight + 10, //add some height to dialog so iframe doesn't appear cut-off
                width: param.dialogWidth + 10,  //add some width to dialog so iframe doesn't appear cut-off
                resizable: param.dialogResizable,
                draggable: param.dialogDraggable,
                closeOnEscape: false
            });

            //close context menu if it exists and it's open
            if (typeof oPopup != "undefined") {
                if (oPopup.isOpen) {
                    oPopup.hide();
                }
            } else if (typeof newPopup != "undefined") {
                if (newPopup.isOpen) {
                    newPopup.hide();
                }
            }

            //open dialog
            $("#modal_dialog").dialog("option", "title", param.dialogTitle);

            if (param.dialogEvent != null) {
                var oEvent = param.dialogEvent;
                $("#modal_dialog").dialog("option", "position", [(oEvent.clientX), (oEvent.clientY)]).dialog("open");
            } else {
                $("#modal_dialog").dialog({ position: { my: "center", at: "top+15%", of: window } }).dialog("open");
            }
        } else { //open custom modal dialog DIV element
            if (param.dialogDivID != null) {
                $('#' + param.dialogDivID + '').dialog({
                    modal: true,
                    autoOpen: false,
                    open: function (ev, ui) {
                        //Set dialog's DIV and container DIV values: ---
                        $('#modal_content').attr('width', '' + (param.dialogWidth - 10) + '');
                        $('#modal_content').attr('height', '' + (param.dialogHeight - 10) + '');

                        $('#' + param.dialogDivID + '').removeClass('hide').addClass('show');
                        $('#modal_content').removeClass('hide').addClass('show');

                        if (param.dialogEvent == null) {
                            $('#' + param.dialogDivID + '').dialog('option', 'position', 'center');
                        }

                        //pass value 
                        if (param.dialogArguments != null && param.dialogArguments != '') {
                            modalDialog.passArgument(param.dialogArguments, param.dialogArgumentsName);
                        }
                    },
                    close: function () {
                        //Reset dialog's DIV and container DIV values: ---
                        $('#' + param.dialogDivID + '').removeClass('show').addClass('hide');
                        $('#modal_content').removeClass('show').addClass('hide');

                        $('#' + param.dialogDivID + '').attr('title', '');
                        $('#modal_content').attr('width', '');
                        $('#modal_content').attr('height', '');

                        //return values
                        modalDialog.returnValue();
                    },
                    height: param.dialogHeight,
                    width: param.dialogWidth,
                    resizable: param.dialogResizable,
                    draggable: param.dialogDraggable,
                    closeOnEscape: false
                });

                //close context menu if it exists and it's open
                if (typeof oPopup != "undefined") {
                    if (oPopup.isOpen) {
                        oPopup.hide();
                    }
                } else if (typeof newPopup != "undefined") {
                    if (newPopup.isOpen) {
                        newPopup.hide();
                    }
                }

                //open dialog
                $('#' + param.dialogDivID + '').dialog("option", "title", param.dialogTitle);

                if (param.dialogEvent != null) {
                    var oEvent = param.dialogEvent;
                    $('#' + param.dialogDivID + '').dialog("option", "position", [(oEvent.clientX), (oEvent.clientY)]).dialog("open");
                } else {
                    $('#' + param.dialogDivID + '').dialog({ position: { my: "center", at: "top+15%", of: window } }).dialog("open");
                }
            } else {
                alert("Missing required parameter to open new window.");
            }
        }
    },
    passArgument: function (sArgumentValue, sArgumentName) {
        if (sArgumentValue != null && sArgumentValue != '') {
            if (sArgumentName != null && sArgumentName != '') {
                localStorage.setItem('' + sArgumentName + '', sArgumentValue);
            } else {
                localStorage.setItem('pass_argument', sArgumentValue);
            }
        }
    },
    getArgument: function (sArgumentName) {
        var sReturnValue = "";
        var sSavedName = "";

        if (sArgumentName != null && sArgumentName != '') {
            sSavedName = sArgumentName;
        }else{
            sSavedName = 'pass_argument';
        }

        //Get value from localStorage
        sReturnValue = localStorage.getItem('' + sSavedName + '');
        
        if (sReturnValue != null && sReturnValue != '') {
            //Remove and return local storage value
            localStorage.removeItem('' + sSavedName + '');
            return sReturnValue;
        }
    },
    saveValue: function (sReturnFieldID, sReturnValue) {
        if (sReturnValue != null && sReturnValue != '') {
            localStorage.setItem('return_fieldid', sReturnFieldID);
            localStorage.setItem('return_value', sReturnValue);
        }
    },
    returnValue: function () {
        var sReturnFieldID = "";
        var sReturnValue = "";

        //Get value from localStorage
        sReturnFieldID = localStorage.getItem('return_fieldid');
        sReturnValue = localStorage.getItem('return_value');

        //Replaces showModalDialog's returnValue JS code; usually on buttons page
        if (sReturnFieldID != null && sReturnFieldID != '') { //return value to field in parent page
            $("#" + sReturnFieldID + "").val(""); //clear original value
            $("#" + sReturnFieldID + "").val(sReturnValue); //enter return value from modal dialog
        } else { //return value to modal dialog's default hidden field
            $("#modal-returnValue").val(""); //clear original value
            $("#modal-returnValue").val(sReturnValue); //enter return value from modal dialog
        }

        //Remove local storage value
        localStorage.removeItem('return_fieldid');
        localStorage.removeItem('return_value');
    },
    cancel: function (bReload) {
        //Replaces showModalDialog code that closes the showModalDialog window; usually on buttons page
        try {
            if ($("#modal_dialog").length != 0) {
                //cancel modal div to body tag
                if ($('#modal_dialog').dialog('isOpen') === true) {
                    $('#modal_dialog').dialog('close');
                }
            }
        }
        catch(ex){}

        try {
            if ($("div.content-dialog").length != 0) {
                $("div.content-dialog").each(function () {
                    if ($('#' + $(this).attr("id") + '').dialog('isOpen') === true) {
                        $('#' + $(this).attr("id") + '').dialog('close');
                    }
                });
            }
        }
        catch (ex) {}

        if (bReload === true) {
           window.location.reload(true);
        }
    },
    customize: function (sDialogOption, sDialogValue, bDialogValidate) {
        //customize the parent pages' modal dialog window; works best when child hasn't initialized modalDialog function
        //If so, call function that will open the child page's dialog from parent page; see PlatformList.asp's UpdatePlatform and AddPlatform.
        switch (sDialogOption) {
            case "title":
                //Update title of open Iframe Dialog: ---
                if ($("#modal_dialog").length != 0) {
                    if ($('#modal_dialog').dialog('isOpen') === true) {
                        $('#modal_dialog').dialog({ title: sDialogValue });
                    }
                }
                //Update title of open Content Dialog: --
                if ($("div.content-dialog").length != 0) {
                    $("div.content-dialog").each(function () {
                        if ($('#' + $(this).attr("id") + '').dialog('isOpen') === true) {
                            $('#' + $(this).attr("id") + '').dialog({ title: sDialogValue });
                        }
                    });
                }
                break;
            case "beforeclose":
                //Add beforeClose method to open Iframe Dialog: ---
                if ($("#modal_dialog").length != 0) {
                    if ($('#modal_dialog').dialog('isOpen') === true) {
                        $("#modal_dialog").dialog({
                            beforeClose: function (ev, ui) {
                                if (bDialogValidate !== true) {
                                    return false;
                                } else {
                                    return true;
                                }
                            }
                        });
                    }
                }
                //Add beforeClose method to open Content Dialog: ---
                if ($("div.content-dialog").length != 0) {
                    $("div.content-dialog").each(function () {
                        if ($('#' + $(this).attr("id") + '').dialog('isOpen') === true) {
                            $('#' + $(this).attr("id") + '').dialog({
                                beforeClose: function (ev, ui) {
                                    if (bDialogValidate !== true) {
                                        return false;
                                    } else {
                                        return true;
                                    }
                                }
                            });
                        }
                    });
                }
                break;
            case "show-icon":
                //Show close icon of open Iframe Dialog: ---
                if ($("#modal_dialog").length != 0) {
                    if ($('#modal_dialog').dialog('isOpen') === true) {
                        $('#modal_dialog').dialog({ dialogClass: 'show-close-icon' });
                    }
                }
                //Show close icon of open Content Dialog: --
                if ($("div.content-dialog").length != 0) {
                    $("div.content-dialog").each(function () {
                        if ($('#' + $(this).attr("id") + '').dialog('isOpen') === true) {
                            $('#' + $(this).attr("id") + '').dialog({ dialogClass: 'show-close-icon' });
                        }
                    });
                }
                break;
            case "hide-icon":
                //Hide close icon of open Iframe Dialog: ---
                if ($("#modal_dialog").length != 0) {
                    if ($('#modal_dialog').dialog('isOpen') === true) {
                        $('#modal_dialog').dialog({ dialogClass: 'hide-close-icon' });
                    }
                }
                //Hide close icon of open Content Dialog: --
                if ($("div.content-dialog").length != 0) {
                    $("div.content-dialog").each(function () {
                        if ($('#' + $(this).attr("id") + '').dialog('isOpen') === true) {
                            $('#' + $(this).attr("id") + '').dialog({ dialogClass: 'hide-close-icon' });
                        }
                    });
                }
                break;
            default: break;
        }
    }
};

//************************************************************************
//Description:	To avoid using JS global variables, create the globalVariable  
//              object to get and set values 
//************************************************************************
var globalVariable = {
    save: function (sArgumentValue, sArgumentName) {
        if (sArgumentValue == 0) { //convert zero to string
            sArgumentValue = "0";
        }
        if (sArgumentValue != null && sArgumentValue != '') {
            if (sArgumentName != null && sArgumentName != '') {
                localStorage.setItem('' + sArgumentName + '', ''+sArgumentValue+'');
            } else {
                alert('Can not save global variable, missing parameter.');
            }
        }
    },
    get: function (sArgumentName) {
        var sReturnValue = "";
        var sSavedName = "";

        if (sArgumentName != null && sArgumentName != '') {
            sSavedName = sArgumentName;
        } else {
            alert('Can not save global variable, missing parameter.');
        }

        //Get value from localStorage
        sReturnValue = localStorage.getItem('' + sSavedName + '');

        if (sReturnValue != null && sReturnValue != '') {
            //Remove and return local storage value
            localStorage.removeItem('' + sSavedName + '');
            return sReturnValue;
        }
    }
};

//*****************************************************************
//Description:  Load Jquery UI datepicker for AMO date fields
//Function:     load_datePicker();
//*****************************************************************
function load_datePicker() {
    var $browser = get_browser();
    var $host = window.location.hostname;
    var dCurrentDate = '';
    var sDateFieldID = '';

    //if date field is disabled, remove dateselection class from field else make field readonly 
    $("input[type='text'].dateselection, input[type='text'].dateselection-validate").each(function (index, key) {
        if ($(this).prop('disabled')) {
            if ($(this).hasClass("dateselection")) {
                $(this).removeClass("dateselection");
            } else if ($(this).hasClass("dateselection-validate")) {
                $(this).removeClass("dateselection-validate");
            }
        } else {
            $(this).change(function () {
                validateDate(this);
            });
        }
    });

    //Loop thru each field with dateselection class and add jquery datepicker widget
    if ($(".dateselection").length > 0) {
        $(".dateselection").datepicker({
            showOn: 'button',
            buttonText: 'Select Date',
            buttonImageOnly: true,
            buttonImage: 'http://'+$host+'/Excalibur/images/calendar.gif',
            dateFormat: "mm/dd/yy",
            changeMonth: true,
            changeYear: true,
            firstDay: 7
        });
    }

    //Loop thru each field with dateselection-validate class and add jquery datepicker widget
    //Can use dateselection-validate when JQuery datePicker needs to validate selected date using page's JS function
    if ($(".dateselection-validate").length > 0) {
        $(".dateselection-validate").datepicker({
            showOn: 'button',
            buttonText: 'Select Date',
            buttonImageOnly: true,
            buttonImage: 'http://' + $host + '/Excalibur/images/calendar.gif',
            dateFormat: "mm/dd/yy",
            changeMonth: true,
            changeYear: true,
            firstDay: 7,
            beforeShow: function (input, inst) {
                //get existing date in dateselection field
                dCurrentDate = $(this).datepicker("getDate");
                dCurrentDate = $.datepicker.formatDate("mm/dd/yy", dCurrentDate)
                if (dCurrentDate != null || dCurrentDate != '') {
                    globalVariable.save(dCurrentDate, 'date_selection_currentdate');
                }
            },
            onSelect: function (dateText, inst) {
                dCurrentDate = globalVariable.get('date_selection_currentdate');
                sDateFieldID = $(this).attr('id');
                //if exists on page, process validate function; can add more functions
                if (typeof cmdDate_ScheduleResult !== 'undefined' && typeof cmdDate_ScheduleResult === 'function') {
                    cmdDate_ScheduleResult(sDateFieldID, dCurrentDate);
                }
            }
        });
    }

    $(".ui-datepicker-trigger").css('cursor', 'pointer');
    if ($browser === 'Chrome' || $browser === 'Netscape') {
        $(".ui-datepicker-trigger").css("margin-bottom", "-6px");
    } else {
        $(".ui-datepicker-trigger").css("margin-bottom", "-4px");
    }
    $(".ui-datepicker-trigger").css("margin-left", "3px");

    $(".ui-datepicker").css('margin-left', '-50px', 'margin-top', '-50px');

}

//*****************************************************************
//Description:  Detect Browser Type
//Function:     get_browser();
//*****************************************************************
function get_browser() {
    var N = navigator.appName, ua = navigator.userAgent, tem;
    var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
    if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
    M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
    return M[0];
}

function validateDate(e) {
    var originalMonth = e.value.split("/")[0];
    var originalDay = e.value.split("/")[1];
    var originalYear = e.value.split("/")[2];
    if (originalMonth === undefined || originalDay === undefined || originalYear === undefined || e.value.split("/")[3] !== undefined) {
        alert("Date must be in mm/dd/yyyy format.");
        e.value = "";
        return false;
    }

    var date = new Date(originalYear, originalMonth - 1, originalDay);
    if (date.getFullYear() != originalYear
        || date.getMonth() != originalMonth - 1
        || date.getDate() != originalDay) {
        alert("Invalid date detected : " + e.value);
        e.value = "";
        return false;
    }
}


<%@  language="VBScript" %>
<!-- #include file = "../includes/noaccess.inc" -->
<html>
<head>
    <title>Agency</title>
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script language="JavaScript" src="../includes/client/Common.js"></script>
    <script>
        $(function () {
            $("input:button").button();
        });
    </script>
    <script id="clientEventHandlersJS" language="javascript">
<!--

    function VerifyEmail(src) {
        var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
        var regex = new RegExp(emailReg);
        return regex.test(src);
    }

    function VerifyStatus() {
        with (window.parent.frames["UpperWindow"].frmStatus) {
            switch (cboStatus.value) {
                case 'L':
                    if (!validateTextInput(txtLeveragedSystem, 'Leveraged System Selection')) { return false; }
                    break;
                case 'NS':
                    return true;
                    break;
                case 'O':
                    //if (!validateTextInput(txtProjectedDate, 'Availability Date')){	return false; }
                    if (!validateDateInput(txtProjectedDate, 'Availability Date')) { return false; }
                    break;
                case 'SU':
                    //if (!validateTextInput(txtProjectedDate, 'Availability Date')){	return false; }
                    if (!validateDateInput(txtProjectedDate, 'Availability Date')) { return false; }
                    break;
                case 'C':
                    break;
                case 'P':
                    if (!validateTextInput(txtNotes, 'Notes')) { return false; }
                    break;
                default:
                    break;
            }

            if (cboPorDcr != undefined && !validateTextInput(cboPorDcr, 'Added By')) { return false; }
            if (cboDcr != undefined && !validateTextInput(cboDcr, 'Added By DCR')) { return false; }
        }
        return true;
    }

    function VerifyLeverage() {
        var blnSuccess = true;
        return blnSuccess;
    }

    function cmdCancel_onclick(pulsarplusDivId) {
        //pulsarplusDivId should be a parameter 
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            var iframeName = parent.window.name;
            if (iframeName != '') {
                parent.window.parent.CloseIframeDialog();
            } else {
               
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel();
                } else {     
                    window.parent.close();
                }
            }
        }
    }

    function getSelectValues(select) {
      var result = [];
      var options = select && select.options;
      var opt;

      for (var i=0, iLen=options.length; i<iLen; i++) {
        opt = options[i];

        if (opt.selected) {
          result.push(opt.value || opt.text);
        }
      }
      return result;
    }
    
    function cmdOK_onclick(pulsarplusDivId, CountryID, ReleaseID) {
        var blnAll = true;
        var i;
        var sReturnValue;

        if (window.parent.frames["UpperWindow"].frmStatus.hidBatchUpdate.value == 'True') {
            //Batch Update
            var releaseSelectValue = getSelectValues(window.parent.frames["UpperWindow"].frmStatus.release_select);
            var countrySelectValue = getSelectValues(window.parent.frames["UpperWindow"].frmStatus.country_select);

            if (releaseSelectValue == '' && countrySelectValue == '') {
                alert("Batch update must be select one more on Release Selection or Country Selection!");
                return;
            }

            window.parent.frames["UpperWindow"].frmStatus.hidBatchUpdateRelease.value = releaseSelectValue;
            window.parent.frames["UpperWindow"].frmStatus.hidBatchUpdateCountry.value = countrySelectValue;
            window.parent.frames["UpperWindow"].frmStatus.hidSave.value = true;
            window.parent.frames["UpperWindow"].frmStatus.hidClose.value = true;
            window.parent.frames["UpperWindow"].frmStatus.submit();
            parent.window.parent.modalDialog.cancel();
        }
        else if (CountryID != '' && ReleaseID != '') {
            //once update
            sReturnValue = window.parent.frames["UpperWindow"].frmStatus.cboStatus.value;
            if (sReturnValue == 2 && window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value == '') {
                alert('Please Select a Date to Save.');
            }

            window.frmButtons.cmdCancel.disabled = true;
            window.frmButtons.cmdOK.disabled = true;
            window.parent.frames["UpperWindow"].frmStatus.hidSave.value = true;
            window.parent.frames["UpperWindow"].frmStatus.hidClose.value = true;
            window.parent.frames["UpperWindow"].frmStatus.submit();
            parent.window.parent.modalDialog.cancel();
        }
        else {
            //original
            if (window.parent.frames["UpperWindow"].frmStatus) {
                if (VerifyStatus()) {
                    sReturnValue = window.parent.frames["UpperWindow"].frmStatus.cboStatus.value;
                    if (sReturnValue == 'O' && window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value != '') {
                        sReturnValue = window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value;
                    }
                    if (sReturnValue == 'SU' && window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value != '') {
                        sReturnValue = window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value;
                    }
                    window.frmButtons.cmdCancel.disabled = true;
                    window.frmButtons.cmdOK.disabled = true;
                    if (pulsarplusDivId == undefined || pulsarplusDivId == "") {
                        if (parent.window.parent.document.getElementById('modal_dialog')) {
                            parent.window.parent.ShowAgencyStatusResults(sReturnValue);
                        } else {
                            window.returnValue = sReturnValue;
                        }
                    }
                    window.parent.frames["UpperWindow"].frmStatus.hidSave.value = true;
                    window.parent.frames["UpperWindow"].frmStatus.hidClose.value = true;
                    window.parent.frames["UpperWindow"].frmStatus.submit();
                }
            }
            else if (window.parent.frames["UpperWindow"].frmLeverage) {
                if (VerifyLeverage()) {
                    window.parent.frames["UpperWindow"].frmLeverage.SaveMode.value = true;
                    window.parent.frames["UpperWindow"].frmLeverage.submit();
                }
            }
        }

        parent.parent.location.reload();

        return;
    }

    function document_OnLoad() {
        if (typeof (window.parent.frames["UpperWindow"].document.all["hidEdit"]) == 'object') {
            if (window.parent.frames["UpperWindow"].document.all["hidEdit"].value.toLowerCase() == 'false') {
                window.frmButtons.cmdOK.disabled = true;
            }
        }
    }

//-->
    </script>
    <style type="text/css">
        .modal_button {
            background-color: #0096d6 !important;
            border: none;
            color: #00bfff;
            padding: 5px 15px !important;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 11px;
        }
    </style>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
    <form id="frmButtons" action="AgencyButtons.asp" method="post">
        <table border="0" cellspacing="1" cellpadding="1" align="right">

            <%If Request.QueryString("pulsarplusDivId") <> "" Then
            %>
            <tr>
                <td>
                    <input type="button" value="Save" id="cmdOK" name="cmdOK" class="" onclick="return cmdOK_onclick('<%=Request("pulsarplusDivId")%>')" />
                </td>
                <td>
                    <%If Request.QueryString("AgencyPage") <> "" Then%>
                    <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" class="" onclick="return cmdCancel_onclick('')" />
                    <%Else %>
                    <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" class="" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')" />
                    <%End if %>
                </td>
            </tr>
            <%Else	%>
            <tr>
                <td>
                    <input type="button" value="Save" id="cmdOK" name="cmdOK" class="modal_button" onclick="return cmdOK_onclick('<%=Request("pulsarplusDivId")%>', '<%=Request("CountryID")%>', '<%=Request("ReleaseID")%>')" />
                </td>
                <td>
                    <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" class="modal_button" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')" />
                </td>
            </tr>
            <%End If %>
        </table>
    </form>
</body>
</html>

<%@  language="VBScript" %>


<html>
<head>
    <meta name="VI60_DefaultClientScript" content="JavaScript">

    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">


    <script id="clientEventHandlersJS" language="javascript">
<!--

    function isValidEmail(strEmail) {
        var emailReg = "^[\\w-_\.+#]*[\\w-_\.#]\@([\\w]+\\.)+[\\w]+[\\w]$";
        var regex = new RegExp(emailReg);
        return regex.test(strEmail);
    }

    function isValidEmailList(strList) {
        var AddressArray;

        var i;
        AddressArray = strList.split(";");
        for (i = 0; i < AddressArray.length; i++) {
            if (!isValidEmail(AddressArray[i].replace(/^\s+|\s+$/g, ""))) {
                if (AddressArray[i].replace(/^\s+|\s+$/g, "") != "") {
                    alert(AddressArray[i] + " is not a valid email address.");
                    window.parent.frames["UpperWindow"].frmMain.txtNotify.focus();
                    return false;
                }
                else if (AddressArray[i].replace(/^\s+|\s+$/g, "") == "" && i != (AddressArray.length - 1)) {
                    alert("Missing email address.  Ensure that you have only one semicolon between each address.");
                    window.parent.frames["UpperWindow"].frmMain.txtNotify.focus();
                    return false;
                }
            }
        }
        return true;
    }

    function VerifySave() {
        window.parent.frames["UpperWindow"].frmMain.txtNotify.value = window.parent.frames["UpperWindow"].frmMain.txtNotify.value.replace(/;;/g, ";");
        return isValidEmailList(window.parent.frames["UpperWindow"].frmMain.txtNotify.value);

    }

    function cmdCancel_onclick() {
        var pulsarplusDivId = document.getElementById("pulsarplusDivId");
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel(false);
            } else {
                window.parent.close();
            }
        }
    }

    function cmdFinish_onclick() {
        if (VerifySave()) {
            cmdCancel.disabled = true;
            cmdNext.disabled = true;
            cmdPrevious.disabled = true;
            cmdFinish.disabled = true;
            cmdSave.disabled = true;
            window.parent.frames["UpperWindow"].AlertPreviewSection.innerHTML = "";
             window.parent.frames["UpperWindow"].frmMain.txtIsRTMAsDraft.value = 0;
            window.parent.frames["UpperWindow"].frmMain.txtEmailPreview.value = window.parent.frames["UpperWindow"].PreviewPage.innerHTML;
            window.parent.frames["UpperWindow"].frmMain.submit();
        }
    }

    function cmdSave_onclick() {
        if (VerifySave()) {
             cmdCancel.disabled = true;
            cmdNext.disabled = true;
            cmdPrevious.disabled = true;
            cmdFinish.disabled = true;
            cmdSave.disabled = true;
            window.parent.frames["UpperWindow"].frmMain.txtIsRTMAsDraft.value = 1;
            window.parent.frames["UpperWindow"].frmMain.submit();
        }
    }

    function cmdPrevious_onclick() {
        switch (window.parent.frames["UpperWindow"].CurrentState) {
            case "Preview":
                window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                break;
            case "Alerts":
                if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Patches";
                else if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Images";
                else if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Restore";
                else if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                else if (window.parent.frames["UpperWindow"].frmMain.chkFW.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "FW";
                else if (window.parent.frames["UpperWindow"].frmMain.chkSCMX.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "SCMX";
                else
                    window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "Patches":
                if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Images";
                else if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Restore";
                else if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                else
                    window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "Images":
                if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "Restore";
                else if (window.parent.frames["UpperWindow"].frmMain.chkFW.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "FW";
                else if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                else
                    window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "Restore":
                if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                    window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                else
                    window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "BIOS":
                window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "FW": //standalone with images
                window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
            case "SCMX": //standalone
                window.parent.frames["UpperWindow"].CurrentState = "General";
                break;
        }
        window.parent.frames["UpperWindow"].ProcessState();
    }

    function cmdNext_onclick() {
        var blnSuccess = true;
        var strSeriesList = "";
        if (window.parent.frames["UpperWindow"].ValidateTab(window.parent.frames["UpperWindow"].CurrentState)) {
            switch (window.parent.frames["UpperWindow"].CurrentState) {
                case "General":
                    if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkFW.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "FW";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkSCMX.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "SCMX";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Restore";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";

                    if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked && !window.parent.frames["UpperWindow"].frmMain.chkRestore.checked && !window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].LoadAlerts("&ReportType=2");
                    else
                        window.parent.frames["UpperWindow"].LoadAlerts("");
                    break;
                case "BIOS":
                    if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Restore";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "FW":
                    if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "Restore":
                    if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "Images":
                    if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "Patches":
                    window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "SCMX":
                    window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                    break;
                case "Alerts":
                    window.parent.frames["UpperWindow"].CurrentState = "Preview";
                    break;
            }

            if (blnSuccess)
                window.parent.frames["UpperWindow"].ProcessState();
        }
    }



    function getSelectedRadioValue(buttonGroup) {
        var i = getSelectedRadio(buttonGroup);
        if (i == -1) {
            return "";
        } else {
            if (buttonGroup[i]) {
                return buttonGroup[i].value;
            } else {
                return buttonGroup.value;
            }
        }
    }


    function getSelectedRadio(buttonGroup) {
        if (1 == 0) {

        }
        else {
            if (buttonGroup[0]) {
                for (var i = 0; i < buttonGroup.length; i++) {
                    if (buttonGroup[i].checked) {
                        return i
                    }
                }
            } else {
                if (buttonGroup.checked) { return 0; }
            }
            return -1;
        }
    }


//-->
    </script>
</head>
<style>
    input {
        FONT-SIZE: 10pt;
        FONT-FAMILY: Verdana;
    }
</style>
<body bgcolor="ivory">

    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>

            <td>
                <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" language="javascript" onclick="return cmdCancel_onclick()"></td>
            <td width="10"></td>
            <td>
                <input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" language="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
            <td>
                <input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" language="javascript" onclick="return cmdNext_onclick()"></td>
            <td width="10"></td>
              <td> <input type="button" value="Save" id="cmdSave" name="cmdSave" disabled language="javascript" onclick="return cmdSave_onclick()"></td>
            <td width="10"></td>
            <td>
                <input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled language="javascript" onclick="return cmdFinish_onclick()"></td>

        </tr>
    </table>
</body>
</html>

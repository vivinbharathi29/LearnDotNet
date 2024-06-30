<%@ Language=VBScript %>
<%' 03/11/2016, Herb, Merged with PBI 16007 and changeset 15097%>
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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

    //herb
    function VerifySave() {
        var elmNotify;
        var arrNotify;
        elmNotify = window.parent.frames["UpperWindow"].frmMain.txtNotify;
        if (elmNotify.length == "undefined"){
            arrNotify = new Array();
            arrNotify[0]=elmNotify;
        }else{
            arrNotify = elmNotify;
        }

        var isValidEmailListAll = true;
        for(var j =0; j < arrNotify.length; j++){
            arrNotify[j].value = arrNotify[j].value.replace(/;;/g, ";");
            isValidEmailListAll = isValidEmailListAll && isValidEmailList(arrNotify[j].value);
        }
        
        return isValidEmailListAll;

    }

    function cmdCancel_onclick() {
        var pulsarplusDivId = document.getElementById("pulsarplusDivId");
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                var bRefresh = false;
                if (cmdCancel.value == "Close") {
                    bRefresh = true;
                }
                parent.window.parent.modalDialog.cancel(bRefresh);
            } else {
                if (cmdCancel.value == "Close") {
                    window.returnValue = 1;
                }
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
            window.parent.frames["UpperWindow"].document.all.tabPreview.style.display = "none";
            window.parent.frames["UpperWindow"].document.all.lblTitle.innerText = "Processing ... ... ...     Please Wait.";
            cmdCancel.value = "Close";

            var boolCon = confirm("Start processing RTM.");
            if (boolCon == true) {
                window.parent.frames["UpperWindow"].submitMutipleRTM();
                window.parent.frames["UpperWindow"].document.all.lblTitle.innerText = "Now you can close this window.";
                cmdCancel.disabled = false;

            } else {
                cmdCancel.disabled = false;
                cmdNext.disabled = true;
                cmdPrevious.disabled = false;
                cmdFinish.disabled = false;
                window.parent.frames["UpperWindow"].document.all.tabPreview.style.display = "";
                window.parent.frames["UpperWindow"].document.all.lblTitle.innerText = "Review Selected Information";
                cmdCancel.value = "Cancel";
            }

            //window.parent.frames["UpperWindow"].AlertPreviewSection.innerHTML = "";
            //window.parent.frames["UpperWindow"].frmMain.submit();
        }

    }

    function cmdPrevious_onclick() {
        switch (window.parent.frames["UpperWindow"].CurrentState) {
            case "Preview":
                window.parent.frames["UpperWindow"].CurrentState = "Alerts";
                break;
            case "Alerts":
                //Herb
                if (parseInt(window.parent.frames["UpperWindow"].intAlertStatus) == parseInt(window.parent.frames["UpperWindow"].intProducts) - 1) {
                    if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Restore";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "BIOS";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "General";                    
                } else {
                    preAlert();
                }

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
                    else if (window.parent.frames["UpperWindow"].frmMain.chkRestore.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Restore";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Images";
                    else if (window.parent.frames["UpperWindow"].frmMain.chkPatch.checked)
                        window.parent.frames["UpperWindow"].CurrentState = "Patches";
                    else
                        window.parent.frames["UpperWindow"].CurrentState = "Alerts";

                    if (window.parent.frames["UpperWindow"].frmMain.chkBIOS.checked && !window.parent.frames["UpperWindow"].frmMain.chkRestore.checked && !window.parent.frames["UpperWindow"].frmMain.chkImages.checked)
                        //window.parent.frames["UpperWindow"].LoadAlerts("&ReportType=2");
                        window.parent.frames["UpperWindow"].LoadAlertsByProductAll("&ReportType=2");
                    else
                        window.parent.frames["UpperWindow"].LoadAlertsByProductAll("");
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
                case "Alerts":
                    //Herb
                    if (window.parent.frames["UpperWindow"].intAlertStatus ==0) {
                        window.parent.frames["UpperWindow"].CurrentState = "Preview";
                    } else {
                        nextAlert();
                    }
                    
                    break;
            }

            if (blnSuccess)
                window.parent.frames["UpperWindow"].ProcessState();
        }
    }

    //Herb
    function nextAlert() {
        window.parent.frames["UpperWindow"].intAlertStatus -= 1;
    }

    //Herb
    function preAlert() {
        window.parent.frames["UpperWindow"].intAlertStatus += 1;
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
</SCRIPT>
</head>
<STYLE>
input
{
    FONT-SIZE: 10pt;	
    FONT-FAMILY: Verdana;	
}
</STYLE>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>

			<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>



<!--		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
--></TR></table>
     <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
</body>
</html>
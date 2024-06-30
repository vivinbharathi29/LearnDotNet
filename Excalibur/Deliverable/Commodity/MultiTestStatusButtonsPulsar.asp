<%@ Language=VBScript %>

<html>
<head>
    <title></title>
    <meta name="VI60_DefaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" type="text/javascript">
    <!--

    function VerifySave() {
        var blnSuccess = true;
        var blnFound = false;
        var blnFoundComplete = false;

        var i;


        blnFoundComplete = false;
        if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotEngineer").value == "True") {
            var strRequired = window.parent.frames["UpperWindow"].document.getElementById("txtPilotDateRequired").value.indexOf(window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value);
            if (window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value == "")
                strRequired = -1


            //Pilot - Clear out invalid formatted dates if Date is not the selected status
            //if ( (! isDate(window.parent.frames["UpperWindow"].frmStatus.txtPilotDate.value)) && window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.selectedIndex].value != "2")
            if ((!isDate(window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value)) && strRequired == -1)
                window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value = "";


            //Look for a problem with "Complete" status, if status = dropped
            if (window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").option[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value == "6") {
                if (typeof (window.parent.frames["UpperWindow"].document.getElementById("txtMultiID").length) == "undefined") {
                    if (window.parent.frames["UpperWindow"].document.getElementById("txtMultiID").TestStatus != "5" && window.parent.frames["UpperWindow"].document.getElementById("txtMultiID").checked)
                        blnFoundComplete = true;
                }
                else {
                    for (i = 0; i < window.parent.frames["UpperWindow"].document.getElementById("txtMultiID").length; i++)
                        if (window.parent.frames["UpperWindow"].document.getElementById("txtMultiID" + i).TestStatus != "5" && window.parent.frames["UpperWindow"].document.getElementById("txtMultiID" + i).checked)
                            blnFoundComplete = true;
                }
            }
            if (blnFoundComplete && window.parent.frames["UpperWindow"].document.getElementById("txtCommodityPM").value == "True" && window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value == "6") {
                if (window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value == "5")
                    blnFoundComplete = false;
            }

            if (blnFoundComplete) {
                alert("You can not set the Pilot status to \"Complete\" if any deliverables selected are not \"QComplete\".");
                window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").focus();
                blnSuccess = false;
            }
                //			else if (window.parent.frames["UpperWindow"].frmStatus.txtPilotDate.value != "" && window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.selectedIndex].value == "2" && (! isDate(window.parent.frames["UpperWindow"].frmStatus.txtPilotDate.value)))
            else if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value != "" && strRequired != -1 && (!isDate(window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value))) {
                alert("You must supply a valid pilot date if one is entered.");
                window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").focus();
                blnSuccess = false;
            }
                //			else if (window.parent.frames["UpperWindow"].frmStatus.txtPilotDate.value == "" && window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboPilotStatus.selectedIndex].value =="2")
            else if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value == "" && strRequired != -1) {
                alert("You must supply a valid pilot date.");
                window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").focus();
                blnSuccess = false;
            }
            else if (window.parent.frames["UpperWindow"].document.getElementById("txtCommodityPM").value != "True" && window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value == "") {
                alert("You must enter a new pilot status.");
                window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").focus();
                blnSuccess = false;
            }

            if (window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].value == "2")
                window.parent.frames["UpperWindow"].document.getElementById("txtPilotStatusText").value = window.parent.frames["UpperWindow"].frmStatus.txtPilotDate.value;
            else
                window.parent.frames["UpperWindow"].document.getElementById("txtPilotStatusText").value = window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboPilotStatus").selectedIndex].text;


        }


        //Check Qualification Status.
        if (blnSuccess && window.parent.frames["UpperWindow"].document.getElementById("txtCommodityPM").value == "True") {
            //Qual - Clear out invalid formatted dates if Date is not the selected status
            if ((!isDate(window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").value)) && window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value != "3")
                window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value = "";


            if (window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").value != "" && window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value == "3" && (!isDate(window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").value))) {
                alert("You must supply a valid date if one is entered.");
                window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").focus();
                blnSuccess = false;
            }
            else if (window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").value == "" && window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value == "3") {
                alert("You must supply a valid date.");
                window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").focus();
                blnSuccess = false;
            }
            else if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotEngineer").value != "True" && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value == "") {
                alert("You must enter a new status.");
                window.parent.frames["UpperWindow"].document.getElementById("cboStatus").focus();
                blnSuccess = false;
            }

            if (window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value == "3")
                window.parent.frames["UpperWindow"].document.getElementById("txtStatusText").value = window.parent.frames["UpperWindow"].document.getElementById("txtTestDate").value;
            else
                window.parent.frames["UpperWindow"].document.getElementById("txtStatusText").value = window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].text;
        }



        return blnSuccess;
    }

    function cmdCancel_onclick(pulsarplusDivId) {
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else { 
            window.parent.Cancel();
        }
    }

    function cmdOK_onclick() {
        var blnAll = true;
        var i;
        if (window.parent.frames["UpperWindow"].document.getElementById("txtUpdatableVersionCount").value == "" || window.parent.frames["UpperWindow"].document.getElementById("txtUpdatableVersionCount").value == "0") {
            window.parent.close();
            return;
        }

        if (VerifySave()) {
            cmdCancel.disabled = true;
            cmdOK.disabled = true;
            window.parent.frames["UpperWindow"].frmStatus.submit();
        }

    }

    function enableButton() {
        cmdCancel.disabled = false;
        cmdOK.disabled = false;
    }
    //-->
</script>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"  ></TD>
</TR></table>
</body>
</html>

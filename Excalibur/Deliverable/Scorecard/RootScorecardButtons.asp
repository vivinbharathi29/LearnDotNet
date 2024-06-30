<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
    var blnSuccess = true;

    if (window.parent.frames["UpperWindow"].frmMain.txtExecutiveSummary.value.length > 400 && window.parent.frames["UpperWindow"].frmMain.txtExecutiveSummary.value != window.parent.frames["UpperWindow"].frmMain.txtExecutiveSummaryTemplate.value) {
        alert("The Executive Summary field can not contain more than 400 characters");
        window.parent.frames["UpperWindow"].frmMain.txtExecutiveSummary.focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].frmMain.txtHPPeopleProcess.value.length > 400 && window.parent.frames["UpperWindow"].frmMain.txtHPPeopleProcess.value != window.parent.frames["UpperWindow"].frmMain.txtHPPeopleProcessTemplate.value) {
        alert("The HP People and Process field can not contain more than 400 characters");
        window.parent.frames["UpperWindow"].frmMain.txtHPPeopleProcess.focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].frmMain.txtHPEquipment.value.length > 400 && window.parent.frames["UpperWindow"].frmMain.txtHPEquipment.value != window.parent.frames["UpperWindow"].frmMain.txtHPEquipmentTemplate.value) {
        alert("The HP Equipment field can not contain more than 400 characters");
        window.parent.frames["UpperWindow"].frmMain.txtHPEquipment.focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].frmMain.txtSupplierPeopleProcess.value.length > 400 && window.parent.frames["UpperWindow"].frmMain.txtSupplierPeopleProcess.value != window.parent.frames["UpperWindow"].frmMain.txtSupplierPeopleProcessTemplate.value) {
        alert("The Supplier People and Process field can not contain more than 400 characters");
        window.parent.frames["UpperWindow"].frmMain.txtSupplierPeopleProcess.focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].frmMain.txtSupplierDeliverables.value.length > 400 && window.parent.frames["UpperWindow"].frmMain.txtSupplierDeliverables.value != window.parent.frames["UpperWindow"].frmMain.txtSupplierDeliverablesTemplate.value) {
        alert("The Supplier Software field can not contain more than 400 characters");
        window.parent.frames["UpperWindow"].frmMain.txtSupplierDeliverables.focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].frmMain.txtAction.value.length > 800 && window.parent.frames["UpperWindow"].frmMain.txtAction.value != window.parent.frames["UpperWindow"].frmMain.txtActionTemplate.value) {
        alert("The Actions field can not contain more than 800 characters");
        window.parent.frames["UpperWindow"].frmMain.txtAction.focus();
        blnSuccess = false;
    }

	return blnSuccess;
}

function cmdCancel_onclick() {
		    var iframeName = parent.window.name;
    if (iframeName != '') {
        parent.window.parent.ClosePopUp();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =false;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmMain.submit();
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>
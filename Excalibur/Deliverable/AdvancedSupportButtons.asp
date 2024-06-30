<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<script src="../Scripts/Pulsar2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	//if(window.parent.frames["UpperWindow"].frmMain.txtEOLDate.value != "")
		//if(! isDate(window.parent.frames["UpperWindow"].frmMain.txtEOLDate.value))
		//	{
		//	blnSuccess=false;
		//	alert("EOL Date is not a valid date.");
		//	window.parent.frames["UpperWindow"].frmMain.txtEOLDate.focus();
		//	}
	return blnSuccess;
}

function cmdCancel_onclick() {
    if (isFromPulsar2()) {
        closePulsar2Popup(false);
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    } else {
        window.parent.Cancel();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmTarget.submit();
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
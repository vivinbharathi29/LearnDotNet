<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script language="JavaScript" src="../includes/client/Common.js.old"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifyStatus()
{
	return true;
	
	with (window.parent.frames["UpperWindow"].frmMain)
	{
		if (!validateTextInput(cboDcr, 'DCR')){	return false; }
	}
	return true;
}

function cmdCancel_onclick() 
{
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() 
{
	var sReturnValue;

	if (VerifyStatus())
	{
		sReturnValue = window.parent.frames["UpperWindow"].frmMain.cboDcr.value;
		window.parent.frames["UpperWindow"].frmMain.submit();
	}
	
	return;
}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Save" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>
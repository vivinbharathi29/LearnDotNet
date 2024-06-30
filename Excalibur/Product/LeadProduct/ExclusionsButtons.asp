<%@ Language=VBScript %>

<!-- #include file = "../../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	if (window.parent.frames["UpperWindow"].frmMain.optSelected.checked)
		{
		if ( typeof(window.parent.frames["UpperWindow"].frmMain.chkVersions.length) == "undefined")
			{
			if (!window.parent.frames["UpperWindow"].frmMain.chkVersions.checked)
				{
				blnSuccess=false;
				alert("You must select at least one version.");
				}
			}
		else
			{
			var blnFound = false;
			var i;
			for (i=0;i<	window.parent.frames["UpperWindow"].frmMain.chkVersions.length;i++)
				{
				if (window.parent.frames["UpperWindow"].frmMain.chkVersions(i).checked)
					blnFound=true;
				}
			if (!blnFound)
				{
				blnSuccess=false;
				alert("You must select at least one version.");
				}
			}
		}
	if (window.parent.frames["UpperWindow"].frmMain.txtComments.length > 8000  && blnSuccess)
		{
		alert("The comments can not be more that 8000 characters.");
		window.parent.frames["UpperWindow"].frmMain.txtComments.focus();
		blnSuccess=false;
		}
	return blnSuccess;
}

function cmdCancel_onclick() {
    //window.parent.close();
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {

        window.parent.close();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
		cmdCancel.disabled =true;
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
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>
</body>
</html>
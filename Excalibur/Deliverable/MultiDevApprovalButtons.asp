<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function VerifySave(){
	var blnSuccess = true;	
	var blnFound = false;
	var blnFoundComplete = false;
	
	var i;
	
	
	if (window.parent.frames["UpperWindow"].frmStatus.txtComments.value == "" && window.parent.frames["UpperWindow"].frmStatus.NewValue.value=="2")
		{
		alert("You must enter comments when rejecting a release.");
		window.parent.frames["UpperWindow"].frmStatus.txtComments.focus();
		blnSuccess = false;
		}
			

	return blnSuccess;
}

function cmdCancel_onclick() {
    window.parent.Cancel();
}

function cmdOK_onclick() {
	var blnAll = true;
	var i;

	if (VerifySave())
		{

			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmStatus.submit();
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
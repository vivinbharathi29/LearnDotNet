<%@ Language=VBScript %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifyStatus()
{
	with (window.parent.frames["UpperWindow"].frmMain)
	{
		if (selAvCategory.value == 0)
		{
			alert('Please choose a valid Feature Category');
			return false;
		}
	}
	return true;
}

function cmdCancel_onclick() 
{
	window.parent.close();
}

function cmdOK_onclick() 
{
	var blnAll = true;
	var i;
	var sReturnValue;
	
	if (VerifyStatus())
	{
		window.frmButtons.cmdOK.disabled =true;
		window.frmButtons.cmdCancel.disabled =true;
		window.parent.frames["UpperWindow"].frmMain.hidFunction.value="save";
		window.parent.frames["UpperWindow"].frmMain.submit();
	}
	
	return;
}

function document_OnLoad()
{
	window.frmButtons.cmdOK.disabled = true;
	if (typeof(window.parent.frames["UpperWindow"].document.all["hidMode"]) == 'object')
	{
		//if (window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add'||window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit')
			window.frmButtons.cmdOK.disabled = false;
	}
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmButtons"  action=avButtons.asp method=post>
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Save" id=cmdOK name=cmdOK onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>
</FORM>
</body>
</html>
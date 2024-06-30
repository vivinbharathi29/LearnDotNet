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
	if (VerifyStatus())
	{
		window.frmButtons.cmdOK.disabled =true;
		window.frmButtons.cmdCancel.disabled =true;
		window.parent.frames["UpperWindow"].frmMain.hidMode.value="save";
		window.parent.frames["UpperWindow"].frmMain.submit();
	    //call the pop-up close here will cause the saving not called some times, so move the closing to the kmatDetail.asp after the mode is set to 'Close' liek the old asp way

		//if (window.parent.frames["UpperWindow"]) {
		//    parent.window.parent.modalDialog.cancel(true);
		//} else {
		//    window.parent.close();
		//}
	}
	
	return;
}

function document_OnLoad()
{
	window.frmButtons.cmdOK.disabled = true;
	if (typeof(window.parent.frames["UpperWindow"].document.all["hidMode"]) == 'object')
	{
		if (window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add'||window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit')
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
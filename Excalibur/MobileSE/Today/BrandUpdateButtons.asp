<%@ Language=VBScript %>

<!-- #include file = "../../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function ltrim ( s ) 
{ 
	return s.replace( /^\s*/, "" ) 
} 

function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

function VerifySave(){
	var blnSuccess = true;	
	
    if (window.parent.frames["UpperWindow"].frmMain.cboNew.selectedIndex == 0)
		{
		alert("You must select the new Brand.");
		window.parent.frames["UpperWindow"].frmMain.cboNew.focus();
		blnSuccess = false;
		}
	return blnSuccess;
}

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
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
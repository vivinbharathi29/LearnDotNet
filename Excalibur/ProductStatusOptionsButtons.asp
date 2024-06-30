<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "includes/Date.asp" -->

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
	if ( (window.parent.frames["UpperWindow"].frmMain.txtStartDt.value != "") && (! isDate(window.parent.frames["UpperWindow"].frmMain.txtStartDt.value)) )
	{
		alert('Start Date must be a valid date format.');
		window.parent.frames["UpperWindow"].frmMain.txtStartDt.focus();
		return false
	}
	if ( (window.parent.frames["UpperWindow"].frmMain.txtEndDt.value != "") && (! isDate(window.parent.frames["UpperWindow"].frmMain.txtEndDt.value)) )
	{
		alert('End Date must be a valid date format.');
		window.parent.frames["UpperWindow"].frmMain.txtEndDt.focus();
		return false
	}
	
	return true;
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else{
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].cmdOK_onclick();
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"  ></TD>
</TR></table>
</body>
</html>
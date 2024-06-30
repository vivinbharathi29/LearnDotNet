<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

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
	
	return blnSuccess;
}

function cmdCancel_onclick() {
    var iframeName = parent.window.name;
    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {

        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    } else {
        if (iframeName != '') {
            parent.window.parent.ClosePropertiesDialog();
        } else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			if (typeof(window.parent.frames["UpperWindow"].frmRequirement.txtSpecification) != "undefined")
				window.parent.frames["UpperWindow"].frmRequirement.txtSpecification.value = window.parent.frames["UpperWindow"].frames["myEditor"].document.body.innerHTML;
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmRequirement.submit();
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">

    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>
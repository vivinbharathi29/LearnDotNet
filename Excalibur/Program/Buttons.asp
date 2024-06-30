<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

function VerifySave(){
	var blnSuccess = true;	
		//if (window.parent.frames["UpperWindow"].frmChange.txtSummary.value == "")
	//	{
	//		alert("You must specify a summary of your image strategy");
	//		window.parent.frames["UpperWindow"].frmChange.txtSummary.focus();
	//		blnSuccess = false;
	//	}
	
	return blnSuccess;
}

function cmdCancel_onclick(pulsarplus) {
    if (pulsarplus == 'true') {
         parent.window.parent.closeExternalPopup();
    }
    else if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() {
	var blnOK = true;
	var i;
	if (VerifySave())
		{
			//Name
			if (window.parent.frames["UpperWindow"].frmUpdate.txtName.value == "")
				{
					window.parent.frames["UpperWindow"].frmUpdate.txtName.focus();
					alert("Program name is required.")
					blnOK = false;
				}
			else if (window.parent.frames["UpperWindow"].frmUpdate.cboProgramGroup.selectedIndex < 1)
				{
					window.parent.frames["UpperWindow"].frmUpdate.cboProgramGroup.focus();
					alert("Cycle Type is required.")
					blnOK = false;
				}

			if (blnOK)
				{				
				cmdCancel.disabled =true;
				cmdOK.disabled =true;
				window.parent.frames["UpperWindow"].frmUpdate.submit();
				}
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplus")%>')"  ></TD>
</TR></table>
</body>
</html>
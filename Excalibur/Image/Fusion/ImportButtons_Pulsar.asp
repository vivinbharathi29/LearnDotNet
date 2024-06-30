<%@ Language=VBScript %>
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
//	if (window.parent.frames["UpperWindow"].frmChange.txtSummary.value == "")
//		{
//			alert("You must specify a summary of your image strategy");
//			window.parent.frames["UpperWindow"].frmChange.txtSummary.focus();
//			blnSuccess = false;
//		}
	
	return blnSuccess;
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    }
    else {
        parent.window.parent.ClosePropertiesDialog();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
	{
	    var intImageExists = 0, strImage = "";
	    var response = false;   
	    for (i = 0; i < window.parent.frames["UpperWindow"].frmImport.chkSelected.length; i++) {
	        if (window.parent.frames["UpperWindow"].frmImport.chkSelected[i].checked)
	        {
	            if (window.parent.frames["UpperWindow"].frmImport.chkSelected[i].value.split('-')[2] == "1")
	            {
	                intImageExists = 1;
	                break;
	            }
	        }    
	    }
	    
	    if (intImageExists == 1)
	        response = confirm("Some Operating System(s) selected for importing already exists in the image definition.  Do you want to import the Localizations for the selected Operating System(s)?");
	    if (response)
	        window.parent.frames["UpperWindow"].frmImport.txtImportLocalizations.value = "1";
	    else
	        window.parent.frames["UpperWindow"].frmImport.txtImportLocalizations.value = "0";
	    cmdCancel.disabled = true;
		cmdOK.disabled =true;
		window.parent.frames["UpperWindow"].frmImport.submit();
	}
}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" align="right">
	<tr>
        <TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>	
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"  ></TD>        	
</TR></table>
</body>
</html>
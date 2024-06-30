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
		//if (window.parent.frames["UpperWindow"].frmChange.txtSummary.value == "")
	//	{
	//		alert("You must specify a summary of your image strategy");
	//		window.parent.frames["UpperWindow"].frmChange.txtSummary.focus();
	//		blnSuccess = false;
	//	}
	if (!window.parent.frames["UpperWindow"].frmUpdate.txtVendorName.value)
	{
		window.alert("Vendor Name is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtVendorName;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtDriverCategory.value)
	{
		window.alert("Driver Category is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtDriverCategory;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtOperatingSystem.value)
	{
		window.alert("Operating System is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtOperatingSystem;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtVersionPass.value)
	{
		window.alert("Version/Pass is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtVersionPass;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtLinktoDriver.value)
	{
		window.alert("Link to Driver is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtLinktoDriver;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtFileName.value)
	{
		window.alert("File Name is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtFileName;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtPlatform.value)
	{
		window.alert("Platform(s) is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtPlatform;
		blnSuccess = false;					
		return blnSuccess;
	}
	else if (!window.parent.frames["UpperWindow"].frmUpdate.txtDateNeeded.value)
	{
		window.alert("Date Needed is Required.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtDateNeeded;
		blnSuccess = false;					
		return blnSuccess;
	}
	
	
	return blnSuccess;
}

function cmdCancel_onclick() {
		window.parent.close();
}

function cmdOK_onclick() {
	var blnAll = true;
	var i;
	if (VerifySave())
	{
		cmdCancel.disabled =true;
		cmdOK.disabled =true;
		window.parent.frames["UpperWindow"].frmUpdate.submit();
	}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Send Email" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>
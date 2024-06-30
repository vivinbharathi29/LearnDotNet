<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../../includes/Date.asp" -->

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

	if (window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate.value && ! isDate(window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate.value))
		{
		window.alert("Target Date is invalid.");
		FocusOn = window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate;
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
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>
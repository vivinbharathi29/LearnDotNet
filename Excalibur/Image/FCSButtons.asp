<%@ Language=VBScript %>
<html>
<head>
<title></title>
<meta name="VI60_DefaultClientScript" content="JavaScript" />
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
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

function cmdCancel_onclick() {
    if ('<%=Request("isFromPulsarPlus")%>' != '')
        parent.window.parent.closeExternalPopup();
    else
		window.parent.close();
}

function cmdOK_onclick() {
	var blnAll = true;
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmUpdate.submit();
		}

}


//-->
</script>
</head>
<body bgcolor="ivory">
    <table border="0"  cellspacing="1"  cellpadding="1" align="right">
	    <tr>
		    <td><input type="button" value="OK" id="cmdOK" name="cmdOK" onclick="return cmdOK_onclick()" /></td>
		    <td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" onclick="return cmdCancel_onclick()"  /></td>
        </tr>
    </table>
</body>
</html>
<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>

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

function cmdCancel_onclick() {
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.close();
    }
}

function cmdOK_onclick() {
	var blnAll = true;
	var i;
	if (VerifySave())
	{
		//for (i=0;i<window.parent.frames["UpperWindow"].frmUpdate.cboProductStatus.length;i++)
		//{
			//if (window.parent.frames["UpperWindow"].frmUpdate.cboProductStatus[i].options[window.parent.frames["UpperWindow"].frmUpdate.cboProductStatus[i].selectedIndex].value == 0)
				//alert("Recommended");
			//else
				//alert("Critical");
		//}
		
		//alert(window.parent.frames["UpperWindow"].frmUpdate.txtProductIDs.value);
		//alert(window.parent.frames["UpperWindow"].frmUpdate.cboVersion.value);
			
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
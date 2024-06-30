<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../includes/Date.asp" -->

function isNumeric(sText)
{
   var ValidChars = "0123456789";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }
function VerifySave(){
	blnSuccess = true;
    if (window.parent.frames["UpperWindow"].frmUpdate.txtRev.value == "" )
		{
			blnSuccess=false;
			alert("Internal Rev is required.");
			window.parent.frames["UpperWindow"].frmUpdate.txtRev.focus();
		}
	else if(! isNumeric(window.parent.frames["UpperWindow"].frmUpdate.txtRev.value))
		{
			blnSuccess=false;
			alert("Internal Rev must be a number.");
			window.parent.frames["UpperWindow"].frmUpdate.txtRev.focus();
		}
	else if(window.parent.frames["UpperWindow"].frmUpdate.txtPNRev.value == "")
		{
			blnSuccess=false;
			alert("Part Number Rev is required.");
			window.parent.frames["UpperWindow"].frmUpdate.txtPNRev.focus();
		}
	else if(! isNumeric(window.parent.frames["UpperWindow"].frmUpdate.txtPNRev.value))
		{
			blnSuccess=false;
			alert("Part Number Rev must be a number.");
			window.parent.frames["UpperWindow"].frmUpdate.txtPNRev.focus();
		}
	return blnSuccess;
}

function cmdCancel_onclick() {
		window.parent.close();
}

function cmdOK_onclick() {
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
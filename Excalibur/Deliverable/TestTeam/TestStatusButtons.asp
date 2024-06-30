<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	if( (window.parent.frames["UpperWindow"].frmMain.cboStatus.selectedIndex == 2 || window.parent.frames["UpperWindow"].frmMain.cboStatus.selectedIndex == 3) && window.parent.frames["UpperWindow"].frmMain.txtNotes.value=="" )
		{
		blnSuccess=false;
		alert("Test Notes are required for the selected status.");
		window.parent.frames["UpperWindow"].frmMain.txtNotes.focus();
		}
	else if( window.parent.frames["UpperWindow"].frmMain.cboStatus.selectedIndex == 4 && window.parent.frames["UpperWindow"].frmMain.txtNotes.value.toLowerCase().indexOf("waive") > -1 )
		{
		blnSuccess=false;
		alert("You can not use the word 'waive' in Watch notes.");
		window.parent.frames["UpperWindow"].frmMain.txtNotes.focus();
		}
	else if( ! (window.parent.frames["UpperWindow"].frmMain.txtReceived.value=="" || IsNumeric(window.parent.frames["UpperWindow"].frmMain.txtReceived.value)))
		{
		blnSuccess=false;
		alert("The Total Received field must be a number if it is supplied.");
		window.parent.frames["UpperWindow"].frmMain.txtReceived.focus();
		}
	
	return blnSuccess;
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else
     window.parent.Close();
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled = true;
			window.parent.frames["UpperWindow"].frmMain.submit();
		}
}

function IsNumeric(sText)
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

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD style=display:><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"  ></TD>
</TR></table>
</body>
</html>
<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	var blnFound = false;
	

	//SE Validations
	if (window.parent.frames["UpperWindow"].frmMain.txtSETestLead.value=="true")
		{
		if( (window.parent.frames["UpperWindow"].frmMain.cboIntegrationStatus.selectedIndex == 2 || window.parent.frames["UpperWindow"].frmMain.cboIntegrationStatus.selectedIndex == 3) && window.parent.frames["UpperWindow"].frmMain.txtIntegrationNotes.value=="" )
			{
			blnSuccess=false;
			alert("SE Test Notes are required for the selected status.");
			window.parent.frames["UpperWindow"].frmMain.txtIntegrationNotes.focus();
			}
		else if ( window.parent.frames["UpperWindow"].frmMain.cboIntegrationStatus.selectedIndex == 0  && window.parent.frames["UpperWindow"].frmMain.txtIntegrationNotes.value!="" )
			{
			blnSuccess=false;
			alert("You must select an SE Test Status if you enter SE Test Notes.");
			window.parent.frames["UpperWindow"].frmMain.cboIntegrationStatus.focus();
			}
		}

	//ODM Validations
	if (blnSuccess && window.parent.frames["UpperWindow"].frmMain.txtODMTestLead.value=="true")
		{
		if( (window.parent.frames["UpperWindow"].frmMain.cboODMStatus.selectedIndex == 2 || window.parent.frames["UpperWindow"].frmMain.cboODMStatus.selectedIndex == 3) && window.parent.frames["UpperWindow"].frmMain.txtODMNotes.value=="" )
			{
			blnSuccess=false;
			alert("ODM Test Notes are required for the selected status.");
			window.parent.frames["UpperWindow"].frmMain.txtODMNotes.focus();
			}
		else if ( window.parent.frames["UpperWindow"].frmMain.cboODMStatus.selectedIndex == 0  && window.parent.frames["UpperWindow"].frmMain.txtODMNotes.value!="" )
			{
			blnSuccess=false;
			alert("You must select an ODM Test Status if you enter ODM Test Notes.");
			window.parent.frames["UpperWindow"].frmMain.cboODMStatus.focus();
			}
		}

	//WWAN Validations
	if (blnSuccess && window.parent.frames["UpperWindow"].frmMain.txtWWANTestLead.value=="true")
		{
		if( (window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.selectedIndex == 2 || window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.selectedIndex == 3) && window.parent.frames["UpperWindow"].frmMain.txtWWANNotes.value=="" )
			{
			blnSuccess=false;
			alert("WWAN Test Notes are required for the selected status.");
			window.parent.frames["UpperWindow"].frmMain.txtWWANNotes.focus();
			}
		else if( window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.selectedIndex == 4 && window.parent.frames["UpperWindow"].frmMain.txtWWANNotes.value.toLowerCase().indexOf("waive") > -1 )
			{
			blnSuccess=false;
			alert("You can not use the word 'waive' in Watch notes.");
			window.parent.frames["UpperWindow"].frmMain.txtWWANNotes.focus();
			}
		else if ( window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.selectedIndex == 0  && window.parent.frames["UpperWindow"].frmMain.txtWWANNotes.value!="" )
			{
			blnSuccess=false;
			alert("You must select an WWAN Test Status if you enter WWAN Test Notes.");
			window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.focus();
			}
		}
	
   //DEV Validations
	if (blnSuccess && window.parent.frames["UpperWindow"].frmMain.txtDEVTestLead.value=="true")
		{
		if( (window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.selectedIndex == 2 || window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.selectedIndex == 3) && window.parent.frames["UpperWindow"].frmMain.txtDEVNotes.value=="" )
			{
			blnSuccess=false;
			alert("DEV Test Notes are required for the selected status.");
			window.parent.frames["UpperWindow"].frmMain.txtDEVNotes.focus();
			}
		else if( window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.selectedIndex == 4 && window.parent.frames["UpperWindow"].frmMain.txtDEVNotes.value.toLowerCase().indexOf("waive") > -1 )
			{
			blnSuccess=false;
			alert("You can not use the word 'waive' in Watch notes.");
			window.parent.frames["UpperWindow"].frmMain.txtDEVNotes.focus();
			}
		else if ( window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.selectedIndex == 0  && window.parent.frames["UpperWindow"].frmMain.txtDEVNotes.value!="" )
			{
			blnSuccess=false;
			alert("You must select an DEV Test Status if you enter DEV Test Notes.");
			window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.focus();
			}
		}

	//Make sure at least one status is checked
	if (blnSuccess && window.parent.frames["UpperWindow"].frmMain.cboIntegrationStatus.selectedIndex == 0 && window.parent.frames["UpperWindow"].frmMain.cboODMStatus.selectedIndex == 0 && window.parent.frames["UpperWindow"].frmMain.cboWWANStatus.selectedIndex == 0 && window.parent.frames["UpperWindow"].frmMain.cboDEVStatus.selectedIndex == 0)
		{
		blnSuccess=false;
		alert("You must enter a status to continue.");
		}
	


	//Make sure at least one deliverable is checked.
	
	
	if (blnSuccess)
    {
		if (typeof(window.parent.frames["UpperWindow"].frmMain.lstID.length)!="undefined")
        {
			for (i=0;i<window.parent.frames["UpperWindow"].frmMain.lstID.length;i++)
				if (window.parent.frames["UpperWindow"].frmMain.lstID[i].checked)
					blnFound=true;
		}
		else
		{
			if (window.parent.frames["UpperWindow"].frmMain.lstID.checked)
				blnFound=true;
		}	
		
		if (! blnFound)
			{
			blnSuccess=false;
			alert("You must select at least one deliverable to continue.");
			}
	
		}
	
	return blnSuccess;
}

function cmdCancel_onclick(pulsarplusDivId) {
    //window.parent.close();
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.Cancel();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
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
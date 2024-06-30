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
    var blnFound = false;
    var i;

    var strRequired = window.parent.frames["UpperWindow"].document.getElementById("txtDateRequired").value.indexOf("," + window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value + ",");
    var strCommentsRequired = window.parent.frames["UpperWindow"].document.getElementById("txtCommentsRequired").value.indexOf("," + window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value + ",");
	
    //Clear out invalid formatted dates if Scheduled is not the selected status
    if ( (! isDate(window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value)) && strRequired == -1)
        window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value="";
	
    if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value != "" && strRequired != -1 && (! isDate(window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value)))
    {
        alert("You must supply a valid date if one is entered.");
        window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].document.getElementById("txtComments").value == "" && strCommentsRequired != -1 )
    {
        alert("You must supply comments when entering this Pilot status.");
        window.parent.frames["UpperWindow"].document.getElementById("txtComments").focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value == "" && strRequired != -1)
    {
        alert("You must supply a valid date.");
        window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").focus();
        blnSuccess = false;
    }
    else if (window.parent.frames["UpperWindow"].document.getElementById("txtQualStatus").value != "QComplete" && window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value =="6")
    {
        alert("You can not set the pilot run status to \"Complete\" until the test status is set to \"QComplete\".");
        window.parent.frames["UpperWindow"].document.getElementById("cboStatus").focus();
        blnSuccess = false;
    }
	
    if (window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].value =="2")
        window.parent.frames["UpperWindow"].document.getElementById("txtStatusText").value = window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].text + "|" + window.parent.frames["UpperWindow"].document.getElementById("txtPilotDate").value;
    else
        window.parent.frames["UpperWindow"].document.getElementById("txtStatusText").value = window.parent.frames["UpperWindow"].document.getElementById("cboStatus").options[window.parent.frames["UpperWindow"].document.getElementById("cboStatus").selectedIndex].text;

    return blnSuccess;
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else
     window.parent.Cancel();
}

function cmdOK_onclick(keepItOpen) {
    var blnAll = true;
    var i;
    if (VerifySave())
    {
        cmdSave.disabled = true;
        cmdCancel.disabled =true;
        cmdOK.disabled =true;
        window.parent.frames["UpperWindow"].frmStatus.txtKeepItOpen.value = keepItOpen;
        window.parent.frames["UpperWindow"].frmStatus.submit();
    }

}

function enableButton()
{
    cmdCancel.disabled =false;
    cmdOK.disabled = false;
    cmdSave.disabled = false;
}
//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
        <TD><INPUT type="button" value="Save & Close" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick(false)"></TD>
        <TD><INPUT type="button" value="Save" id=cmdSave name=cmdSave LANGUAGE=javascript onclick="return cmdOK_onclick(true)"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></TD>
</TR></table>
</body>
</html>

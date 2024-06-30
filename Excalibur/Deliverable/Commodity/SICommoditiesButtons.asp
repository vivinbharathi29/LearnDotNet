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


    function VerifySave() {
        var blnSuccess = true;
		/*if (window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value != "" && ! isDate(window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value))
		{
			alert("You must supply a valid date if one is entered.");
			window.parent.frames["UpperWindow"].frmStatus.txtTestDate.focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value == "" && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value =="3")
		{
			alert("You must supply a valid date.");
			window.parent.frames["UpperWindow"].frmStatus.txtTestDate.focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmStatus.cboWhy.selectedIndex >2 && window.parent.frames["UpperWindow"].frmStatus.cboDCR.selectedIndex==0)
		{
			alert("You must select an approved DCR number.");
			window.parent.frames["UpperWindow"].frmStatus.cboDCR.focus();
			blnSuccess = false;
		}
		
		if (window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value =="3")
			window.parent.frames["UpperWindow"].frmStatus.txtStatusText.value =	window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value;
		else
			window.parent.frames["UpperWindow"].frmStatus.txtStatusText.value =	window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].text;
	*/
        return blnSuccess;
    }

    function cmdCancel_onclick() {
        var pulsarplusDivId = document.getElementById("pulsarplusDivId");
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            window.parent.close();
        }
    }

    function cmdOK_onclick() {
        var i;
        if (VerifySave()) {
            //cmdCancel.disabled =true;
            cmdOK.disabled = true;
            window.parent.frames["UpperWindow"].frmMain.submit();
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
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</body>
</html>
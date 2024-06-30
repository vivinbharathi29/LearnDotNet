<%@ Language=VBScript %>

<!-- #include file = "../../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
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


    if (window.parent.frames["UpperWindow"].frmStatus.txtStatusLoaded.value!=window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value == "5")
        {
        if (window.parent.frames["UpperWindow"].frmStatus.txtTestingComplete.value=="0")
            {
             if (! window.confirm("This deliverable has not completed all required testing. Are you sure you want to set it to " + window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].text + "?"))
                return false;
            }
        else if (window.parent.frames["UpperWindow"].frmStatus.txtTestingComplete.value=="2")
            {
             if (! window.confirm("This deliverable TTS status is still Pending. Are you sure you want to set it to " + window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].text + "?"))
                return false;
            }
        else if (window.parent.frames["UpperWindow"].frmStatus.txtTestingComplete.value=="3")
            {
             if (! window.confirm("This deliverable has not completed all required testing and the TTS status is still Pending. Are you sure you want to set it to " + window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].text + "?"))
                return false;
            }
        }
        
	var strCommentsRequired = window.parent.frames["UpperWindow"].txtCommentsRequired.value.indexOf("," + window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value + ",");
	
	//Clear out invalid formatted dates if Date is not the selected status
	if ( (! isDate(window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value)) && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value != "3")
		window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value="";
	
	if (typeof(window.parent.frames["UpperWindow"].frmStatus.lstSub.length) == "undefined")
		{	
		if(window.parent.frames["UpperWindow"].frmStatus.lstSub.checked)
			blnFound=true;
		}
	else
		{
		for (i=0;i<window.parent.frames["UpperWindow"].frmStatus.lstSub.length;i++)
			if(window.parent.frames["UpperWindow"].frmStatus.lstSub[i].checked)
				{
				blnFound=true;
				break;
				}
		}
	if (!blnFound && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value !="0")		
		{
			alert("You must select at least one subassembly.");
			//window.parent.frames["UpperWindow"].frmStatus.lstSub[0].focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmStatus.txtComments.value == "" && strCommentsRequired != -1 )
		{
			alert("You must supply comments when entering this Qualification status.");
			window.parent.frames["UpperWindow"].frmStatus.txtComments.focus();
			blnSuccess = false;
		}
		
		else if (window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value != "" && window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value == "3" && (! isDate(window.parent.frames["UpperWindow"].frmStatus.txtTestDate.value)))
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
		else if (window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].value =="5" && window.parent.frames["UpperWindow"].frmStatus.chkRiskRelease.checked)
			window.parent.frames["UpperWindow"].frmStatus.txtStatusText.value =	"Risk Release";
		else
			window.parent.frames["UpperWindow"].frmStatus.txtStatusText.value =	window.parent.frames["UpperWindow"].frmStatus.cboStatus.options[window.parent.frames["UpperWindow"].frmStatus.cboStatus.selectedIndex].text;
	return blnSuccess;
}

function cmdCancel_onclick(pulsarplusDivId) {
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
	var blnAll = true;
	var i;
	if (VerifySave())
		{

			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmStatus.submit();
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"  ></TD>
</TR></table>
</body>
</html>
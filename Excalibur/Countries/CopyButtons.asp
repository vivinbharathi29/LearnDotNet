<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	if (window.parent.frames["UpperWindow"].frmMain.ddlProducts)
	{	
		if (window.parent.frames["UpperWindow"].frmMain.ddlProducts.value == 0 || window.parent.frames["UpperWindow"].frmMain.ddlProducts.value == '')
		{
			alert('Please Select a Product Version');
			return false;
		}
	}
	else
		return false;

	if (window.parent.frames["UpperWindow"].frmMain.ddlBrand)
	{	
		if (window.parent.frames["UpperWindow"].frmMain.ddlBrand.value == 0 || window.parent.frames["UpperWindow"].frmMain.ddlBrand.value == '')
		{
			alert('Please Select a Brand');
			return false;
		}
	}
	else
	    return false;

	var bReleaseChecked = false;
	if (window.parent.frames["UpperWindow"].frmMain.chkCopyReleases) {
	    if (typeof (window.parent.frames["UpperWindow"].frmMain.chkCopyReleases.length) == "undefined") {
	        if (!window.parent.frames["UpperWindow"].frmMain.chkCopyReleases.checked == true) {
	            alert('Please Select a Release');
	            return false;
	        }
	    }
	    else
	    {
            for (var i = 0; i < window.parent.frames["UpperWindow"].frmMain.chkCopyReleases.length; i++) {
	            if (window.parent.frames["UpperWindow"].frmMain.chkCopyReleases[i].checked) {
	                bReleaseChecked = true;
	                break;
	            }
	        }
	        if (bReleaseChecked == false) {
	            alert('Please Select a Release');
	            return false;
	        }
	    }	   
	}
	else
	    return false;

    return true;
}

function cmdCancel_onclick() {
    var pulsarplusDivId = document.getElementById('hdnTabName');
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
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


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
    <input type="hidden" id="hdnTabName" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>
<%@ Language=VBScript %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function VerifyStatus() {
        with (window.parent.frames["SCMUpperWindow"].frmMain) {
        }
        return true;
    }

    function cmdCancel_onclick() {
        var objReturn = new Object();
        objReturn.Refresh = "0";
        objReturn.CategoryRules = "";
        var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            //PBI 10633 - task 20275 - Convert the SCM Category Details popup to jQuery
            parent.window.parent.CloseSCMCategoryPropertiesDialog(objReturn);
        }
    }

    function cmdOK_onclick() {
        var blnAll = true;
        var i;
        var sReturnValue;

        if (VerifyStatus()) {
            window.frmSCMButtons.cmdSCMOK.disabled = true;
            window.frmSCMButtons.cmdSCMCancel.disabled = true;
            window.parent.frames["SCMUpperWindow"].frmMain.hidFunction.value = "save";
            window.parent.frames["SCMUpperWindow"].frmMain.submit();
        }

        return;
    }

    function document_OnLoad() {
        window.frmSCMButtons.cmdSCMOK.disabled = true;
        if (typeof (window.parent.frames["SCMUpperWindow"].document.all["hidMode"]) == 'object') {
            if (window.parent.frames["SCMUpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add' || window.parent.frames["SCMUpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit')
                window.frmSCMButtons.cmdSCMOK.disabled = false;
        }
    }

    //-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmSCMButtons"  action=SCMCatButtons.asp method=post>
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Save" id=cmdSCMOK name=cmdSCMOK onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdSCMCancel name=cmdSCMCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>
</FORM>
    <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>
<%@ Language=VBScript %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsMarketingUser
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager

Dim AppRoot
AppRoot = Session("ApplicationRoot")

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity

	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If

	Set Security = Nothing
'##############################################################################	


If Request("AVID") <> Request("hidAVID") And Request("hidAVID") <> "" Then
	Response.Redirect "avButtons.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request.Form("hidAVID")
End If


Dim rs, dw, cn, cmd
Dim iPrev, iCur, iNext

If LCase(Request("MODE")) <> "add" Then
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
	dw.CreateParameter cmd, "@p_GpgDescription", adVarchar, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_UPC", adChar, adParamInput, 12, ""
	dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
	dw.CreateParameter cmd, "@p_KMAT", adChar, adParamInput, 6, ""
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	If Trim(Request("AVID")) <> "" Then
		Do Until rs.EOF Or iCur = CLng(Trim(Request("AVID")))
			iPrev = iCur
			iCur = rs("AvDetailID")
			rs.MoveNext
		Loop

		If Not rs.EOF Then
			iNext = rs("AvDetailID")
		End If
	End If

	rs.Close
End If	

%>
<html>
<head>
<script language="JavaScript" src="../includes/client/Common.js"></script>

<% If m_IsMarketingUser Then %>
<SCRIPT type="text/javascript">
<!--
    function VerifyStatus() { return true; }
//-->
</SCRIPT>
<% Else %>
<SCRIPT type="text/javascript">
<!--
    function VerifyStatus() {
    with (window.parent.frames["UpperWindow"].frmMain) {
            //Check if Shared AV and Config Rules differ within KMAT//
            //if (ValidateConfigRule(frmButtons.AVID.value, frmButtons.BID.value, txtConfigRules.value) == 1) {
            //    var answer = confirm("Changes to this Shared AV's Configuration Rule will be made in all SCMs for the same KMAT where the Shared AV is used.");
            //    if (answer == false)
            //        return false;
            //}
        //alert(window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value);
        if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
            if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "True") {
                txtAvGpgDescription.value = txtAVName3.value
                txtMarketingDesc.value = txtAVName5.value
                txtMarketingDescPMG.value = txtAVName7.value
            }
        }
		//if (!validateTextInput(txtAvGpgDescription, 'Gpg Description')){ return false; }
            if (AvNo.value == "" && txtAvGpgDescription.value == "") {
			alert("Av Number or GPG Description is required.");
			return false;
        }
		
            if (txtAvGpgDescription.value == hidGpgDescription.value && hidMode.value == "clone") {
			alert("Gpg Description must be changed");
			return false;
        }

		if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
		    if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "False") {
		        if (!validateTextAreaSize(txtMarketingDesc, 40, 'Marketing Description')) { return false; }
		        if (!validateTextAreaSize(txtConfigRules, 1600, 'Config Rules')) { return false; }
		        if (!validateTextAreaSize(txtManufacturingNotes, 1600, 'Manufacturing Notes')) { return false; }
		    }
		}
		var iBrandCount = 0;
            if (hidAVID == "") {
                for (i = 0; i <= chkBrand.length - 1; i++) {
				if (chkBrand[i].checked)
					iBrandCount++;
			}
                if (iBrandCount == 0) {
				alert('Select At Least One Brand');
				return false;
			}
		}
		
            if (!validateDateInput(txtGSEndDt, "GS End Date")) { return false; }

            if (cboCategory.value == 0) {
			alert('Please choose a valid Feature Category');
			return false;
        }

        if ((hidParentID.value == 0 && AvNo.value != "") || (hidAVID.value == "" && AvNo.value != "")) {
            //check if NON-Localized AV contains a # char
            var index = AvNo.value.indexOf("#");

            if (index > -1 && hidAVID.value == "") {
                alert("You are creating new NON-Localized AV.  Please remove the # character and any characters after the #.");
                return false;
            }
            else if (index > -1) {
                alert("This AV is a NON-Localized AV.  Please remove the # character and any characters after the #.");
                return false;
            }
        }
	}
	return true;
}

    function ValidateConfigRule(AVID, BID, ConfigRules) {
        var parameters = "function=ValidateConfigRule&AVID=" + AVID + "&BID=" + BID + "&ConfigRules=" + encodeURIComponent(ConfigRules);
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {        //Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "<%=AppRoot %>/SCM/ValidateSharedAVConfigRule.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
        if (request.responseText == 'PASS') {
            return 0;
        } else {
            return 1;
        }
    }
//-->
</SCRIPT>
<% End If %>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function cmdCancel_onclick() {
	window.parent.close();
}

    function cmdOK_onclick() {
	var blnAll = true;
	var i;
	var sReturnValue;

	if (VerifyStatus()) {
	    if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
	        if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "True") {
	            window.parent.frames["UpperWindow"].SaveNameElements();
	        } else {
	            window.parent.frames["UpperWindow"].frmMain.strNameElements.value = "";
	        }
	    }
            window.frmButtons.cmdOK.disabled = true;
		window.frmButtons.cmdCancel.disabled = true;

		//alert("GPG Desc: " + window.parent.frames["UpperWindow"].frmMain.txtAvGpgDescription.value);
		//alert("Marketing Desc: " + window.parent.frames["UpperWindow"].frmMain.txtMarketingDesc.value);
		//alert("Marketing Desc PMG: " + window.parent.frames["UpperWindow"].frmMain.txtMarketingDescPMG.value);
	        
            window.parent.frames["UpperWindow"].frmMain.hidFunction.value = "save";
		window.parent.frames["UpperWindow"].frmMain.submit();
	}
	
	return;
}

    function cmdPrev_onclick() {
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidPrev.value;
	frmButtons.hidAVID.value = frmButtons.hidPrev.value;
	cmdOK_onclick();
	frmButtons.submit();
}

    function cmdNext_onclick() {
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidNext.value;
	frmButtons.hidAVID.value = frmButtons.hidNext.value;
	cmdOK_onclick();
	frmButtons.submit();
}

    function cmdClone_onclick() {
	var strID;
	var productVersionID = frmButtons.PVID.value;
	var avDetailID = frmButtons.AVID.value;
	var productBrandID = frmButtons.BID.value;
        strID = window.parent.showModalDialog("avFrame.asp?Mode=clone&PVID=" + productVersionID + "&AVID=" + avDetailID + "&BID=" + productBrandID, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
	document.location.reload();
}
    function document_OnLoad() {
	window.frmButtons.cmdOK.disabled = true;
        if (typeof (window.parent.frames["UpperWindow"].document.all["hidMode"]) == 'object') {
            if (window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add' || window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit' || window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'clone')
			window.frmButtons.cmdOK.disabled = false;
	}
	
        if (typeof (window.frmButtons.cmdNext) == 'object') {
		if (window.frmButtons.hidNext.value == '')
			window.frmButtons.cmdNext.disabled = true;
	}	
        if (typeof (window.frmButtons.cmdPrev) == 'object') {
		if (window.frmButtons.hidPrev.value == '')
			window.frmButtons.cmdPrev.disabled = true;
	}
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmButtons"  action=avButtons.asp method=post>
<INPUT type="HIDDEN" id=PVID name=PVID value="<%= Request("PVID")%>">
<INPUT type="HIDDEN" id=AVID name=AVID value="<%= Request("AVID")%>">
<INPUT type="HIDDEN" id=BID name=BID value="<%= Request("BID")%>">
<INPUT type="HIDDEN" id=Mode name=Mode value="<%= Request("Mode")%>">
<INPUT type="HIDDEN" id=hidPrev name=hidPrev value="<%= iPrev%>">
<INPUT type="HIDDEN" id=hidNext name=hidNext value="<%= iNext%>">
<INPUT type="HIDDEN" id=hidAVID name=hidAVID value="<%= Request("AVID")%>">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" WIDTH=100%>
	<tr><TD width=33% align=left>
<% If LCase(Request("MODE")) <> "clone" And  m_EditModeOn Then %>			
			<INPUT type="button" value="Clone" id=cmdClone name=cmdClone onclick="return cmdClone_onclick()"></TD>
<% Else %>
			&nbsp;
<% End If %>
		<TD width=33% align=center>
<% If LCase(Request("MODE")) <> "add" And LCase(Request("MODE")) <> "clone" Then %>			
			<INPUT type="button" accesskey="P" value="Prev" id=cmdPrev name=cmdPrev onclick="return cmdPrev_onclick()">
			<INPUT type="button" accesskey="N" value="Next" id=cmdNext name=cmdNext onclick="return cmdNext_onclick()">
<% Else %>
			&nbsp;
<% End If %>
		</TD>
		<TD width=33% align=right><INPUT accesskey="S" type="button" value="Save" id=cmdOK name=cmdOK onclick="return cmdOK_onclick()">
			<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>
</FORM>
</body>
</html>
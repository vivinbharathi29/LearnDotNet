<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file = "../includes/lib_debug.inc" -->
<%
Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sMode				    : sMode = Request.QueryString("Mode")
Dim sFunction			    : sFunction = Request.Form("hidFunction")
Dim sSCMCategory			: sSCMCategory = ""
Dim sManufacturingNotes	    : sManufacturingNotes = ""
Dim sMarketingDescription   : sMarketingDescription = ""
Dim sConfigRules		    : sConfigRules = ""
Dim iBrandID			    : iBrandID = Request.QueryString("BID")
Dim sCatMin					: sCatMin = ""
Dim sCatMax					: sCatMax = ""
Dim sIsDesktop  
Dim sRulesSyntax			: sRulesSyntax = ""

Dim m_ProductVersionID	    : m_ProductVersionID = Request("PVID")
Dim m_SCMCategoryID         : m_SCMCategoryID = Request("SCMID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName


'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	
'
'PC Can do any thing
'Marketing can change description

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Sub Main()
'
'TODO: Get CatDetail Data
'
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	sIsDesktop = rs("IsDesktop") 
	rs.Close

		Set cmd = dw.CreateCommandSP(cn, "usp_SCM_SelectSCMCategoryDetail")
		dw.CreateParameter cmd, "@p_intProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_intSCMCategoryID", adInteger, adParamInput, 8, Trim(Request("SCMID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		If Not rs.EOF Then
			sSCMCategory = rs("SCMCategory") & ""
			sConfigRules = rs("ConfigRules") & ""
			sManufacturingNotes = rs("ManufacturingNotes") & ""
			sMarketingDescription = rs("MarketingDescription") & ""
			sCatMin = rs("CatMin") & ""
			sCatMax = rs("CatMax") & ""
			sRulesSyntax = rs("RulesSyntax") & ""
		End If
		
		rs.Close
		
		If isnull(sCatMin) then sCatMin = ""
		If isnull(sCatMax) then sCatMax = ""
End Sub

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	Dim iAvId

	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	sIsDesktop = rs("IsDesktop") 
	rs.Close

	if Request.Form("txtCatMin") = "" then sCatMin = -1 else sCatMin = Request.Form("txtCatMin")
	if Request.Form("txtCatMax") = "" then sCatMax = -1 else sCatMax = Request.Form("txtCatMax")

	If Request.Form("txtConfigRules") <> Request.Form("hidConfigRulesDefault") _
		Or Request.Form("txtManufacturingNotes") <> Request.Form("hidManufacturingNotesDefault") _
		Or Request.Form("txtMarketingDescription") <> Request.Form("hidMarketingDescriptionDefault") _
		Or sCatMin <> Request.Form("hidCatMin") _
		Or sCatMax <> Request.Form("hidCatMax") _
		Or Request.Form("txtRulesSyntax") <> Request.Form("hidRulesSyntax") Then
	
		if sCatMin = -1 then sCatMin = null
		if sCatMax = -1 then sCatMax = null

		cn.BeginTrans
		'Save CategoryDetail data
		Set cmd = dw.CreateCommandSP(cn, "usp_SCM_UpdateSCMCategoryDetail")
		cmd.NamedParameters = True
		dw.CreateParameter cmd, "@p_intProductBrandID", adInteger, adParamInput, 8, Request("BID")
		dw.CreateParameter cmd, "@p_intSCMCategoryID", adInteger, adParamInput, 8, Request("SCMID")
		dw.CreateParameter cmd, "@p_chrConfigRules", adVarchar, adParamInput, 2000, Request.Form("txtConfigRules")
		dw.CreateParameter cmd, "@p_chrManufacturingNotes", adVarchar, adParamInput, 2000, Request.Form("txtManufacturingNotes")
		dw.CreateParameter cmd, "@p_chrMarketingDescription", adVarchar, adParamInput, 2000, Request.Form("txtMarketingDescription")
		dw.CreateParameter cmd, "@p_chrLastUpdUser", adVarchar, adParamInput, 50, m_UserFullName
		dw.CreateParameter cmd, "@p_intCatMin", adInteger, adParamInput, 8, sCatMin
		dw.CreateParameter cmd, "@p_intCatMax", adInteger, adParamInput, 8, sCatMax
	    dw.CreateParameter cmd, "@p_chrRetMsg",adVarchar, adParamOutput, 256,""
		'if sIsDesktop = True then
			dw.CreateParameter cmd, "@p_RulesSyntax", adVarchar, adParamInput, 512, Request.Form("txtRulesSyntax") 
		'else
		'	dw.CreateParameter cmd, "@p_RulesSyntax", adVarchar, adParamInput, 512, ""
		'end if
		returnValue = dw.ExecuteNonQuery(cmd)

		If cmd("@p_ChrRetMsg") <> "" Then
			Response.Write cmd("@p_ChrRetMsg")
			Response.End
		End If

		sFunction = "close"
		cn.CommitTrans
	End If
End Sub

If LCase(sFunction) = "save" Then
	Call Save()
    Call Main()
Else
	Call Main()
End If
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css">
    <script type="text/javascript">
        function Body_OnLoad() {
            var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
            switch (frmMain.hidFunction.value) {

                case "close":
                    var objReturn = new Object();
                    objReturn.Refresh = "1";
                    objReturn.CategoryMD = frmMain.txtMarketingDescription.value;
                    objReturn.CategoryRules = frmMain.txtConfigRules.value;
                    objReturn.CategoryRuleSyntax = frmMain.txtRulesSyntax.value;
                    objReturn.CatMin = frmMain.txtCatMin.value;
                    objReturn.CatMax = frmMain.txtCatMax.value;
                    objReturn.SCMCategoryID = frmMain.hidFCID.value; //pass SCMCategoryID to refersh the grid - bug 22270
                    //PBI 10633 - task 20275 - Convert the SCM Category Details popup to jQuery
                    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                        parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                        // For Closing current popup if Called from pulsarplus
                        parent.window.parent.closeExternalPopup();
                    }
                    else {
                        parent.window.parent.CloseSCMCategoryPropertiesDialog(objReturn);
                    }
                    break;
                case "save":
                    //PBI 10633 - task 20275 - Convert the SCM Category Details popup to jQuery
                    var objReturn = new Object();
                    objReturn.Refresh = "1";

                    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                        parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                        // For Closing current popup if Called from pulsarplus
                        parent.window.parent.closeExternalPopup();
                    }
                    else {
                        parent.window.parent.CloseSCMCategoryPropertiesDialog(objReturn);
                    }
                    break;
            }
            if (typeof (window.parent.frames["SCMLowerWindow"].frmSCMButtons) == 'object') {
                if (window.frmMain.hidMode.value.toLowerCase() == 'edit' || window.frmMain.hidMode.value.toLowerCase() == 'add')
                    window.parent.frames["SCMLowerWindow"].frmSCMButtons.cmdSCMOK.disabled = false;
            }

        }

        function checkInteger(e, minnumber, maxnumber) {
            // Get ASCII value of key that user pressed
            if (!e) e = window.event;
            var charCode = e.keyCode ? e.keyCode : e.which;
            // charCode of 48 = 0, 57 = 9
            if ((charCode - 48 >= minnumber && charCode - 48 <= maxnumber) || charCode == 8)	// numbers or backspace
                return;
            else // otherwise, discard character 
                if (window.event)
                    e.returnValue = null; // IE
                else
                    e.preventDefault(); // Firefox
        }

        function textCounter(field, countfield, maxlimit) {
            if (field.value.length > maxlimit)
                field.value = field.value.substring(0, maxlimit);
            else {
                if (countfield != null)
                    countfield.innerHTML = maxlimit - field.value.length;
            }
        }

    </script>
</head>
<body onload="Body_OnLoad()">
    <%'= response.Write(request.QueryString) %>
    <form method="post" id="frmMain">
        <input id="hidMode" name="hidMode" type="hidden" value="<%= LCase(sMode)%>">
        <input id="hidFunction" name="hidFunction" type="hidden" value="<%= LCase(sFunction)%>">
        <input id="BID" name="BID" type="hidden" value="<%= iBrandID%>">
        <input id="hidFCID" name="hidFCID" type="hidden" value="<%=m_SCMCategoryID%>" />
        <input id="hidCatName" name="hidCatName" type="hidden" value="<%= sSCMCategory%>">
        <textarea rows="5" id="hidConfigRulesDefault" name="hidConfigRulesDefault" style="display: none;"><%= sConfigRules %></textarea>
        <textarea rows="5" id="hidManufacturingNotesDefault" name="hidManufacturingNotesDefault" style="display: none;"><%= sManufacturingNotes %></textarea>
        <textarea rows="5" id="hidMarketingDescriptionDefault" name="hidMarketingDescriptionDefault" style="display: none;"><%= sMarketingDescription %></textarea>
        <input id="hidCatMin" name="hidCatMin" type="hidden" value="<%= sCatMin %>">
        <input id="hidCatMax" name="hidCatMax" type="hidden" value="<%= sCatMax %>">
        <input id="hidRulesSyntax" name="hidRulesSyntax" type="hidden" value="<%= sRulesSyntax %>">
        <table class="FormTable" bgcolor="cornsilk" width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="tan">
            <tr>
                <th>SCM Category:</th>
                <td><%= PrepForWeb(sSCMCategory)%></td>
            </tr>
            <tr>
                <th>Configuration Rules:</th>
                <td>
                    <textarea rows="5" id="txtConfigRules" name="txtConfigRules" style="width: 300px"><%= sConfigRules%></textarea></td>
            </tr>
            <%'if sIsDesktop then %>
            <tr>
                <th>Rules Syntax:<br />
                    <br />
                    <br />
                    <span class="Label" style="font-weight: normal;">Remaining characters: </span>
                    <span class="LabelHeader"><span id="tLen2"><% if len(sRulesSyntax)>0 then response.write 512-len(sRulesSyntax) else response.write "512" %></span></span></th>
                <td>
                    <textarea rows="5" id="txtRulesSyntax" name="txtRulesSyntax" maxlength="512" style="width: 300px; text-transform: uppercase;"
                        onkeydown="textCounter(this.form.txtRulesSyntax, document.getElementById('tLen2'), 512);"
                        onkeyup="textCounter(this.form.txtRulesSyntax, document.getElementById('tLen2'), 512);"
                        onchange="this.value = this.value.toUpperCase();"><%= sRulesSyntax%></textarea>
                </td>
            </tr>
            <% 'end if %>
            <tr>
                <th>Min/Max:<br />
                    <span style="font-weight: normal;">If left blank, it will be blank in the SCM report</span>
                </th>
                <td>
                    <table>
                        <tr>
                            <td>
                                <input type="text" id="txtCatMin" name="txtCatMin" maxlength="1" style="width: 80px;" value="<%= sCatMin %>" onkeypress="return checkInteger(event, 0, 1)" /></td>
                            <td style="vertical-align: middle;">0 or 1</td>
                            <td style="width: 10px;"></td>
                            <td>
                                <input type="text" id="txtCatMax" name="txtCatMax" maxlength="3" style="width: 80px;" value="<%= sCatMax %>" onkeypress="return checkInteger(event, 0, 9)" /></td>
                            <td style="vertical-align: middle;">0 to 999</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <th>Manufacturing Notes:</th>
                <td>
                    <textarea rows="5" id="txtManufacturingNotes" name="txtManufacturingNotes" style="width: 300px"><%= sManufacturingNotes%></textarea></td>
            </tr>
            <tr>
                <th>Marketing Description:</th>
                <td>
                    <textarea rows="5" id="txtMarketingDescription" name="txtMarketingDescription" style="width: 300px"><%= sMarketingDescription %></textarea></td>
            </tr>
        </table>
    </form>
    <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>

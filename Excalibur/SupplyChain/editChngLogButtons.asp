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

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity

	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
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

<SCRIPT type="text/javascript">
<!--
function VerifyStatus() {return true;}
//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function cmdCancel_onclick() 
{
    var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        parent.window.parent.closeExternalPopup();
    }
    else {
        window.parent.close();
    }
}

function cmdOK_onclick() 
{
	var blnAll = true;
	var i;
	var sReturnValue;
	
	if (VerifyStatus())
	{
		window.frmButtons.cmdOK.disabled =true;
		window.frmButtons.cmdCancel.disabled =true;
		window.parent.frames["UpperWindow"].frmMain.hidFunction.value="save";
		window.parent.frames["UpperWindow"].frmMain.submit();
	}
	
	return;
}

function document_OnLoad()
{
	window.frmButtons.cmdOK.disabled = true;
	if (typeof(window.parent.frames["UpperWindow"].document.all["hidMode"]) == 'object')
	{
	    if ((window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add' || window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit' || window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'clone') && window.parent.frames["UpperWindow"].document.all["HasAccess"].value.toLowerCase() == 'true')
			window.frmButtons.cmdOK.disabled = false;
	}
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmButtons"  action=avButtons.asp method=post>
<INPUT type="HIDDEN" id=CLID name=PVID value="<%= Request("CLID")%>">
    <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" WIDTH=100%>
	<tr><TD width=33% align=left>
			&nbsp;</TD>
		<TD width=33% align=center>
			&nbsp;
		</TD>
		<TD width=33% align=right><INPUT type="button" value="Save" id=cmdOK name=cmdOK onclick="return cmdOK_onclick()">
			<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>
</FORM>
</body>
</html>
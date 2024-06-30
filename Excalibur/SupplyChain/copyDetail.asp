<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sMode				: sMode = Request.QueryString("Mode")
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim sFeatureCat			: sFeatureCat = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sConfigRules		: sConfigRules = ""
Dim iBrandID			: iBrandID = ""
Dim sAvList				: sAvList = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
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
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvFeatureDetail")
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Trim(Request("FCID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		If Not rs.EOF Then
			sFeatureCat = rs("AvFeatureCategory")
			sConfigRules = rs("ConfigRules")
			sManufacturingNotes = rs("ManufacturingNotes")
		End If
		
		rs.Close
		
End Sub

Sub Search()
	Dim sSearch : sSearch = CSTR(Request.Form("txtSearch"))
	Dim sReplace : sReplace = CSTR(Request.Form("txtReplace"))
	Dim sNewGpgDescription
	Dim sFeatureCat

	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
	dw.CreateParameter cmd, "@p_GpgDescription", adVarchar, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_UPC", adChar, adParamInput, 12, ""
	dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
	dw.CreateParameter cmd, "@p_KMAT", adChar, adParamInput, 6, Trim(Request.Form("txtKMAT"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	sAvList = ""
			
	Do Until rs.EOF
		If (instr(1, rs("GpgDescription"), sSearch, 1)) > 0 Then
			If sFeatureCat <> rs("AvFeatureCategory") Then
				sFeatureCat = rs("AvFeatureCategory")
				sAvList = sAvList & "<TR><TD ID=Feature ColSpan=2>" & sFeatureCat & "</TD></TR>"
			End If
			
			sNewGpgDescription = Replace(rs("GpgDescription")&"", sSearch, sReplace, 1, -1, vbTextCompare)
			sAvList = sAvList & "<TR><TD><INPUT type=checkbox id=cbx" & rs("AvDetailID") & " name=cbx" & rs("AvDetailID") & "  checked>" & _
								"<INPUT type=hidden id=AvDetailID name=AvDetailID value='" & rs("AvDetailID") & "'>" & _
								"<INPUT type=hidden id=FeatureCat" & rs("AvDetailID") & " name=FeatureCat" & rs("AvDetailID") & "  value='" & rs("FeatureCategoryID") & "'>" & _
								"<INPUT type=hidden id=GPGDescription" & rs("AvDetailID") & " name=GPGDescription" & rs("AvDetailID") & "  value='" & sNewGpgDescription & "'>" & _
								"<INPUT type=hidden id=MktgDescription" & rs("AvDetailID") & " name=MktgDescription" & rs("AvDetailID") & "  value='" & rs("MarketingDescription") & "'>" & _
								"<TEXTAREA style='display:none;' id=ConfigRules" & rs("AvDetailID") & " name=ConfigRules" & rs("AvDetailID") & ">" & rs("ConfigRules") & "</TEXTAREA>" & _
								"<INPUT type=hidden id=IdsSkus" & rs("AvDetailID") & " name=IdsSkus" & rs("AvDetailID") & "  value='" & rs("IdsSkus_YN") & "'>" & _
								"<INPUT type=hidden id=IdsCto" & rs("AvDetailID") & " name=IdsCto" & rs("AvDetailID") & " value='" & rs("IdsCto_YN") & "'>" & _
								"<INPUT type=hidden id=RctoSkus" & rs("AvDetailID") & " name=RctoSkus" & rs("AvDetailID") & " value='" & rs("RctoSkus_YN") & "'>" & _
								"<INPUT type=hidden id=RctoCto" & rs("AvDetailID") & " name=RctoCto" & rs("AvDetailID") & " value='" & rs("RctoCto_YN") & "'>" & _
								"<INPUT type=hidden id=Weight" & rs("AvDetailID") & " name=Weight" & rs("AvDetailID") & " value='" & rs("Weight") & "'></TD>" & _
								"<TD>" & sNewGpgDescription & "</TD></TR>"
		End If						
		rs.MoveNext
	Loop
		
	rs.Close
		
	If Trim(sAvList) = "" Then
		sAvList = "<TR><TD Colspan=2><B>No AVs Matching the Search Criteria Were Found</B></TD><TR>"
	Else
		sMode = "add"
	End If
		
End Sub

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim i
	Dim returnValue
	Dim iaAvId : iaAvId = Split(Request.Form("AvDetailID"),",")
	Dim cbxValue, sGpgDescription, sMktgDescription, sConfigRules, sIdsSkus, sIdsCto, sRctoSkus, sRctoCto, sWeight, sFeatureCatID
	Dim iAvId
	
	cn.BeginTrans
	
	For i = LBound(iaAvid) To UBound(iaAvid)
		iaAvid(i) = Trim(iaAvid(i))
	Next
	
	For i = LBound(iaAvId) To UBound(iaAvId)
		iAvId = ""
		cbxValue = Request.Form("cbx" & iaAvid(i))
		sGpgDescription = Request.Form("gpgdescription" & iaAvid(i))
		sMktgDescription = Request.Form("MktgDescription" & iaAvid(i))
		sConfigRules = Request.Form("ConfigRules" & iaAvid(i))
		sIdsSkus = Request.Form("idsskus" & iaAvid(i))
		sIdsCto = Request.Form("idscto" & iaAvid(i))
		sRctoSkus = Request.Form("rctoskus" & iaAvid(i))
		sRctoCto = Request.Form("rctocto" & iaAvid(i))
		sWeight = Request.Form("weight" & iaAvid(i))
		sFeatureCatID = Request.Form("FeatureCat" & iaAvid(i))
		
		If LCase(cbxValue) = "on" Then
			
			Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvDetail")
			cmd.NamedParameters = True
			dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
			dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, sFeatureCatID
			dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, sGpgDescription
			dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, sMktgDescription
    		'dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, true
			dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
			dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamOutput, 8, ""
			returnValue = dw.ExecuteNonQuery(cmd)
		    
			iAvId = cmd("@p_AvDetailID")

			If returnValue = 0 Then
				' Abort Transaction
				Response.Write "Error Saving Detail " & sGpgDescription & " : AvDetailID=" & iAvID & " : ReturnValue=" & returnValue & "<BR>"
				cn.RollbackTrans()
				Exit Sub
			End If

			Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand")
    		cmd.NamedParameters = True
			dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
			dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
			dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
			dw.CreateParameter cmd, "@p_ProgramVersion", adVarchar, adParamInput, 5, ""
			dw.CreateParameter cmd, "@p_ConfigRules", adVarchar, adParamInput, 800, sConfigRules
			dw.CreateParameter cmd, "@p_ManufacturingNotes", adVarchar, adParamInput, 800, ""
			dw.CreateParameter cmd, "@p_IdsSkus_YN", adChar, adParamInput, 1, sIdsSkus
			dw.CreateParameter cmd, "@p_IdsCto_YN", adChar, adParamInput, 1, sIdsCto
			dw.CreateParameter cmd, "@p_RctoSkus_YN", adChar, adParamInput, 1, sRctoSkus
			dw.CreateParameter cmd, "@p_RctoCto_YN", adChar, adParamInput, 1, sRctoCto
    		dw.CreateParameter cmd, "@p_SortOrder", adInteger, adParamInput, 4, ""
			dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, ""
       		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, 1
       		dw.CreateParameter cmd, "@p_GSEndDt", adDate, adParamInput, 8, ""
       		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
			returnValue = dw.ExecuteNonQuery(cmd)
		
			If returnValue <> 1 Then
				' Abort Transaction
				Response.Write "Error Saving Product Info " & Request("BID") & " : AvDetailID=" & iAvID & "<BR>"
				cn.RollbackTrans()
				Exit Sub
			End If

		End If
	
	Next
	
	sFunction = "close"
	cn.CommitTrans

End Sub

Select Case LCase(sFunction)
	Case "save"
		Call Save()
	Case "search"
		Call Search()
	Case Else
		Call Main()
End Select
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/excalibur.css">
<LINK rel="stylesheet" type="text/css" href="../scm/style.css">
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT type="text/javascript">
function Body_OnLoad()
{
	switch (frmMain.hidFunction.value)
	{
		case "close":
			window.close();
			break;
	}		
	
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
	}
}


function btnSearch_OnClick()
{
	with (frmMain)
	{
		if (!validateTextInput(txtKMAT, 'KMAT')){ return false; }
		if (!validateTextInput(txtSearch, 'Search String')){ return false; }
		if (!validateTextInput(txtReplace, 'Replace String')){ return false; }
		hidFunction.value = 'search';
		submit();
	}
}
</SCRIPT>
</HEAD>
<BODY OnLoad="Body_OnLoad()">
<FORM method=post id=frmMain>
<INPUT id="hidMode" name="hidMode" type=HIDDEN value=<%= LCase(sMode)%>>
<INPUT id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<INPUT id="BID" name="BID" type=HIDDEN value="<%= iBrandID%>">
<TABLE class="FormTable" bgcolor=cornsilk WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=1 bordercolor=tan>
	<TR>
		<TH>KMAT</TH>
		<TD><INPUT class=txtbox id="txtKMAT" name="txtKMAT" type=text maxlength=6 size=7 value=<%= Request.Form("txtKMAT")%>>-999</TD>
	</TR>
	<TR>
		<TH>Search String</TH>
		<TD><INPUT class=txtbox id="txtSearch" name="txtSearch" type=text value=<%= Request.Form("txtSearch")%>></TD>
	</TR>
	<TR>
		<TH>Replace String</TH>
		<TD><INPUT class=txtbox id="txtReplace" name="txtReplace" type=text value=<%= Request.Form("txtReplace")%>></TD>
	</TR>
</TABLE>
<P align=right><INPUT type="button" value="Search" id=btnSearch name=btnSearch onclick="btnSearch_OnClick()"></P>
<TABLE Class="tblResults">
<TR>
	<TH>&nbsp;</TH>
	<TH width=100%>GPG Description</TH></TR>
<%= sAvList%>
</TABLE>
</FORM>
</BODY>
</HTML>

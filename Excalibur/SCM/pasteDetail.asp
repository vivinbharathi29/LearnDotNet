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
Dim sCategoryOpt		: sCategoryOpt = ""
Dim iCategoryOpt		: iCategoryOpt = ""

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
	Set cmd = dw.CreateCommandSP(cn, "usp_ListAvFeatureCategories")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.EOF
		sCategoryOpt = sCategoryOpt & "<OPTION Value='" & rs("AvFeatureCategoryID") & "'"
		If iCategoryOpt = rs("AvFeatureCategoryID") Then
			sCategoryOpt = sCategoryOpt & " SELECTED "
		End If
		sCategoryOpt = sCategoryOpt & ">" & rs("AvFeatureCategory") & "</OPTION>" & VbCrLf
		rs.MoveNext
	Loop

	rs.Close
		
End Sub

Sub Save()
On Error Goto 0
	Dim saAvNumbers	: saAvNumbers = Split(Request.Form("txtAvNumbers"),vbcrlf)
	Dim i, iAvId, returnValue, saAvDetail
	
	cn.BeginTrans
	
	For i = LBound(saAvNumbers) to UBound(saAvNumbers)
	response.Write UBound(saAvNumbers) & "<br>"

		If Trim(saAvNumbers(i)) <> "" Then
			' Add Record AV Table
			saAvDetail = Split(saAvNumbers(i), vbTab)
			response.Write UBound(saAvDetail) & "<br>"
			If (Trim(saAvDetail(0)) <> "" OR Trim(saAvDetail(2)) <> "") And Ubound(saAvDetail) > 11 Then
				Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvDetail")
				cmd.NamedParameters = True
				dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Trim(saAvDetail(0))
				dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Request.Form("selAvCategory")
				dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Trim(saAvDetail(2))
				dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Trim(saAvDetail(3))
				dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
				dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamOutput, 8, ""
				returnValue = dw.ExecuteNonQuery(cmd)
		
				iAvId = cmd("@p_AvDetailID")
                Response.Write "AvID:"& iAvid & "<br>"
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
				dw.CreateParameter cmd, "@p_ConfigRules", adVarchar, adParamInput, 800, Trim(saAvDetail(7))
				dw.CreateParameter cmd, "@p_ManufacturingNotes", adVarchar, adParamInput, 800, ""
				dw.CreateParameter cmd, "@p_IdsSkus_YN", adChar, adParamInput, 1, GetYNfromX(saAvDetail(8))
				dw.CreateParameter cmd, "@p_IdsCto_YN", adChar, adParamInput, 1, GetYNfromX(saAvDetail(9))
				dw.CreateParameter cmd, "@p_RctoSkus_YN", adChar, adParamInput, 1, GetYNfromX(saAvDetail(10))
				dw.CreateParameter cmd, "@p_RctoCto_YN", adChar, adParamInput, 1, GetYNfromX(saAvDetail(11))
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
		End If
	Next
	
	sFunction = "close"
	cn.CommitTrans

End Sub

Function GetYNfromX( value )
	If InStr(1,value, "x", vbTextCompare) Then
		GetYNfromX = "Y"
	Else
		GetYNfromX = "N"
	End If
End Function

If LCase(sFunction) = "save" Then
	Call Save()
Else
	Call Main()
End If
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/excalibur.css">
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
		//if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
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
		<TH>Feature Category</TH>
		<TD>			<SELECT id=selAvCategory name=selAvCategory>
				<OPTION VALUE=0>--- Please Make a Selection ---</OPTION>
				<%= sCategoryOpt%>			</SELECT>
		</TD>
	</TR>
	<TR>
		<TH>AV Numbers<p style="font-weight:normal;"><font color=red size=1>Copy the AV row from Excel and past it into the text box to the right.  You may copy more than one row at a time.</font></p></TH>
		<TD><TEXTAREA rows=15 id=txtAvNumbers name=txtAvNumbers style="width:300px"></TEXTAREA></TD>
	</TR>
</TABLE>
</FORM>
</BODY>
</HTML>

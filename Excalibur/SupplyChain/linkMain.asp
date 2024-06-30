<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim i
Dim sMode			: sMode = Request.Form("hidMode") : If sMode = "" Then sMode = Request.QueryString("Mode")
Dim sFunction		: sFunction = Request("Function")
Dim sKmat
Dim sEzcKmat
Dim sPlantCd
Dim sProjectCd
Dim sSalesOrg
Dim sAvList

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
If Request.Form("ddlProducts") <> m_ProductVersionID And Request.Form("ddlProducts") <> "" Then
	m_ProductVersionID = Request.Form("ddlProducts")
End If
Dim m_ProductBrandID	: m_ProductBrandID = Request.Form("ddlBrand")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName

Dim m_Table
Dim m_ListBox
Dim m_JSArray


m_IsSysAdmin = false
m_IsProgramCoordinator = false
m_IsConfigurationManager = false
m_EditModeOn = false



'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	'm_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	'm_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
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


Function GetProductVersion( ProductVersionID )
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	GetProductVersion = rs("version") & ""
	
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Function

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

	If Len(Trim(Request("PVID"))) = 0 Or Len(Trim(Request("BID"))) = 0 Then
		Response.Write "<H3>Insufficient information provided.</H3>"
		Response.End
	End If

	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3>"
		Response.End
	End If

	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListProductVersionBrands")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
	If rs.eof And rs.bof Then
		Response.Write "<H3>No Active Products were found.</H3>"
		Response.End
	End If
	
	Dim Rows
	Dim ReturnCode
	Dim ProductVersionID
	Dim sProductOptions
	Dim sLastProductVersion	

	Do Until rs.EOF
		'
		' Create Javascrip Array
		'
		If sLastProductVersion <> rs("ProductVersionID") Then
			sLastProductVersion = rs("ProductVersionID")
			m_JSArray = m_JSArray & "'EOF');" & vbcrlf & "assocArray['ddlProducts=" & rs("ProductVersionID") & "'] = new Array("
			sProductOptions = sProductOptions & "<OPTION Value=" & rs("ProductVersionID") 
			If CLng(sLastProductVersion) = CLng(m_ProductVersionID) Then
				sProductOptions = sProductOptions & " SELECTED"
			End If
			sProductOptions = sProductOptions & ">" & Trim(rs("ProductName")) & " " & Trim(rs("ProductVersion")) & "</OPTION>"
		End If
		m_JSArray = m_JSArray & "'" & rs("ProductBrandID") & "','" & rs("BrandName") & "',"
		
		rs.movenext
	Loop
	
	m_JSArray = Right(m_JSArray, LEN(m_JSArray) - 7) & "'EOF');"
	m_JSArray = "<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>" & vbcrlf & _
		"<!--" & vbcrlf & _
		"var assocArray = new Object();" & vbcrlf & _
		m_JSArray & vbcrlf & _
		"//-->" & vbcrlf & _
		"</SCRIPT>"
		

	m_ListBox = sProductOptions
	
	rs.close
	set rs = nothing
	
End Sub

Sub Search()
	Dim sSearch : sSearch = Request.Form("txtSearch")
	Dim sReplace : sReplace = Request.Form("txtReplace")
	Dim sNewGpgDescription
	Dim sFeatureCat

	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request.Form("ddlProducts"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request.Form("ddlBrand"))
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
	dw.CreateParameter cmd, "@p_GpgDescription", adVarchar, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_UPC", adChar, adParamInput, 12, ""
	dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
	dw.CreateParameter cmd, "@p_KMAT", adChar, adParamInput, 6, ""
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	sAvList = ""
			
	Do Until rs.EOF
		If sFeatureCat <> rs("AvFeatureCategory") Then
			sFeatureCat = rs("AvFeatureCategory")
			sAvList = sAvList & "<TR><TD ID=Feature ColSpan=3>" & sFeatureCat & "</TD></TR>"
		End If
			
		sAvList = sAvList & "<TR><TD><INPUT type=checkbox id=cbx" & rs("AvDetailID") & " name=cbx" & rs("AvDetailID") & ">" & _
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
							"<TD>" & rs("AvNo")&"" & "</TD>" & _
							"<TD>" & rs("GpgDescription")&"" & "</TD></TR>"
		rs.MoveNext
	Loop
		
	rs.Close
		
	If Trim(sAvList) = "" Then
		sAvList = "<TR><TD Colspan=3><B>No AVs Matching the Search Criteria Were Found</B></TD><TR>"
	Else
		sMode = "add"
	End If
		
	Call Main()
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
		iAvId = iaAvid(i)
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
			
			Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand")
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
			dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, ""
      		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, 1
       		dw.CreateParameter cmd, "@p_GSEndDt", adDate, adParamInput, 8, ""
       		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
			returnValue = dw.ExecuteNonQuery(cmd)
		
			If returnValue <> 1 Then
				' Abort Transaction
				Response.Write returnValue
				cn.RollbackTrans()
				Exit Sub
			End If

		End If
	
	Next
	
	sFunction = "close"
	cn.CommitTrans

End Sub

Select Case LCase(Request.Form("hidMode"))
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
<SCRIPT type="text/javascript">
<!--
function Body_OnLoad()
{

	//searchResults.style.display = "";
	
	if (frmMain.hidFunction.value == "LinkFrom")
		frmMain.btnSearch.className="Button";

	if (frmMain.hidFunction.value == "close")
		window.close();

	// -- If user is allowed to edit the settings enable the OK button -- 	
	
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
	}
	
	
	listboxItemSelected(frmMain.ddlProducts,frmMain.ddlBrand);
	if (frmMain.hidBrand.value != "")
		frmMain.ddlBrand.value = frmMain.hidBrand.value;
}

function comboItemSelected(oList1,oList2)
{
	if (oList2!=null)
	{
		clearComboOrList(oList2);
		if (oList1.selectedIndex == -1)
			oList2.options[oList2.options.length] = new Option('Please make a selection from the list', '');
		else 
			fillCombobox(oList2, oList1.name + '=' + oList1.options[oList1.selectedIndex].value);
	}
}

function listboxItemSelected(oList1,oList2)
{
	if (oList2!=null)
	{
		clearComboOrList(oList2);
		if (oList1.selectedIndex == -1)
			oList2.options[oList2.options.length] = new Option('Please make a selection from the list', '');
		else 
			fillListbox(oList2, oList1.name + '=' + oList1.options[oList1.selectedIndex].value);
	}
}

function clearComboOrList(oList)
{
	for (var i = oList.options.length - 1; i >= 0; i--)
	{
		oList.options[i] = null;
	}
	oList.selectedIndex = -1;
	if (oList.onchange)	oList.onchange();
}

function fillCombobox(oList, vValue)
{
	if (vValue != '') 
	{
		if (assocArray[vValue])
		{
			oList.options[0] = new Option('Please make a selection', '');
			var arrX = assocArray[vValue];
			for (var i = 0; i < arrX.length; i = i + 2)
			{
				if (arrX[i] != 'EOF') oList.options[oList.options.length] = new Option(arrX[i + 1].split('&amp;').join('&'), arrX[i]);
			}
			if (oList.options.length == 1)
			{
				oList.selectedIndex=0;
				if (oList.onchange) oList.onchange();
			}
		} 
		else 
		{
			oList.options[0] = new Option('None found', '');
		}
	}
}

function fillListbox(oList, vValue)
{
	if (vValue != '') 
	{
		if (assocArray[vValue])
		{
			oList.options[oList.options.length] = new Option('-- Make A Selection --');
			var arrX = assocArray[vValue];
			for (var i = 0; i < arrX.length; i = i + 2)
			{
				if (arrX[i] != 'EOF') oList.options[oList.options.length] = new Option(arrX[i + 1].split('&amp;').join('&'), arrX[i]);
			}
			if (oList.options.length == 1)
			{
				oList.selectedIndex=0;
				if (oList.onchange)	oList.onchange();
			}
		} 
		else 
		{
			oList.options[0] = new Option('None found', '');
		}
	}
}

function btnSearch_OnClick()
{
	frmMain.hidMode.value = "search";
	frmMain.submit();
}
//-->
</SCRIPT>
<%= m_JSArray %>
</HEAD>
<BODY OnLoad="Body_OnLoad()">
<FORM method=post id=frmMain>
<INPUT id="hidMode" name="hidMode" type=HIDDEN value="<%= LCase(sMode)%>">
<INPUT id="hidFunction" name="hidFunction" type=HIDDEN value="<%= sFunction%>">
<INPUT id="hidBrand" name="hidBrand" type=HIDDEN value="<%= m_ProductBrandID%>">
<TABLE class="FormTable" bgcolor=cornsilk WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
<tr><th>Product Version</th><th>Brand</th><th>&nbsp;&nbsp;</th></tr>
<tr><td>
<SELECT class=TextBox id=ddlProducts name=ddlProducts onChange='listboxItemSelected(this.form.ddlProducts,this.form.ddlBrand);'>
<%= m_ListBox%>
</SELECT>
</td><td>
<SELECT class=TextBox id=ddlBrand name=ddlBrand></SELECT>
</td>
<td><INPUT type="button" class=Hidden value="Search" id=btnSearch name=btnSearch onclick="btnSearch_OnClick()"></td></tr></table>
<BR>

<TABLE Class="tblResults">
<TR>
	<TH nowrap width=25>&nbsp;</TH>
	<TH nowrap width=50>Av No.</TH>
	<TH nowrap>GPG Description</TH></TR>
<%= sAvList%>
</TABLE>
</FORM>
</BODY>
</HTML>



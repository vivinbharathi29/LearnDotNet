<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
'Response.Write Request.QueryString
'Response.End

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim i
Dim sMode				: sMode = Request.QueryString("Mode")
Dim sAvNo				: sAvNo = ""
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim sGeoList			: sGeoList = ""
Dim sRegionList			: sRegionList = ""
Dim sReadyDates			: sReadyDates = ""
Dim sAvName				: sAvName = ""
Dim m_JSArray			: m_JSArray = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim m_IsSupplyChainUser


'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

' Debug Section
'
'	If Security.CurrentUserID = 1396 Then
'		m_IsSysAdmin = False
'		Security.CurrentUserID = 1288
'		Response.Write Security.CurrentUserID
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSysEngProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSystemTeamLead(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramCoordinator(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write m_ProductVersionID
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If


	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_IsSupplyChainUser = Security.UserInRole(m_ProductVersionID, "SUPPLYCHAIN")
	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Or m_IsSupplyChainUser Then
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

Function GetBoolValue( value )

	Select Case UCase( value )
		Case "Y"
			GetBoolValue = true
		Case "N"
			GetBoolValue = false
		Case 1
			GetBoolValue = true
		Case 0
			GetBoolValue = false
		Case "T"
			GetBoolValue = true
		Case "F"
			GetBoolValue = false
		Case Else
			GetBoolValue = false
	End Select

End Function

Sub Main()

	Dim sLastGeo
	Dim sLastRegion
	Dim sAllOption
	Dim sAllArray
	
	sAllOption = "<OPTION Value=0>All</OPTION>"
	sAllArray = "assocArray['lbGeo=0'] = new Array("

	Set cmd = dw.CreateCommandSP(cn, "usp_ListConfigCodes")
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.EOF
	
		If sLastGeo <> rs("geoshortname") Then
			sLastGeo = rs("geoshortname")
			sGeoList = sGeoList & "<OPTION Value=" & rs("geoid") & ">" & rs("geoshortname") & "</OPTION>"
			m_JSArray = m_JSArray & "'EOF');" & vbcrlf & "assocArray['lbGeo=" & rs("geoid") & "'] = new Array("
		End If
		
		sRegionList = sRegionList & "<OPTION Value=" & rs("optionconfig") & ">" & rs("optionconfig") & " - " & rs("name") & "</OPTION>"
		
		m_JSArray = m_JSArray & "'" & rs("optionconfig") & "','" & rs("optionconfig") & " - " & rs("name") & "',"
		sAllArray = sAllArray & "'" & rs("optionconfig") & "','" & rs("optionconfig") & " - " & rs("name") & "',"

		rs.MoveNext
	Loop
	rs.Close
	
	sGeoList = sAllOption & sGeoList
    If LEN(m_JSArray) > 7 Then
	    m_JSArray = Right(m_JSArray, LEN(m_JSArray) - 7) & "'EOF');"
    End If
	sAllArray = sAllArray & "'EOF');"
	m_JSArray = "<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>" & vbcrlf & _
		"<!--" & vbcrlf & _
		"var assocArray = new Object();" & vbcrlf & _
		sAllArray & m_JSArray & vbcrlf & _
		"//-->" & vbcrlf & _
		"</SCRIPT>"
		
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvMrDates")
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sLastGeo = ""
	Do Until rs.EOF
		If sLastGeo <> rs("geoshortname") Then
			sLastGeo = rs("geoshortname")
			sReadyDates = sReadyDates & "<TR id=region><td colspan=3>" & sLastGeo & "</td></tr>"
		End If
		
		
		sReadyDates = sReadyDates & "<TR><TD nowrap>" & rs("optionconfig") & "</TD><TD nowrap>" & rs("name") & _
			"</TD><TD>" & rs("readydate") & "</TD></TR>" 
	
		rs.MoveNext
	Loop
	rs.Close
	
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sAvName = rs("AvNo") & " - " & rs("GPGDescription")
	
	rs.Close
		

End Sub

Function GetSkuCtoValue( value )
	If lcase(Value) = "y" Then
		GetSkuCtoValue = "X"
	Else
		GetSkuCtoValue = "&nbsp;"
	End If
End Function

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	Dim iAvId
	Dim saConfigCode : saConfigCode = Split(Request.Form("lbRegion"), ",")
	Dim configCode
	Dim i : i = 0
	
	cn.BeginTrans

	For Each configCode In Request.Form("lbRegion")
		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvMrDate")
		cmd.NamedParameters = True
		dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, Request.QueryString("AVID")
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request.QueryString("BID")
		dw.CreateParameter cmd, "@p_OptionConfig", adVarchar, adParamInput, 5, configCode
		dw.CreateParameter cmd, "@p_ReadyDate", adDate, adParamInput, 8, Request.Form("txtDate")
		returnValue = dw.ExecuteNonQuery(cmd)

		If returnValue <> 1 Then
			' Abort Transaction
			Response.Write returnValue
			cn.RollbackTrans()
			Exit Sub
		End If
	Next

	cn.CommitTrans
	
	sFunction = ""
	Call Main()
	
End Sub

If LCase(sFunction) = "save" Then
	Call Save()
Else
	Call Main()
End If
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../style/excalibur.css">
<script type="text/javascript">
function Body_OnLoad()
{
}

function btnSave_OnClick()
{
	frmMain.hidFunction.value = "save";
	frmMain.submit();
}

function cmdDate_onclick(FieldID) {
	var strID;
	var oldValue = window.frmMain.elements(FieldID).value;
		
	strID = window.showModalDialog("../mobilese/today/caldraw1.asp",FieldID,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) == "undefined")
		return
	
	window.frmMain.elements(FieldID).value = strID;
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

function fillListbox(oList, vValue)
{
	if (vValue != '') 
	{
		if (assocArray[vValue])
		{
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
			
            for (var i=0; i<oList.length; i++) 
                oList[i].selected = oList[i].checked = true

		} 
		else 
		{
			oList.options[0] = new Option('None found', '');
		}
	}
}
</script>
<%= m_JSArray %>
</head>
<body OnLoad="Body_OnLoad()">
<form method="post" id="frmMain">
<input id="hidMode" name="hidMode" type="HIDDEN" value="<%= LCase(sMode)%>">
<input id="hidFunction" name="hidFunction" type="HIDDEN" value="<%= LCase(sFunction)%>">
<input id="hidAVID" name="hidAVID" type="HIDDEN" value="<%= Request("AVID")%>">
<p><%= sAvName%></p>
<p><select class="selectbox" size="5" id="lbGeo" name="lbGeo" onChange="listboxItemSelected(this.form.lbGeo,this.form.lbRegion);">

<%= sGeoList%>

</select>
&nbsp;
<select class="selectbox" size="5" id="lbRegion" name="lbRegion" multiple>

<%= sRegionList%>

</select>
</p>
<p>
<table class="FormTable" WIDTH="100%" style="border-collapse:collapse;">
<tr>
	<th>Select Date</th>
	<td>
		<input type="text" id="txtDate" name="txtDate">
		<a href="javascript: cmdDate_onclick('txtDate')"><img ID="picTarget" SRC="/mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a></td></tr>
<tr>
	<td ColSpan="2" Align="Right">
		<input class="button" type="button" value="Save" id="btnSave" name="btnSave" onclick="btnSave_OnClick()"></td></tr>
</table>
</p>
<p>
<table Class="FormTable" WIDTH="100%" style="border-collapse:collapse;">
<tr>
	<th>Config&nbsp;Cd.</th>
	<th>Country</th>
	<th width="100%">Ready Date</th></tr>
<%= sReadyDates%>	
</table>
</p>
</form>
</body>
</html>

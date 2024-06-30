<%@ Language=VBScript %>

<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeout = 600
%>
<!-- #include file="../includes/ExcelExport.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/xml-helper.asp" --> 

<%
Dim cn
Dim cmd
Dim dw
Dim rs
Dim xlDoc
Dim fileName
Dim dtYear
Dim dtMonth
Dim dtDay
Dim platformName
Dim scmName
Dim dtPddLock
Dim iSheetRowCount
Dim sText
Dim sLanguages
Dim sBrand
Dim sRegion
Dim sCategory
Dim sRootName
Dim sContacts
Dim iProgramVersion
Dim dtLastPublish
Dim dtLastPublishDt

dtYear = DatePart("yyyy", Now())
dtMonth = DatePart("m", Now())
dtDay = DatePart("d", Now())

Select Case dtMonth
	Case 1
		dtMonth = "JAN"
	Case 2
		dtMonth = "FEB"
	Case 3
		dtMonth = "MAR"
	Case 4
		dtMonth = "APR"
	Case 5
		dtMonth = "MAY"
	Case 6
		dtMonth = "JUN"
	Case 7
		dtMonth = "JUL"
	Case 8
		dtMonth = "AUG"
	Case 9
		dtMonth = "SEP"
	Case 10
		dtMonth = "OCT"
	Case 11
		dtMonth = "NOV"
	Case 12
		dtMonth = "DEC"
End Select

If Len(dtDay) = 1 Then
	dtDay = "0" & dtDay
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")


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
	
	Set Security = Nothing
'##############################################################################	

If m_EditModeOn And Request("Publish") = "True" Then
	Set cmd = dw.CreateCommandSP(cn, "usp_InsertScmPublishEvent")
	cmd.NamedParameters = True
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	dw.CreateParameter cmd, "@p_UserName", adVarChar, adParamInput, 50, m_UserFullName
	cmd.Execute
End If

Function ConvertYN( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		ConvertYN = ""
	ElseIf Trim(UCase( value )) = "Y" Then
		ConvertYN = "X"
	End If

End Function

Function GetCellStyle( ChangedBit )
	If ChangedBit Then
		GetCellStyle = "s32"
	Else
		GetCellSTyle = "s31"
	End If
End Function

'
' Get Scm Name
'
Set cmd = dw.CreateCommandSP(cn, "usp_ListScmNames")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, "" 'Request("PVID") 
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

iProgramVersion = Trim(XmlSafe(rs("Version") & ""))

scmName = Trim(XmlSafe(rs("name"))&"") & " " & Trim(XmlSafe(rs("version")&"")) & " " & Trim(XmlSafe(rs("seriesname")&""))

fileName = "SCM_" & Trim(rs("seriesname")) & "_" & Trim(rs("name")) & "_" & Trim(rs("version")) & "_" & dtMonth & "_" & dtDay & "_" & dtYear & ".xls"

fileName = Replace(fileName, "/", "_")

If m_EditModeOn And Request("Publish") = "True" Then
	dtLastPublish = now()
else
	dtLastPublish = rs("LastPublishDt")&""
End If

dtLastPublishDt = rs("LastPublishDt")&""
If Not IsDate(dtLastPublishDt) Then
    dtLastPublishDt = NOW()
End If

rs.Close

Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
Response.ContentType = "application/vnd.ms-excel"
'Response.ContentType = "text/xml"

sContacts = "<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>"
sContacts = sContacts & "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">"
sContacts = sContacts & "<DocumentProperties xmlns=""urn:schemas-microsoft-com:office:office""><Version>11.6408</Version></DocumentProperties>"
sContacts = sContacts & "<ExcelWorkbook xmlns=""urn:schemas-microsoft-com:office:excel""><WindowHeight>4785</WindowHeight><WindowWidth>7650</WindowWidth><WindowTopX>7665</WindowTopX><WindowTopY>-15</WindowTopY><TabRatio>735</TabRatio><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows><ActiveSheet>2</ActiveSheet></ExcelWorkbook>"
sContacts = sContacts & "<Styles>"
sContacts = sContacts & "<Style ss:ID=""Default"" ss:Name=""Normal""><Alignment ss:Vertical=""Top""/><Borders/><Font/><Interior/><NumberFormat/><Protection/></Style>"
sContacts = sContacts & "<Style ss:ID=""m19709554""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""m19709564""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""m19709574""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s11""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders></Style>"
sContacts = sContacts & "<Style ss:ID=""s21""><Font x:Family=""Swiss"" ss:Size=""20"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s22""><Font x:Family=""Swiss"" ss:Size=""12"" ss:Color=""#0000FF"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s23""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s24""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s25""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#00CCFF"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s26""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#99CC00"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s27""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#CCFFCC"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s28""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#FF99CC"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s29""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s30""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s31""><Alignment ss:Vertical=""Center"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders></Style>"
sContacts = sContacts & "<Style ss:ID=""s32""><Alignment ss:Vertical=""Center"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#FFFF99"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s33""><Alignment ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:Bold=""1""/><Interior/><NumberFormat/><Protection/></Style>"
sContacts = sContacts & "<Style ss:ID=""s34""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#00FFFF"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s35""><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s36""><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/><NumberFormat ss:Format=""m/d/yyyy;@""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s37""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s38""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s39""><Alignment ss:Horizontal=""Right"" ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s40""><Alignment ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s41""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><NumberFormat ss:Format=""m/d/yyyy;@""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s42""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders></Style>"
sContacts = sContacts & "<Style ss:ID=""s43""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s44""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/><NumberFormat ss:Format=""m/d/yyyy;@""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s45""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s46""><Alignment ss:Horizontal=""Left"" ss:Vertical=""Top""/><Interior ss:Color=""#CCFFCC"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s47""><Borders><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s48""><Borders><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Size=""12"" ss:Bold=""1""/><NumberFormat ss:Format=""m/d/yyyy;@""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s56""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s57""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior/><NumberFormat/><Protection/></Style>"
sContacts = sContacts & "<Style ss:ID=""s58""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:Rotate=""90"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s59""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Color=""#FF0000"" ss:Bold=""1""/><Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s60""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Color=""#FF0000"" ss:Bold=""1""/><Interior ss:Color=""#00CCFF"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s61""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Color=""#FF0000"" ss:Bold=""1""/><Interior ss:Color=""#99CC00"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s62""><Alignment ss:Horizontal=""Center"" ss:Vertical=""Top""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/><Interior ss:Color=""#000000"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s63""><Alignment ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders></Style>"
sContacts = sContacts & "<Style ss:ID=""s64""><Alignment ss:Vertical=""Top"" ss:WrapText=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s65""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s66""><NumberFormat ss:Format=""Short Date""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s67""><Alignment ss:Vertical=""Center""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Swiss"" ss:Bold=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s68""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Interior ss:Color=""#FFFF99"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s71""><Font ss:StrikeThrough=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s72""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:StrikeThrough=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s73""><Alignment ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:StrikeThrough=""1""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s74""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:StrikeThrough=""1""/><Interior ss:Color=""#FFFF99"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s75""><Alignment ss:Vertical=""Top"" ss:WrapText=""1""/><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:StrikeThrough=""1""/><Interior ss:Color=""#FFFF99"" ss:Pattern=""Solid""/></Style>"
sContacts = sContacts & "<Style ss:ID=""s80""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font ss:Color=""#FF0000"" ss:Bold=""1""/><Interior ss:Color=""#000000"" ss:Pattern=""Solid""/></Style>"

sContacts = sContacts & "</Styles>"

Set cmd = dw.CreateCommandSP(cn, "usp_SelectSystemTeam")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 
Set rs = dw.ExecuteCommandReturnRS(cmd)

sContacts = sContacts & "<Worksheet ss:Name=""Using the SCM - Contacts&amp;Terms""><Names><NamedRange ss:Name=""Print_Area"" ss:RefersTo=""='Using the SCM - Contacts&amp;Terms'!R1C1:R47C6""/></Names>"
sContacts = sContacts & "<Table x:FullColumns=""1"" x:FullRows=""1""><Column ss:AutoFitWidth=""0"" ss:Width=""141""/><Column ss:AutoFitWidth=""0"" ss:Width=""171.75""/><Column ss:AutoFitWidth=""0"" ss:Width=""194.25""/><Column ss:AutoFitWidth=""0"" ss:Width=""180.75""/><Column ss:Width=""99.75""/><Column ss:AutoFitWidth=""0"" ss:Width=""179.25""/>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""26.25""><Cell ss:StyleID=""s21""><Data ss:Type=""String"">" & scmName & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0""/><Row ss:AutoFitHeight=""0"" ss:Height=""15.75""><Cell ss:StyleID=""s22""><Data ss:Type=""String"">Contacts</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s23""><Data ss:Type=""String"">Business Function</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">Name</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s24""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">"
'
' Insert PC's Name Here
'
sContacts = sContacts & Trim(XmlSafe(rs("pc_name")&""))
sContacts = sContacts & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s25""><Data ss:Type=""String"">Marketing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">"
'
' Insert Marketing Contact Here
'
Dim sMarketing
If Trim(XmlSafe(rs("commercialmarketing_name")&"")) <> "" Then
	sMarketing = ", " & Trim(XmlSafe(rs("commercialmarketing_name")&""))
End If
If Trim(XmlSafe(rs("smbmarketing_name")&"")) <> "" Then
	sMarketing = sMarketing & ", " & Trim(XmlSafe(rs("smbmarketing_name")&""))
End If
If Trim(XmlSafe(rs("consumermarketing_name")&"")) <> "" Then
	sMarketing = sMarketing & ", " & Trim(XmlSafe(rs("consumermarketing_name")&""))
End If
sMarketing = Mid(sMarketing, 2)
sContacts = sContacts & sMarketing
sContacts = sContacts & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row> <Cell ss:StyleID=""s26""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String""><NamedCell ss:Name=""Print_Area""/>"
'
' Insert Marketing Ops Contact Here
'
sContacts = sContacts & Trim(XmlSafe(rs("MarketingOps_name")&""))
sContacts = sContacts & "</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s27""><Data ss:Type=""String"">Supply Chain</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">"
'
' Insert Supply Chain Contact Here
'
sContacts = sContacts & Trim(XmlSafe(rs("supplychain_name")&""))

sContacts = sContacts & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row> <Cell ss:StyleID=""s28""><Data ss:Type=""String"">Plat Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">"
'
' Insert Platform Engineering
'
sContacts = sContacts & Trim(XmlSafe(rs("platformdev_name")&""))

sContacts = sContacts & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s29""><Data ss:Type=""String"">Service</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">"
'
' Insert Service Contact Here
'
sContacts = sContacts & Trim(XmlSafe(rs("service_name")&""))

rs.Close

sContacts = sContacts & "</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row ss:Index=""14"" ss:AutoFitHeight=""0"" ss:Height=""15.75""><Cell ss:StyleID=""s22""><Data ss:Type=""String"">Guidelines for Change control </Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Item</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Action required on SCM Data tab</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Action required on Change Form tab</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Add AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Add in SCM document in Col A</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Add in change log</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Delete AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Strike through AV# in Col A</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Marketing Description</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Correct description</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log (To/From) by AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Typos/Correcting mistakes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Correct mistake</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log (To/From) by AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Date Changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">strike through and put new date</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">none</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Configuration Rules changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">add/delete/correct changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log in comments</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Localization changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">add/delete/correct changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log (To/From) by AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">UPC changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">add/delete/correct changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log (To/From) by AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Weight changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">add/delete/correct changes</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">track on change log (To/From) by AV#</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row ss:Index=""26""><Cell><Data ss:Type=""String"">Prior to 1st publish (1.0): The initial creation to initial publish, only AV#s need to be tracked on change control.</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row ss:Index=""29"" ss:AutoFitHeight=""0"" ss:Height=""15.75""><Cell ss:MergeAcross=""1"" ss:StyleID=""s22""><Data ss:Type=""String"">SCM Field Documentation</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Column Header (row 4)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Column Subheader (row 5)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Description</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Populated By</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">End User of Data</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s30""><Data ss:Type=""String"">Comments</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">AV# </Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Feature Product ID (PID)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Feature Category</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">n/a</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Main category of the feature</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">KE</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">GPG Description</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">SAP global description of PID</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">GBU, Region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing  Description (30 Characters Max)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Consumer description of PID</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">GPSy</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Program Version where used</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">x.y release of platform</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">GBU, Region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">&quot;</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">&quot;</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">x.y+1 release of platform</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">GBU, Region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">need to define rules</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">CPL Blind Date</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">First date orders can be taken for this AV PID.</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Conrad order checker</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">The date will be shown on the Base Unit row with exceptions only noted in the CPL Blind Date column</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">End of Manufacturing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">End of Manufacturing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">RAS Automation</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">The date will be shown on the Base Unit row with exceptions only noted in the End of Manufacturing column</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Configuration Rules</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Notes of capability and availability by region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">STL Program Manager</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Region</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">IDS-SKUS</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">X = IDS-SKUS Fullfillment/Deployment</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">KE</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">IDS-CTO</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">X = IDS-CTO Fullfillment/Deployment</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">KE</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">RCTO-SKUS</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">X = RCTO-SKUS Fullfillment/Deployment</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">KE</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">RCTO-CTO</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">X = RCTO-CTO Fullfillment/Deployment</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">KE</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">EMEA Offering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">EMEA</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Field only to be used by EMEA</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">EMEA </Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">EMEA</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">UPC</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">UPC code for PID (if applicable)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">NA Retail Web Kiosk CTO</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Weight (in oz)</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Weight of PID</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">???</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">NB will leave blank until user and rules are defined</Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Region or Country</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Localization Suffix</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Region/Country/Suffix where PID is available per Base Unit Date or Suffix Exception Only Date</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Operations</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s11""><Data ss:Type=""String"">Region</Data><NamedCell ss:Name=""Print_Area""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s11""><ss:Data ss:Type=""String"" xmlns=""http://www.w3.org/TR/REC-html40""><B>Notebook usage:  </B><Font>&#10;The Base Unit date is the date applicable to all PIDs.  Exceptions Only are noted for specific PIDs.&#10;&#10;</Font><B>Types of NB Localizations:</B><Font>&#10;1. Localized independent - date per base unit; exceptions only noted per PID&#10;2. Localized by AV# (ex: keyboards) - date per base unit; exceptions only noted per PID&#10;3. Substitution Localization (ex: modems, DIB) - localized per config rules with date per base unit; exceptions only noted per PID</Font></ss:Data><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
sContacts = sContacts & "</Table><WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><PageSetup><Layout x:Orientation=""Landscape""/><Header x:Margin=""0.25""/><Footer x:Margin=""0.26""/><PageMargins x:Bottom=""0.41"" x:Top=""0.5""/></PageSetup><FitToPage/><Print><FitHeight>2</FitHeight><ValidPrinterInfo/><Scale>67</Scale><HorizontalResolution>600</HorizontalResolution><VerticalResolution>600</VerticalResolution></Print><Selected/><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios><Zoom>75</Zoom></WorksheetOptions></Worksheet>"

Response.Write sContacts

'
' Begin Change Log Tab
'
sContacts = "<Worksheet ss:Name=""SCM Change Log"">"
sContacts = sContacts & "<Table ss:ExpandedColumnCount=""8"" x:FullColumns=""1"" x:FullRows=""1""><Column ss:Width=""141""/><Column ss:Width=""129.75""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:AutoFitWidth=""0"" ss:Width=""126.75""/><Column ss:AutoFitWidth=""0"" ss:Width=""126.75""/><Column ss:AutoFitWidth=""0"" ss:Width=""126.75""/><Column ss:Width=""630.75""/>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""27.75""><Cell ss:StyleID=""s21""><Data ss:Type=""String"">" & scmName & "</Data></Cell><Cell ss:Index=""4"" ss:StyleID=""s34""><Data ss:Type=""String"">A</Data></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Added </Data></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""22.5""><Cell ss:StyleID=""s35""><Data ss:Type=""String"">Published</Data></Cell><Cell ss:StyleID=""s36""><Data ss:Type=""String"">" & dtLastPublish & "</Data></Cell><Cell ss:Index=""4"" ss:StyleID=""s37""><Data ss:Type=""String"">C</Data></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Changed </Data></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""24""><Cell ss:StyleID=""s35""><Data ss:Type=""String"">Today's date</Data></Cell><Cell ss:StyleID=""s36"" ss:Formula=""='SCM AV data'!RC[1]""><Data ss:Type=""DateTime"">2005-09-13T00:00:00.000</Data></Cell><Cell ss:Index=""4"" ss:StyleID=""s38""><Data ss:Type=""String"">O</Data></Cell><Cell ss:StyleID=""s31""><Data ss:Type=""String"">Obsoleted</Data></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""40.5""/>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""26.25""><Cell ss:StyleID=""s39""><Data ss:Type=""String"">Changed Date</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Changed By</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">AV#</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Change</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Field Changed</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Change From</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Change To</Data></Cell><Cell ss:StyleID=""s40""><Data ss:Type=""String"">Comment</Data></Cell></Row>"
'
' Begin Change Log Data
'
Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvHistory")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
Set rs = dw.ExecuteCommAndReturnRS(cmd)

Do Until rs.EOF

	If rs("ShowOnScm") Then
	Select Case rs("AvChangeTypeCd")
		Case "A"
			'
			' Added
			'
			sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_Date")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_User")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s34""><Data ss:Type=""String"">" & XmlSafe(rs("AvChangeTypeCd")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ColumnChanged")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("OldValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("NewValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Comments")&"") & "</Data></Cell></Row>"
		Case "C","M"
			'
			' Changed
			'
			sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_Date")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_User")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s37""><Data ss:Type=""String"">C</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ColumnChanged")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("OldValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("NewValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Comments")&"") & "</Data></Cell></Row>"
		Case "O"
			'
			' Obsoleted
			'
			sContacts = sContacts & "<Row><Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_Date")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_User")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s38""><Data ss:Type=""String"">O</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ColumnChanged")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("OldValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("NewValue")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("Comments")&"") & "</Data></Cell></Row>"
		Case "P"
			sContacts = sContacts & "<Row><Cell ss:StyleID=""s45""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_Date")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s45""><Data ss:Type=""String"">" & XmlSafe(rs("Last_Upd_User")&"") & "</Data></Cell>" & _
			    "<Cell ss:StyleID=""s45""></Cell>" & _
			    "<Cell ss:StyleID=""s45""></Cell>" & _
			    "<Cell ss:StyleID=""s45""></Cell>" & _
			    "<Cell ss:StyleID=""s45""></Cell>" & _
			    "<Cell ss:StyleID=""s45""></Cell>" & _
			    "<Cell ss:StyleID=""s45""><Data ss:Type=""String"">" & XmlSafe(rs("Comments")&"") & "</Data></Cell></Row>"

	End Select	
	End If
	rs.MoveNext
Loop
rs.Close

sContacts = sContacts & "</Table><WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">"
sContacts = sContacts & "<PageSetup><Layout x:Orientation=""Landscape""/><Footer x:Margin=""0.19""/><PageMargins x:Bottom=""0.47"" x:Left=""0.37"" x:Right=""0.25"" x:Top=""0.42""/></PageSetup><DisplayPageBreak/>"
sContacts = sContacts & "<FitToPage/><Print><FitHeight>0</FitHeight><ValidPrinterInfo/><Scale>86</Scale><HorizontalResolution>600</HorizontalResolution><VerticalResolution>600</VerticalResolution><Gridlines/></Print>"
sContacts = sContacts & "<Zoom>75</Zoom><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>"
sContacts = sContacts & "</Worksheet>"
Response.Write sContacts

'
' Begin SCM AV Data Worksheet
'
sContacts = ""
'
' Hide MR Dates for regions not supported by the platform
'
'ss:Hidden=""1"" 
Dim sRegionCds
Dim iHidden
Set cmd = dw.CreateCommAndSP(cn, "usp_SelectBrandLocalizationData")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
Set rs = dw.ExecuteCommAndReturnRS(cmd)

If Not rs.EOF Then
	rs.Sort = "OptionConfig"
End If

Do Until rs.EOF
	If instr(UCASE(sRegionCds), UCASE(rs("OptionConfig"))) = 0 Then
		sRegionCds = sRegionCds & rs("OptionConfig") & ","
	End If
	rs.MoveNext
Loop
rs.Close

sRegionCds = UCASE(sRegionCds)

'Response.Write vbcrlf
'Response.Write sregioncds
'Response.End

If instr(sRegionCds, "AC4") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Index=""17"" ss:Width=""55""/>"
If instr(sRegionCds, "ABL") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABC") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABM") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ACH") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABA") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABV") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "UUG") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKB") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABY") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABB") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABF") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABD") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "B1A") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKC") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "A2M") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABT") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABZ") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABH") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABN") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKD") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB9") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ACB") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKR") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKN") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ACQ") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABE") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AK8") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "UUZ") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB8") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABU") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "UUF") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABG") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB5") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ACJ") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ABJ") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "ACF") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB1") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "B1L") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB2") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AB0") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"
If instr(sRegionCds, "AKL") Then iHidden = 0 Else iHidden = 1
sContacts = sContacts & "<Column ss:Hidden=""" & iHidden & """ ss:Width=""55""/>"


sContacts = sContacts & "<Row><Cell ss:StyleID=""s21""><Data ss:Type=""String"">" & scmName & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s43""><Data ss:Type=""String"">Published</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s43""><Data ss:Type=""String"">" & dtLastPublish & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s44""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:Index=""17"" ss:MergeAcross=""39"" ss:StyleID=""s46""><Data ss:Type=""String"">* Provides MR Date by Localized location (date format mm/dd/yy)</Data></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""30""><Cell ss:StyleID=""s47""><Data ss:Type=""String"">Today's date</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s47""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s48"" ss:Formula=""=TODAY()""><Data ss:Type=""DateTime"">2005-09-13T00:00:00.000</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>"

sContacts = sContacts & "<Cell ss:Index=""17"" ss:MergeAcross=""5"" ss:StyleID=""m19709554""><Data ss:Type=""String"">Americas</Data></Cell><Cell ss:MergeAcross=""24"" ss:StyleID=""m19709564""><Data ss:Type=""String"">EMEA</Data></Cell><Cell ss:MergeAcross=""10"" ss:StyleID=""m19709574""><Data ss:Type=""String"">APD</Data></Cell></Row>"

sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""57.75""><Cell ss:StyleID=""s56""><Data ss:Type=""String"">AV# </Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">Feature Category</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">GPG Description&#10;(no ' or &quot; quote marks)</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s56""><Data ss:Type=""String"">Marketing  Description</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">Program Version where used</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">CPL Blind Date</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s56""><Data ss:Type=""String"">End of Manufacturing</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">Configuration Rules</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">IDS-SKUS</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">IDS-CTO</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">RCTO-SKUS</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">RCTO-CTO</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s56""><Data ss:Type=""String"">UPC</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">Weight (in oz)</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s56""><Data ss:Type=""String"">Global&#10;Series&#10;Modules&#10;(GS End Date)</Data><NamedCell ss:Name=""Print_Titles""/><NamedCell ss:Name=""Print_Area""/></Cell>"

sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Regions</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Brazil</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Canada</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Canadian (French)</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Latin America</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Latin America English</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">United States</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Arabia</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Belgium</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Czech Republic</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Denmark</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Euro</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">France</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Germany</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Greece</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Hungary</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Iceland</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Israel (Hebrew)</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Italy</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Netherlands</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Norway</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Poland</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Portugal</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Russia</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Slovakia</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Slovenia</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">South Africa</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Spain</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Sweden / Finland</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Switzerland</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Turkey</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">United Kingdom</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Asia Pacific</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Australia</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Hong Kong</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">India</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Japan</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Japan (English)</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Korea</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">PRC (CH,EN)</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">PRC Chinese</Data><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s58""><Data ss:Type=""String"">Taiwan</Data><NamedCell ss:Name=""Print_Titles""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s58""><Data ss:Type=""String"">Thailand</Data><NamedCell ss:Name=""Print_Titles""/></Cell></Row>"

sContacts = sContacts & "<Row ss:AutoFitHeight=""1""><Cell ss:StyleID=""s23"" ss:Index=""17""><Data ss:Type=""String"">BRZL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">CAN/ENG</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">FCAN</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">LTNA</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">LAE</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">US</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ARAB</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">EUROA4</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">CZECH</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">DEN</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">EURO</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">FR</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">GR</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">GK</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">HUNG</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ICELAND</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">HE</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ITL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">NL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">NOR</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">POL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">PORT</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">RUSS</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SK</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SA</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SP</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SE/FI</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">SWIS2</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">TURK</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">UK</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">A/P</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AUST</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">HK</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">INDIA</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">JPN2</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">JPN/ENG</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">KOR</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ASIA</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">PRC</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">TW</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">THAI</Data></Cell></Row>"

sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""22.5""><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s59""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell><Cell ss:StyleID=""s60""><Data ss:Type=""String"">Marketing</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s61""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s61""><Data ss:Type=""String"">Marketing Ops</Data><NamedCell ss:Name=""Print_Area""/></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Eng</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell><Cell ss:StyleID=""s59""><Data ss:Type=""String"">BOM Engineering</Data><NamedCell ss:Name=""Print_Area""/></Cell>"

sContacts = sContacts & "<Cell ss:StyleID=""s62""/><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AC4</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABL</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABC</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABM</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ACH</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABA</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABV</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">UUG</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKB</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABY</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABB</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABF</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABD</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">B1A</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKC</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">A2M</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABT</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABZ</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABH</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABN</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKD</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB9</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ACB</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKR</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKN</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ACQ</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABE</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AK8</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">UUZ</Data></Cell>"
sContacts = sContacts & "<Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB8</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABU</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">UUF</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABG</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB5</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ACJ</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ABJ</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">ACF</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB1</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">B1L</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB2</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AB0</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">AKL</Data></Cell></Row>"

Set cmd = dw.CreateCommandSP(cn, "usp_SelectScmDetail")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID") 
Set rs = dw.ExecuteCommandReturnRS(cmd)

Dim sHeader, iRowCount
iRowCount = rs.RecordCount + 90
sHeader = "<Worksheet ss:Name=""SCM AV Data""><Names><NamedRange ss:Name=""Print_Area"" ss:RefersTo=""='SCM AV data'!R1C1:R" + CStr(iRowCount) + "C15""/><NamedRange ss:Name=""Print_Titles"" ss:RefersTo=""='SCM AV data'!C1:C3,'SCM AV data'!R4""/></Names>"
sHeader = sHeader & "<Table x:FullColumns=""1"" x:FullRows=""1"">"
sHeader = sHeader & "<Column ss:Width=""93.75""/><Column ss:Width=""61.5""/><Column ss:Width=""164.25""/><Column ss:Width=""165.75""/><Column ss:Width=""54""/><Column ss:Width=""55""/><Column ss:Width=""55""/><Column ss:Width=""153""/><Column ss:Width=""41.25""/><Column ss:Width=""34.5""/><Column ss:Width=""42""/><Column ss:Width=""42.75""/><Column ss:Width=""90"" ss:Span=""1""/><Column ss:Width=""80""/><Column ss:Index=""16"" ss:Width=""16.5""/>"
response.Write sHeader


Dim iCatId, bEmptyCat, dtLastUpdate

Do Until rs.EOF

	If iCatID <> rs("FeatureCategoryID") Then
		iCatID = rs("FeatureCategoryID")
		'
		' Category Row
		'
		If bEmptyCat Then
			sContacts = sContacts & "<Row></Row><Row></Row>"
		ElseIf iCatID <> 1 Then
			sContacts = sContacts & "<Row></Row>"
		End If		    
		Response.Write sContacts
		
        dtLastUpdate = rs("ConfigRules_LastUpdDt")
        If Not IsDate(dtLastUpdate) Then
            dtLastUpdate = dtLastPublishDt
        End If

        If CDate(dtLastUpdate) > CDate(dtLastPublishDt) Then
		sContacts = "<Row>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s67""><Data ss:Type=""String"">" & Trim(XmlSafe(rs("AvFeatureCategory") & "")) & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & Trim(XmlSafe(rs("CategoryMarketingDescription") & "")) & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("CategoryRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
        Else
		sContacts = "<Row>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s67""><Data ss:Type=""String"">" & Trim(XmlSafe(rs("AvFeatureCategory") & "")) & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & Trim(XmlSafe(rs("CategoryMarketingDescription") & "")) & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("CategoryRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/></Cell></Row>"
         End If
'		    "<Cell ss:StyleID=""s63""><Data ss:Type=""String"">" & Trim(XmlSafe(rs("CategoryRules") & "")) & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _

		bEmptyCat = True

	End If

	If rs("AvNo") & "" <> "" Then
	
		bEmptyCat = False
		
		
		If rs("Status") = "A" And rs("bStatus") & "" = "1" Then
		'
		' Av Row
		'
		sContacts = sContacts & "<Row>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("GPGDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("MarketingDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & iProgramVersion & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("CPLBlindDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("RASDiscontinueDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("ConfigRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & ConvertYN( rs("IdsSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & ConvertYN( rs("IdsCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & ConvertYN( rs("RctoSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & ConvertYN( rs("RctoCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("UPC") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s32""><Data ss:Type=""String"">" & XmlSafe(rs("GSEndDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>"

		ElseIf rs("Status") = "A" Then
		'
		' Av Row
		'
		sContacts = sContacts & "<Row>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bAvNo")) & """><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s31""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bGPGDescription")) & """><Data ss:Type=""String"">" & XmlSafe(rs("GPGDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bMarketingDescription")) & """><Data ss:Type=""String"">" & XmlSafe(rs("MarketingDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bAvNo")) & """><Data ss:Type=""String"">" & iProgramVersion & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bCPLBlindDt")) & """><Data ss:Type=""String"">" & XmlSafe(rs("CPLBlindDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bRASDiscontinueDt")) & """><Data ss:Type=""String"">" & XmlSafe(rs("RASDiscontinueDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bConfigRules")) & """><Data ss:Type=""String"">" & XmlSafe(rs("ConfigRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bIdsSkus")) & """><Data ss:Type=""String"">" & ConvertYN( rs("IdsSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bIdsCto")) & """><Data ss:Type=""String"">" & ConvertYN( rs("IdsCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bRctoSkus")) & """><Data ss:Type=""String"">" & ConvertYN( rs("RctoSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bRctoCto")) & """><Data ss:Type=""String"">" & ConvertYN( rs("RctoCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bUPC")) & """><Data ss:Type=""String"">" & XmlSafe(rs("UPC") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bWeight")) & """><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""" & GetCellStyle(rs("bGSEndDt")) & """><Data ss:Type=""String"">" & XmlSafe(rs("GSEndDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>"
		    
		ElseIf rs("Status") = "O" And  rs("bStatus") & "" = "1"  Then
		'
		' Av Row Strikethrough Highlight
		'
		sContacts = sContacts & "<Row ss:StyleID=""s71"">" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("GPGDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("MarketingDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & iProgramVersion & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("CPLBlindDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("RASDiscontinueDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("ConfigRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & ConvertYN( rs("IdsSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & ConvertYN( rs("IdsCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & ConvertYN( rs("RctoSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & ConvertYN( rs("RctoCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("UPC") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s74""><Data ss:Type=""String"">" & XmlSafe(rs("GSEndDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>"
		ElseIf rs("Status") = "O" Then
		'
		' Av Row Strikethrough
		'
		sContacts = sContacts & "<Row ss:StyleID=""s71"">" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("AvNo")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("GPGDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/><NamedCell ss:Name=""Print_Titles""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("MarketingDescription")&"") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & iProgramVersion & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("CPLBlindDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("RASDiscontinueDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("ConfigRules") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & ConvertYN( rs("IdsSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & ConvertYN( rs("IdsCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & ConvertYN( rs("RctoSkus_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & ConvertYN( rs("RctoCto_YN") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("UPC") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><NamedCell ss:Name=""Print_Area""/></Cell>" & _
		    "<Cell ss:StyleID=""s72""><Data ss:Type=""String"">" & XmlSafe(rs("GSEndDt") & "") & "</Data><NamedCell ss:Name=""Print_Area""/></Cell>"
		End If
		
		'
		' Populate the Mr Dates
		'
		sContacts = sContacts & "<Cell ss:Index=""17"" ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AC4")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABL")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABC")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABM")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ACH")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABA")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABV")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("UUG")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKB")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABY")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABB")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABF")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABD")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("B1A")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKC")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("A2M")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABT")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABZ")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABH")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABN")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKD")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB9")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ACB")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKR")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKN")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ACQ")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABE")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AK8")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("UUZ")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB8")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABU")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("UUF")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABG")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB5")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ACJ")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ABJ")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("ACF")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB1")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("B1L")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB2")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AB0")&"") & "</Data></Cell>"
		sContacts = sContacts & "<Cell ss:StyleID=""s31""><Data ss:Type=""String"">" & XmlSafe(rs("AKL")&"") & "</Data></Cell>"
		sContacts = sContacts & "</Row>"	
	End If
	rs.MoveNext
Loop

rs.Close

sContacts = sContacts & "<Row><Cell ss:Index=""8"" ss:StyleID=""s64""/></Row>"
sContacts = sContacts & "<Row><Cell ss:Index=""8"" ss:StyleID=""s64""/></Row>"
sContacts = sContacts & "<Row ss:StyleID=""s80""><Cell ss:StyleID=""s80"" ss:Index=""2""><Data ss:Type=""String"">Items below this line are for US Retail Web/Kiosk only</Data><NamedCell ss:Name=""Print_Titles""/></Cell></Row>"
sContacts = sContacts & "</Table>"
sContacts = sContacts & "<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">"
sContacts = sContacts & "<PageSetup><Layout x:Orientation=""Landscape""/><Footer x:Margin=""0.17""/><PageMargins x:Bottom=""0.42"" x:Top=""0.44""/></PageSetup>"
sContacts = sContacts & "<Unsynced/><FitToPage/><Print><FitHeight>48</FitHeight><ValidPrinterInfo/><PaperSizeIndex>5</PaperSizeIndex><Scale>58</Scale><HorizontalResolution>600</HorizontalResolution><VerticalResolution>600</VerticalResolution></Print>"
sContacts = sContacts & "<Zoom>75</Zoom><FreezePanes/><SplitHorizontal>5</SplitHorizontal><TopRowBottomPane>5</TopRowBottomPane><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane><ActivePane>0</ActivePane>"
sContacts = sContacts & "<Panes><Pane><Number>3</Number><ActiveRow>1</ActiveRow></Pane><Pane><Number>1</Number></Pane><Pane><Number>2</Number></Pane><Pane><Number>0</Number><ActiveRow>1</ActiveRow><ActiveCol>1</ActiveCol><RangeSelection>R1C1:R1C1</RangeSelection></Pane></Panes>"
sContacts = sContacts & "<ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>"
Response.Write sContacts
'
' Begin SCM Platform Data Worksheet
'
Set cmd = dw.CreateCommandSP(cn, "usp_SelectKmat")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID") 
Set rs = dw.ExecuteCommandReturnRS(cmd)

sContacts = "<Worksheet ss:Name=""SCM Platform data"">"
sContacts = sContacts & "<Table x:FullColumns=""1"">"
sContacts = sContacts & "<Column ss:Width=""120.75""/><Column ss:AutoFitWidth=""0"" ss:Width=""72""/><Column ss:AutoFitWidth=""0"" ss:Width=""85.5""/><Column ss:AutoFitWidth=""0"" ss:Width=""83.25""/>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""26.25""><Cell ss:StyleID=""s21""><Data ss:Type=""String"">" & scmName & "</Data></Cell></Row>"
sContacts = sContacts & "<Row ss:AutoFitHeight=""0"" ss:Height=""15.75""><Cell ss:StyleID=""s22""><Data ss:Type=""String"">Platform Information</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s65""><Data ss:Type=""String"">KMAT Part number</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">" & rs("KMAT") & "</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s65""><Data ss:Type=""String"">EZC KMAT Part number</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">" & rs("EZCKMAT") & "</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s65""><Data ss:Type=""String"">Project Code Number</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">" & rs("ProjectCd") & "</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s65""><Data ss:Type=""String"">Sales Orgs</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">" & rs("SalesOrg") & "</Data></Cell></Row>"
sContacts = sContacts & "<Row><Cell ss:StyleID=""s65""><Data ss:Type=""String"">Plant Code</Data></Cell><Cell ss:StyleID=""s23""><Data ss:Type=""String"">" & rs("PlantCd") & "</Data></Cell></Row>"
sContacts = sContacts & "</Table></Worksheet>"
sContacts = sContacts & "</Workbook>"

rs.Close

Response.Write sContacts
	

If m_EditModeOn And Request("Publish") = "True" Then
	Set cmd = dw.CreateCommandSP(cn, "usp_SetScmPublishDt")
	cmd.NamedParameters = True
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	dw.CreateParameter cmd, "@p_UserName", adVarChar, adParamInput, 50, m_UserFullName
	cmd.Execute
End If
				
Set rs = nothing
Set cn = nothing
Set cmd = nothing
Set dw = nothing
%>

<%@ Language=VBScript %>

<%
Option Explicit
Response.Buffer = True

	
%>
<!-- #include file="../includes/ExcelExport.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../agency/excel_export.asp" -->
<!-- #include file="../countries/excel_export.asp" -->
<!-- #include file="pdd_export_xlstyle.asp" -->
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
Dim dtPddLock
Dim iSheetRowCount
Dim sText
Dim sLanguages
Dim sBrand
Dim sRegion
Dim sCategory
Dim sRootName

dtYear = DatePart("yyyy", Now())
dtMonth = DatePart("m", Now())
dtDay = DatePart("d", Now())

If len(dtMonth) = 1 Then
	dtMonth = "0" & dtMonth
End If
If Len(dtDay) = 1 Then
	dtDay = "0" & dtDay
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

'
' Get Platform Name
'
Set cmd = dw.CreateCommandSP(cn, "spGetProductVersionName")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

platformName = rs("name")
fileName = Replace(platformName, " ", "_") & "_PDD_" & dtYear & dtMonth & dtDay & ".xls"

rs.Close

'
' Get Pdd Lock Date
'
Set cmd = dw.CreateCommandSP(cn, "usp_GetPddLockStatus")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

if rs.EOF and rs.BOF then
    response.Write "<html><head><title>PDD Export Error</title></head><body><h2>PDD Export Error</h2><p>No schedules have been created for this product.  The schedule is required to provide this report with the PRL/PDD Lock date.</p></body></html>"
    response.End
else
    dtPddLock = rs("PDD_Locked_Dt")
end if

rs.Close

xlDoc = OpenXlDoc(excelStyle)


' ****************************************
' Schedule Worksheet
' ****************************************

'
' Get Schedule Data
'
Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData_PDD_Export_Schedule")
dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_ActiveYN", adChar, adParamInput, 1, "Y"
dw.CreateParameter cmd, "@p_ScheduleDefinitionDataID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

sRegion = ""
sBrand = ""
Do Until rs.EOF

	If rs("Schedule_ID") <> sBrand Then
		If sBrand <> "" Then
			xlDoc = xlDoc & OpenXlRow("")
			xlDoc = xlDoc & AddXlCell("", "Bot")
			xlDoc = xlDoc & AddXlCell("", "Default")
			xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
			xlDoc = xlDoc & AddXlCell("", "Default")
			xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
			xlDoc = xlDoc & AddXlCell("", "Default")
			xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
			xlDoc = xlDoc & AddXlCell("", "Default")
			xlDoc = xlDoc & AddXlCell("", "Bot")
			xlDoc = xlDoc & CloseXlRow()

            xlDoc = xlDoc & OpenXlRow("")
            xlDoc = xlDoc & AddXlCell("Exported on " & NOW(), "Default")
            xlDoc = xlDoc & CloseXlRow()

			xlDoc = xlDoc & CloseXlTable()
			xlDoc = xlDoc & CloseXlWorksheet()
		End If
		sBrand = rs("Schedule_ID")
		xlDoc = xlDoc & OpenXlWorksheet(Replace(Left(rs("schedule_name"),31), "/", " "))
		xlDoc = xlDoc & OpenXlTable()
		xlDoc = xlDoc & AddColumn(200)
		xlDoc = xlDoc & AddColumn(5)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(5)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(5)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(55)
		xlDoc = xlDoc & AddColumn(5)
		xlDoc = xlDoc & AddColumn(150)
		xlDoc = xlDoc & AddColumn(5)
		xlDoc = xlDoc & AddColumn(60)
		xlDoc = xlDoc & OpenXlRow("7.5")
		xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 14)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("24")
		xlDoc = xlDoc & AddXlCellSpan(rs("family_name") & " " & rs("schedule_name"), "TitleText", 14)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("7.5")
		xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 14)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("7.5") & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow(30)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlHtmlCell("<B><Font html:Color=""#FFFFFF"">POR&#10;</Font><Font html:Color=""#FF0000"" html:Size=""8"">For Reference Only</Font></B>", "HdrSingleCenter", 2)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCellSpan("Current Commitment", "HdrSingleCenter", 2)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCellSpan("Actual", "HdrSingleCenter", 2)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("Start", "HdrLeftCenter")
		xlDoc = xlDoc & AddXlCell("Finish", "HdrRightCenter")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("Start", "HdrLeftCenter")
		xlDoc = xlDoc & AddXlCell("Finish", "HdrRightCenter")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("Start", "HdrLeftCenter")
		xlDoc = xlDoc & AddXlCell("Finish", "HdrRightCenter")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("Comments", "HdrSingle")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell("Owner", "HdrSingle")
		xlDoc = xlDoc & CloseXlRow()
	End If

	If sRegion <> rs("phase_name") Then
		sRegion = rs("phase_name")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(rs("phase_name"), "CellRegion", 14)
		xlDoc = xlDoc & CloseXlRow()
	End If

	If Trim(UCase(rs("milestone_yn"))) = "Y" Then
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell(rs("item_description"), "CellSingle")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCellSpan(rs("por_start_dt"), "CellSingleCentered",2)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCellSpan(rs("projected_start_dt"), "CellSingleCentered",2)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCellSpan(rs("actual_start_dt"), "CellSingleCentered",2)
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("item_notes"), "CellSingle")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("RoleShortName"), "CellSingle")
		xlDoc = xlDoc & CloseXlRow()
	Else
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell(rs("item_description"), "CellSingle")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("por_start_dt"), "CellLeftCentered")
		xlDoc = xlDoc & AddXlCell(rs("por_end_dt"), "CellRightCentered")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("projected_start_dt"), "CellLeftCentered")
		xlDoc = xlDoc & AddXlCell(rs("projected_end_dt"), "CellRightCentered")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("actual_start_dt"), "CellLeftCentered")
		xlDoc = xlDoc & AddXlCell(rs("actual_end_dt"), "CellRightCentered")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("item_notes"), "CellSingle")
		xlDoc = xlDoc & AddXlCell("", "Default")
		xlDoc = xlDoc & AddXlCell(rs("RoleShortName"), "CellSingle")
		xlDoc = xlDoc & CloseXlRow()
	End If
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & AddXlCell("", "Default")
xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
xlDoc = xlDoc & AddXlCell("", "Default")
xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
xlDoc = xlDoc & AddXlCell("", "Default")
xlDoc = xlDoc & AddXlCellSpan("", "Bot",2)
xlDoc = xlDoc & AddXlCell("", "Default")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Exported on " & NOW(), "Default")
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

xlDoc = xlDoc & CloseXlDoc()
	
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
Response.ContentType = "application/vnd.ms-excel"
Response.Write xlDoc
				
Set rs = nothing
Set cn = nothing
Set cmd = nothing
Set dw = nothing
%>

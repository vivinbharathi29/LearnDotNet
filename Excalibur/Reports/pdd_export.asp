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
' Change Log Worksheet
' ****************************************

xlDoc = xlDoc & OpenXlWorksheet("Change Log")
iSheetRowCount = 0
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(300)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Change Log", "TitleText", 4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & CloseXlRow()

'
' Get Change Requests Opened
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Change Requests Opened", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Submitter", "Hdr")
xlDoc = xlDoc & AddXlCell("DCR No.", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = 6

Set cmd = dw.CreateCommandSP(cn, "spListDCRThisWeek")
dw.CreateParameter cmd, "@Report", adInteger, adParamInput, 8, 2
dw.CreateParameter cmd, "@Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID")
dw.CreateParameter cmd, "@StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)


rs.Sort = rs.Fields(2).Name & " desc"
Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(2), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("Submitter"), InStr(rs("Submitter"), ",") + 2) & ".", "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("ID"), "Cell")
	xlDoc = xlDoc & AddXlCell(rs("Summary"), "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot",4)
xlDoc = xlDoc & CloseXlRow()

'
' Get Change Requests Closed
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Change Requests Closed", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Submitter", "Hdr")
xlDoc = xlDoc & AddXlCell("DCR No.", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "spListDCRThisWeek")
dw.CreateParameter cmd, "@Report", adInteger, adParamInput, 8, 3
dw.CreateParameter cmd, "@Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID")
dw.CreateParameter cmd, "@StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = rs.Fields(2).Name & " desc"
Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(2), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("Submitter"), InStr(rs("Submitter"), ",") + 2) & ".", "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("ID"), "Cell")
	xlDoc = xlDoc & AddXlCell(rs("Summary"), "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

'
' Get Deliverable Matrix Changes
'
'rs.Open "spListDeliverableMatrixUpdates " & clng(strproductID) & ",2," & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Deliverable Matrix Changes", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Changed By", "Hdr")
xlDoc = xlDoc & AddXlCell("Change", "Hdr")
xlDoc = xlDoc & AddXlCell("Deliverable", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "spListDeliverableMatrixUpdates")
dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
dw.CreateParameter cmd, "@Type", adInteger, adParamInput, 8, 2
dw.CreateParameter cmd, "@Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "Updated desc"
Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("updated"), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("username"), InStr(rs("username"), ",") + 2) & ".", "Cell")
	If ucase(trim(rs("Action")&"")) = ucase("Targeted") Then
		xlDoc = xlDoc & AddXlCell("Added", "Cell")
	ElseIf ucase(trim(rs("Action")&"")) = ucase("Target Removed") Then
		xlDoc = xlDoc & AddXlCell("Removed", "Cell")
	Else
		xlDoc = xlDoc & AddXlCell("", "Cell")
	End If
	xlDoc = xlDoc & AddXlCell(rs("Deliverable") & " " & rs("version") & "," & rs("revision") & "," & rs("pass"), "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

'
' Get Schedule Changes
'
'rs.Open "usp_SelectScheduleDataHistory " & strProductID & ", " & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Schedule Changes", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Changed By", "Hdr")
xlDoc = xlDoc & AddXlCell("Schedule", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleDataHistory")
dw.CreateParameter cmd, "@p_ProductID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
dw.CreateParameter cmd, "@p_Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@p_StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@p_EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "last_upd_date desc"
Do Until rs.EOF
	If rs("old_actual_end_dt") & "" = "" And _
		rs("new_actual_end_dt") & "" <> "" And _
		rs("schedule_definition_data_id") = 7 Then
		
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & DrawXlCell(rs.Fields("last_upd_date"), "CellLeft")
		xlDoc = xlDoc & AddXlCell(Left(rs("user_name"), InStr(rs("user_name"), ",") + 2) & ".", "Cell")
		xlDoc = xlDoc & AddXlCell(rs("schedule_name"), "Cell")
		xlDoc = xlDoc & AddXlCell(Trim(rs("item_description") & "") & " achieved on " & FormatDateTime(rs("new_actual_end_dt") & "", vbShortDate), "CellRight")
		xlDoc = xlDoc & CloseXlRow()
	End If

	If rs("old_actual_end_dt") & "" = "" And _
		rs("new_actual_end_dt") & "" = "" Then
		
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & DrawXlCell(rs.Fields("last_upd_date"), "CellLeft")
		xlDoc = xlDoc & AddXlCell(Left(rs("user_name"), InStr(rs("user_name"), ",") + 2) & ".", "Cell")
		xlDoc = xlDoc & AddXlCell(rs("schedule_name"), "Cell")

		If UCASE(Trim(rs("milestone_yn"))) = "Y" Then
		
			sText = "The date for " & rs("item_description")
			If Trim(rs("old_projected_end_dt")&"") = "" Then
				sText = sText & " was set to "
				If rs("new_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_end_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
			Else
				sText = sText & " changed from "
				If rs("old_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("old_projected_end_dt")&"", vbShortDate) 
				Else
					sText = sText & "No Date"
				End If
				sText = sText & " to "
				If rs("new_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_end_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
			End If
		Else
		
			sText = "The dates for " & rs("item_description")
			If Trim(rs("old_projected_start_dt")&"") = "" Then
				sText = sText & " were set to Start:" 
				If rs("new_projected_start_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_start_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
				sText = sText & " Finish:"
				If rs("new_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_end_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
			Else
				sText = sText & " changed from Start:"
				If rs("old_projected_start_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("old_projected_start_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
				sText = sText & " Finish:"
				If rs("old_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("old_projected_end_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
				sText = sText & " to Start:"
				If rs("new_projected_start_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_start_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
				sText = sText & " Finish:"
				If rs("new_projected_end_dt")&"" <> "" Then
					sText = sText & FormatDateTime(rs("new_projected_end_dt")&"", vbShortDate)
				Else
					sText = sText & "No Date"
				End If
			End If
		End If
		xlDoc = xlDoc & AddXlCell(sText, "CellRight")
		xlDoc = xlDoc & CloseXlRow()
		iSheetRowCount = iSheetRowCount + 1
	End If
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

'
' Agency Changes
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Agency Changes", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Changed By", "Hdr")
xlDoc = xlDoc & AddXlCell("Schedule", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyStatusHistory")
dw.CreateParameter cmd, "@p_ProductID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, NULL
dw.CreateParameter cmd, "@p_Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@p_StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@p_EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "date_of_change desc"
Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("date_of_change"), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("changed_by"), InStr(rs("changed_by"), ",") + 2) & ".", "Cell")
	xlDoc = xlDoc & AddXlCell(rs("change_type"), "Cell")
	xlDoc = xlDoc & AddXlCell(rs("change_summary"), "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

'
' Supported Country Changes
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Supported Country Changes", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Changed By", "Hdr")
xlDoc = xlDoc & AddXlCell("DCR", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "usp_SelectProdBrandCountryHistory")
dw.CreateParameter cmd, "@p_ProductID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
dw.CreateParameter cmd, "@p_Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@p_StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@p_EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "last_upd_date desc"
Do Until rs.EOF
	If LCase(rs("ChangeType")) = "added" Then
		sText = rs("Country") & " was " & LCase(rs("ChangeType")) & " to the " & rs("Brand") & " brand."
	Else
		sText = rs("Country") & " was " & LCase(rs("ChangeType")) & " from the " & rs("Brand") & " brand."
	End If

	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("Last_Upd_Date"), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("Last_Upd_User"), InStr(rs("Last_Upd_User"), ",") + 2) & ".", "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("DcrID"), "Cell")
	xlDoc = xlDoc & AddXlCell(sText, "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

'
' Localization Changes
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Localization Changes", "HdrSingle", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Date", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Changed By", "Hdr")
xlDoc = xlDoc & AddXlCell("DCR", "Hdr")
xlDoc = xlDoc & AddXlCell("Summary", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = iSheetRowCount + 3

Set cmd = dw.CreateCommandSP(cn, "usp_SelectProdBrandCountryLocalizationHistory")
dw.CreateParameter cmd, "@p_ProductID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
dw.CreateParameter cmd, "@p_Duration", adInteger, adParamInput, 8, 0
dw.CreateParameter cmd, "@p_StartDt", adDate, adParamInput, 8, dtPddLock
dw.CreateParameter cmd, "@p_EndDt", adDate, adParamInput, 8, Now()
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "last_upd_date desc"
Do Until rs.EOF

	sText = rs("OptionConfig") & "/" & rs("Dash") & " was " & LCase(rs("ChangeType")) & " for " & rs("Country") & " on " & rs("Brand") & "."

	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("Last_Upd_Date"), "CellLeft")
	xlDoc = xlDoc & AddXlCell(Left(rs("Last_Upd_User"), InStr(rs("Last_Upd_User"), ",") + 2) & ".", "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("DcrID"), "Cell")
	xlDoc = xlDoc & AddXlCell(sText, "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Scope Changes Worksheet
' ****************************************

xlDoc = xlDoc & OpenXlWorksheet("Scope Changes")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(300)
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",5)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Scope Changes", "TitleText", 5)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",5)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & CloseXlRow()

'
' Get Approved DCRs
'
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("Approved DCRs", "HdrSingle", 5)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("Number", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Status", "Hdr")
xlDoc = xlDoc & AddXlCell("Approval Dt.", "Hdr")
xlDoc = xlDoc & AddXlCell("Description", "Hdr")
xlDoc = xlDoc & AddXlCell("Submitter", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

iSheetRowCount = 5

Set cmd = dw.CreateCommandSP(cn, "spListApprovedDCRs")
dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

rs.Sort = "id desc"
Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs("ID"), "CellLeft")
	xlDoc = xlDoc & AddXlCell(rs("Status"), "Cell")
	xlDoc = xlDoc & DrawXlCell(rs("ApprovalDate"), "Cell")
	xlDoc = xlDoc & AddXlCell(rs("Summary"), "Cell")
	xlDoc = xlDoc & AddXlCell(Left(rs("Submitter"), InStr(rs("Submitter"), ",") + 2) & ".", "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 5)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Schedule Worksheet
' ****************************************

'
' Get Schedule Data
'
Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData")
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

			xlDoc = xlDoc & CloseXlTable()
			xlDoc = xlDoc & CloseXlWorksheet()
		End If
		sBrand = rs("Schedule_ID")
		xlDoc = xlDoc & OpenXlWorksheet(rs("schedule_name"))
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
		xlDoc = xlDoc & OpenXlRow("7.5")
		xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 12)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("24")
		xlDoc = xlDoc & AddXlCellSpan(rs("family_name") & " " & rs("schedule_name"), "TitleText", 12)
		xlDoc = xlDoc & CloseXlRow()
		xlDoc = xlDoc & OpenXlRow("7.5")
		xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder", 12)
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
		xlDoc = xlDoc & CloseXlRow()
	End If

	If sRegion <> rs("phase_name") Then
		sRegion = rs("phase_name")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(rs("phase_name"), "CellRegion", 12)
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

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Requirements Worksheet
' ****************************************
xlDoc = xlDoc & OpenXlWorksheet("General Requirements")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " General Requirements", "TitleText",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Agency Worksheet
' ****************************************
xlDoc = xlDoc & OpenXlWorksheet("Agency")
xlDoc = xlDoc & OpenXlTable()

xlDoc = xlDoc & DrawAgencyMatrix(Request("PVID"), platformName)

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Product Localization Worksheet
' ****************************************
xlDoc = xlDoc & OpenXlWorksheet("Product Localization")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & AddColumn(50)
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & AddColumn(60)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",11)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Product Localization Matrix", "TitleText",11)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",11)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & CloseXlRow()

'
' Get Product Localization
'

iSheetRowCount = 5

Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductBrandLocalizations")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)

Do Until rs.EOF
	If sBrand <> rs("brand") Then
		If sBrand <> "" Then
			xlDoc = xlDoc & OpenXlRow("")
			xlDoc = xlDoc & AddXlCellSpan("", "Bot", 11)
			xlDoc = xlDoc & CloseXlRow()
		End If
		sBrand = rs("brand")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(sBrand, "SubTitle", 11)
		xlDoc = xlDoc & CloseXlRow()

		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell("Country Region", "HdrLeft")
		xlDoc = xlDoc & AddXlCell("HP Blue Code", "Hdr")
		xlDoc = xlDoc & AddXlCell("CPQ Red Dash", "Hdr")
		xlDoc = xlDoc & AddXlCell("Languages", "Hdr")
		xlDoc = xlDoc & AddXlCell("MUI", "Hdr")
		xlDoc = xlDoc & AddXlCell("Country Cd.", "Hdr")
		xlDoc = xlDoc & AddXlCell("Keyboard", "Hdr")
		xlDoc = xlDoc & AddXlCell("KWL", "Hdr")
		xlDoc = xlDoc & AddXlCell("Power Cord", "Hdr")
		xlDoc = xlDoc & AddXlCell("Doc. Languages", "Hdr")
		xlDoc = xlDoc & AddXlCell("OS Restore Solution", "HdrRight")
		xlDoc = xlDoc & CloseXlRow()
	End If
	
	If sRegion <> rs("GeoID") Then
		sRegion = rs("GeoID")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(rs("Geo"), "CellRegion", 11)
		xlDoc = xlDoc & CloseXlRow()
	End If

	sLanguages =  rs("OSLanguage") & ""
	if trim(rs("OtherLanguage") & "") <> "" then
		sLanguages = sLanguages & "," & rs("OtherLanguage")
	end if	
	
	If IsNull(rs("NonStandard")) Then
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell(rs("CountryRegion"), "CellLeft")
		xlDoc = xlDoc & AddXlCell(rs("OptionConfig"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("Dash"), "Cell")
		xlDoc = xlDoc & AddXlCell(sLanguages, "Cell")
		xlDoc = xlDoc & AddXlCell(rs("MUI"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("CountryCode"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("Keyboard"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("KWL"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("PowerCord"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("DocKits"), "Cell")
		xlDoc = xlDoc & AddXlCell(rs("RestoreMedia"), "CellRight")
		xlDoc = xlDoc & CloseXlRow()
	Else
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCell(rs("CountryRegion"), "CellLeftHighlight")
		xlDoc = xlDoc & AddXlCell(rs("OptionConfig"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("Dash"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(sLanguages, "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("MUI"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("CountryCode"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("Keyboard"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("KWL"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("PowerCord"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("DocKits"), "CellHighlight")
		xlDoc = xlDoc & AddXlCell(rs("RestoreMedia"), "CellRightHighlight")
		xlDoc = xlDoc & CloseXlRow()
	End If
	iSheetRowCount = iSheetRowCount + 1
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 11)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Image Localization Worksheet
' ****************************************
xlDoc = xlDoc & OpenXlWorksheet("Image Localization")
xlDoc = xlDoc & OpenXlTable()

xlDoc = xlDoc & DrawImageLocalizationMatrix(Request("PVID"), platformName)

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

' ****************************************
' Software Worksheet
' ****************************************
xlDoc = xlDoc & OpenXlWorksheet("Software")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(200)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Software", "TitleText",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & CloseXlRow()

Set cmd = dw.CreateCommandSP(cn, "rpt_pdd_Software")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID") 'Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)	

xlDoc = xlDoc & DrawXlHeaderRow(rs)

Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(0), "CellLeft")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(1), "CellCentered")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(2), "CellCentered")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(3), "CellRightCentered")
	xlDoc = xlDoc & CloseXlRow()
	rs.MoveNext
Loop

rs.Close
	
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot",4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

xlDoc = xlDoc & OpenXlWorksheet("Commodities")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(150)
xlDoc = xlDoc & AddColumn(150)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Commodities", "TitleText",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & CloseXlRow()


Set cmd = dw.CreateCommandSP(cn, "usp_pdd_SelectCommodities")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)	

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("ID", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Supplier", "Hdr")
xlDoc = xlDoc & AddXlCell("Model No.", "Hdr")
xlDoc = xlDoc & AddXlCell("POR/DCR", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

Do Until rs.EOF
	If sCategory <> rs("Category") Then
		sCategory = rs("Category")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(sCategory, "HdrSingle", 4)
		xlDoc = xlDoc & CloseXlRow()
	End If
	If sRootName <> rs("RootName") Then
		sRootName = rs("RootName")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(sRootName, "CellRegion", 4)
		xlDoc = xlDoc & CloseXlRow()
	End If
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("VersionID"), "CellLeft")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("Vendor"), "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("ModelNumber"), "Cell")
	Select Case rs("DCRID")
		Case 1
			xlDoc = xlDoc & AddXlCell("POR", "CellRight")
		Case 2
			xlDoc = xlDoc & AddXlCell("HFCN", "CellRight")
		Case Else
			xlDoc = xlDoc & AddXlCell("DCR: " & rs("DCRID"), "CellRight")
	End Select
	xlDoc = xlDoc & CloseXlRow()
	rs.MoveNext
Loop

rs.Close

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()

xlDoc = xlDoc & OpenXlWorksheet("Accessories")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & AddColumn(100)
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("24")
xlDoc = xlDoc & AddXlCellSpan(platformName & " Accessories", "TitleText",4)
xlDoc = xlDoc & CloseXlRow()
xlDoc = xlDoc & OpenXlRow("7.5")
xlDoc = xlDoc & AddXlCellSpan("", "TitleBorder",4)
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

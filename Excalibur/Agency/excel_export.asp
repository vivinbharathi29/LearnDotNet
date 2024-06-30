<%
Function BuildRow(RecordSet)
    Dim rowData, strRow
    Dim statusID, agencyType, agencyName, statusCd, projectedDt, levStatusCd, levProjectedDt
    Dim supportedCountryYN, statusNote, porDcr, pcid
    Dim field

	If RecordSet.Fields(3).Value & "" = "" Then
		supportedCountryYN = "N"
	Else
		supportedCountryYN = "Y"
	End If
	
    strRow = OpenXlRow("")

    'strRow = AddCell(strRow, RecordSet("Country"))

    For Each field In RecordSet.Fields

        If field.name = RecordSet.Fields(0).name Then
			'Do Nothing
        ElseIf field.name = RecordSet.Fields(1).name Then
            strRow = strRow & AddCell(field.value)
        ElseIf field.name = RecordSet.Fields(2).name Then
            'Do Nothing
        ElseIf field.name = RecordSet.Fields(3).name Then
            'Do Nothing
        ElseIf field.value & "" = "" Then
			strRow = strRow & AddCell("")
		Else
            rowData = Split(field.value, "|")

            statusID = rowData(0)
            agencyType = rowData(1)
            agencyName = rowData(2)
            statusCd = rowData(3)
            projectedDt = rowData(4)
            levStatusCd = rowData(5)
            levProjectedDt = rowData(6)
            supportedCountryYN = rowData(7)
            statusNote = rowData(8)
            porDcr = rowData(9)
            pcid = rowData(10)

            strRow = strRow & AddStatusCell(statusID, agencyType, agencyName, statusCd, projectedDt, levStatusCd, levProjectedDt, supportedCountryYN, statusNote, porDcr, pcid)
        End If

    Next

    strRow = strRow & AddXlCell("", "Right") & CloseXlRow()

    BuildRow = strRow

End Function

Function AddCell(Value)
    AddCell = AddXlCell(Value, "AgencyCell")
End Function

Function AddStatusCell(StatusID, AgencyType, AgencyName, StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote, PorDcr, PCID)
    Dim Value, LinkCode
    Dim CssClass

    LinkCode = ""
    Value = GetStatusValue(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote)

    Select Case AgencyType
        Case "C"
            CssClass = GetStatusClass(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, PCID)
            LinkCode = "LANGUAGE=javascript onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()' onclick='return Status_OnClick(" & StatusID & ")'"
        Case "I"
            CssClass = "Illegal"
            Value = AgencyName
        Case "E"
            CssClass = "Embargo"
            Value = AgencyName
        Case "N"
            CssClass = "NotNeeded"
            If PCID = "" Then
                CssClass = "NS_NotNeeded"
            End If
            Value = AgencyName
        Case Else
            CssClass = "Unknown"
            Value = AgencyName
    End Select

    If UCASE(PorDcr) = "DCR" Then
        Value = Value & " *"
    End If

    AddStatusCell = AddXlHtmlCell(Value, CssClass, 1)

End Function

Function GetStatusClass(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, PCID)
    Dim StatusClass

    Select Case UCase(StatusCd)
        Case "O"
            If isDate(ProjectedDt) Then
                If DateDiff("d", Now(), ProjectedDt) > 0 Then
                    StatusClass = "Cert_Open"
                Else
                    StatusClass = "Cert_Late"
                End If
            Else
                StatusClass = "Cert_Late"
            End If
        Case "SU"
            If isDate(ProjectedDt) Then
                If DateDiff("d", Now(), ProjectedDt) > 0 Then
                    StatusClass = "Cert_Open"
                Else
                    StatusClass = "Cert_Late"
                End If
            Else
                StatusClass = "Cert_Late"
            End If
        Case "C"
            StatusClass = "Cert_Complete"
        Case "NS"
            StatusClass = "NotSupported"
        Case "L"
            If LevStatusCd = "C" Then
                StatusClass = "Cert_Complete"
            ElseIf isDate(LevProjectedDt) Then
                If DateDiff("d", Now(), LevProjectedDt) > 0 Then
                    StatusClass = "Cert_Open"
                Else
                    StatusClass = "Cert_Late"
                End If
            Else
                StatusClass = "Cert_Late"
            End If
        Case Else
            StatusClass = "NotNeeded"
    End Select

    If Len(Trim(PCID)) = 0 Then
        StatusClass = "NS_" & StatusClass
    End If

    GetStatusClass = StatusClass
End Function

Function GetStatusValue(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote)

    GetStatusValue = "&nbsp;"

    Select Case UCase(StatusCd)
        Case "O"
            If isDate(ProjectedDt) Then
                GetStatusValue = FormatDateTime(ProjectedDt, vbShortDate)
            Else
                GetStatusValue = "OPEN"
            End If
        Case "SU"
            If isDate(ProjectedDt) Then
                GetStatusValue = FormatDateTime(ProjectedDt, vbShortDate)
            Else
                GetStatusValue = "OPEN"
            End If
        Case "C"
			GetStatusValue = "Certified"
        Case "L"
            GetStatusValue = "Leveraged&#10;" '& vbCrLf
            If LevStatusCd = "C" Then
				GetStatusValue = GetStatusValue & "Certified"
            Else
                If isDate(LevProjectedDt) Then
                    GetStatusValue = GetStatusValue & FormatDateTime(LevProjectedDt, vbShortDate)
                ElseIf LevStatusCd = "NS" Then
                    GetStatusValue = GetStatusValue & "Not Supported"
                Else
                    GetStatusValue = GetStatusValue & "OPEN"
                End If
            End If
        Case "NS"
            If UCase(SupportedCountryYN) = "Y" Then
                GetStatusValue = "Product&#10;" '& vbCrLf
            Else
                GetStatusValue = "Country&#10;" '& vbCrLf
            End If
            GetStatusValue = GetStatusValue & "Not Supported"
        Case "NR"
            GetStatusValue = "Not Requested"
        Case Else
            GetStatusValue = "Invalid Status"
    End Select

    If Len(Trim(StatusNote)) > 0 Then
        'GetStatusValue = GetStatusValue & vbcrlf & "&lt;Note&gt;"
    End If

End Function

Function DrawGroupingRow(Region, ColCount)
    DrawGroupingRow = OpenXlRow("") & AddXlCellSpan(Region, "Region", ColCount) & CloseXlRow()
End Function
	
Function DrawHeaderRow(RecordSet)
	Dim i
	Dim sOutput
	Dim asDeliverables
	Dim asDeliverable
	Dim sMasterUrl
	Dim sUrl
	Dim field
	
	
	sOutput = OpenXlRow("")
	If RecordSet.Fields.Count = 0 Then
		sOutput = sOutput & AddXlCell("There are no agency requirements recorded.", "Hdr")
		Exit Function
	End If
	For Each field IN RecordSet.Fields
		If field.Name = RecordSet.Fields(0).Name Then
			sOutput = sOutput & AddXlCell(field.value, "HdrLeft")
		ElseIf field.Name = RecordSet.Fields(RecordSet.Fields.Count-1).Name Then
			sOutput = sOutput & AddXlCell(field.value, "HdrRightCenter")
		Else
			sOutput = sOutput & AddXlCell(field.value, "HdrCenter")
		End If
	Next
	sOutput = sOutput & CloseXlRow()
	
	DrawHeaderRow = sOutput
End Function

Function DrawAgencyMatrix(ProductVersionID, PlatformName)
	Dim cn
	Dim cmd
	Dim dw
	Dim rs
	Dim i
	Dim sOutput

	Dim colCount
	Dim sRegion
	sRegion = ""

	'
	' Get Data
	'
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "rpt_AgencyPMView")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
	colCount = rs.Fields.Count
	sOutput = AddColumn(150)
	For i = 1 to colCount
		sOutput = sOutput & AddColumn(100)
	Next

	sOutput = sOutput & OpenXlRow("7.5")
	sOutput = sOutput & AddXlCellSpan("", "TitleBorder", colCount)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("24")
	sOutput = sOutput & AddXlCellSpan(PlatformName & " Agency", "TitleText", colCount)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("7.5")
	sOutput = sOutput & AddXlCellSpan("", "TitleBorder", colCount)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("7.5") & CloseXlRow()
	
	sOutput = sOutput & DrawHeaderRow(rs)
		
	If colCount = 0 Then
		Exit Function
	End If
		
	Set rs = rs.NextRecordset
	
	Do Until rs.EOF
		
		If sRegion <> rs("Region") Then
			sRegion = rs("Region")
			sOutput = sOutput & DrawGroupingRow(Replace(Replace(sRegion, "<",""),">",""), colCount)
		End If

		If rs.Fields(3).Value & "" <> "" Then
			sOutput = sOutput & BuildRow(rs)
		End If
			
		rs.MoveNext
	Loop
		
	rs.Close
				
	Set rs = nothing
	Set cn = nothing
	Set cmd = nothing
	Set dw = nothing
	
	DrawAgencyMatrix = sOutput & OpenXlRow("") & AddXlCellSpan("* Country added after POR by DCR", "Bot", colCount) & CloseXlRow()

End Function

%>

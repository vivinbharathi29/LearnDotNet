<%

Sub SendPDDLockNotification(ScheduleID, ProgramName, UserEmail, UserName)
	'
	' Query for the FCS date for the current schedule.  FCS Definition ID = 60
	'		
	Set dw = New DataWrapper
	set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData")
	dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, Request.Form("ScheduleID")
	dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, NULL
	dw.CreateParameter cmd, "@p_ScheduleDefinitionDataID", adInteger, adParamInput, 8, Application("FCS_ScheduleDefinitionID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
		
	If Not rs.EOF Then
		'
		' Send email to the MobileExcalNotification-FCS@hp.com list.
		'
		sSubject = ProgramName & " has reached PRL/PDD Lock"
		sBody = "The " & ProgramName & " schedule has locked." & vbcrlf & _
				vbcrlf & _
				"MV Production Release POR Date: " & FormatDateTime(rs("projected_end_dt"), vbShortDate)
		
		Set oMessage = Server.CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
	
		oMessage.To = "MobileExcalNotification-FCS@hp.com"
		oMessage.From = UserEmail
		oMessage.Subject = sSubject
		oMessage.TextBody = sBody
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send
		Set oMessage = Nothing 		
					
	End If
	
	rs.close
	
	set rs = nothing
	set cn = nothing
	set dw = nothing

End Sub
	
' If FCS Date changed and POR Dates Are Set send a message to the list.
'
Sub SendFCSNotification(Program, MilestoneName, PorEndDt, OldEndDt, NewEndDt, UserEmail, UserName)		
	'
	' Send email to the MobileExcalNotification-FCS@hp.com list.
	'
	
	sSubject = MilestoneName & " date change on " & Program
	sBody = "The " & MilestoneName & " date on the " & Program & " schedule has changed." & vbcrlf & _
			vbcrlf & _
			"POR Date: " & FormatDateTime(PorEndDt, vbShortDate) & vbcrlf & _
			"Old Date: " & FormatDateTime(OldEndDt, vbShortDate) & vbcrlf & _
			"New Date: " & FormatDateTime(NewEndDt, vbShortDate)
	
	Set oMessage = Server.CreateObject("CDO.Message")
	'Set oMessage.Configuration = Application("CDO_Config")
	
	oMessage.To = "MobileExcalNotification-FCS@hp.com"
	oMessage.From = UserEmail
	oMessage.Subject = sSubject
	oMessage.TextBody = sBody
	oMessage.DSNOptions = cdoDSNFailure
	oMessage.Send
	Set oMessage = Nothing 		

End Sub

Sub LogChangeToHistoryTable(Program, MilestoneName, OldStartDt, OldEndDt, NewStartDt, NewEndDt, UserName)

	Set dw = New DataWrapper
	set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_InsertScheduleDataHistory")
	dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, Request.Form("ScheduleID")
	dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, NULL
	dw.CreateParameter cmd, "@p_ScheduleDefinitionDataID", adInteger, adParamInput, 8, Application("FCS_ScheduleDefinitionID")

End Sub


%>
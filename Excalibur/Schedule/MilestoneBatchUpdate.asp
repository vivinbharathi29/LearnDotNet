<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/EmailWrapper.asp" -->
<!-- #include file = "clsSchedule.asp" -->

<%

'Response.Buffer = false

'##############################################################################	
'
' Create Security Object to get User Info
'

Dim m_IsSysAdmin
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_IsSEPMProductsEditor
Dim m_EditModeOn
Dim Security
Dim sUserFullName
Dim	m_UserEmail

	
Set Security = New ExcaliburSecurity

m_IsSysAdmin = Security.IsSysAdmin()

'
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
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If

m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
m_UserEmail = Security.CurrentUserEmail()

sUserFullName = Security.CurrentUser()

If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsSEPMProductsEditor Then
	m_EditModeOn = True
End If

If Not m_EditModeOn Then
	Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
	Response.End
End If

Set Security = Nothing

'##############################################################################	

Dim m_ScheduleID

Sub Main()
	m_ScheduleID = Request("ScheduleID")
	If m_ScheduleID = "" Then
		m_ScheduleID = 178
	End If

	If Request.Form("Mode") = "" Then
		Call Draw()
	Else
		Call Save()
	End If
End Sub

Sub Draw()

	Dim dw
	Dim cn
	Dim rs
	Dim cmd
	Dim sLastPhase
	Dim sActualStart
	Dim sActualEnd
	Dim sProjectedStart
	Dim sProjectedEnd
	Dim sDateStart
	Dim sDateEnd
	Dim sScheduleDataID
	Dim sItemNotes
    Dim sChangeNotes

	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData")

	dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, m_ScheduleID
	dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, "Y"
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Set cmd = Nothing

	Response.Write "<TABLE  ID=TableSchedule cellSpacing=1 cellPadding=1 width='100%' border=1 borderColor=tan bgColor=ivory>"
	Response.Write "<TR><TD nowrap width=100% bgColor=cornsilk vAlign=middle rowspan=2>" & _
		"<INPUT type=hidden name=Program id=Program value='" & rs("family_name") & " " & rs("schedule_name") & "'>" & _
		"<FONT size=1><STRONG>Schedule&nbsp;Item</STRONG></FONT></TD>"
	
	If Trim(Request("Mode")) = "Projected" Then
		Response.Write "<TD nowrap width=280 bgColor=cornsilk Align=center vAlign=middle colspan=2><FONT size=1><STRONG>Current Commitment</STRONG></FONT></TD>"
	Else
		Response.Write "<TD nowrap width=280 bgColor=cornsilk Align=center vAlign=middle colspan=2><FONT size=1><STRONG>Actual</STRONG></FONT></TD>"
	End If
	
	Response.Write "<TD nowrap width=100% bgColor=cornsilk rowspan=2><STRONG><FONT size=1>Change Notes</FONT></STRONG></TD></TR>"
	Response.Write "</TR>"
	
	Response.Write "<TR><TD nowrap width=140 bgColor=cornsilk Align=center vAlign=middle><FONT size=1><STRONG>Start</STRONG></FONT></TD>"
	Response.Write "<TD nowrap width=140 bgColor=cornsilk Align=center vAlign=middle><FONT size=1><STRONG>Finish</STRONG></FONT></TD></TR>"

	Do While Not rs.EOF  

		If sLastPhase <> rs("phase_name") Then
			Response.Write "<TR bgcolor=lightsteelblue><TD nowrap valign=top colspan=8><font color=black size=1 class='text'>" & rs("phase_name") & "</font></TD>"
			sLastPhase = rs("phase_name")
		End If
	
		Response.Write "<TR>"
		
		sScheduleDataID = rs("schedule_data_id")

		sActualStart = rs("actual_start_dt") & ""
		sActualEnd = rs("actual_end_dt") & ""

		sProjectedStart = rs("projected_start_dt") & ""
		sProjectedEnd = rs("projected_end_dt") & ""
		
		sItemNotes = rs("item_notes") & ""

        sChangeNotes = rs("change_notes") & ""
		
		If Request("Mode") = "Projected" Then
			If IsDate(sProjectedStart) Then
				sDateStart = FormatDateTime(sProjectedStart, 2)
			Else
				sDateStart = ""
			End If
			If IsDate(sProjectedEnd) Then
				sDateEnd = FormatDateTime(sProjectedEnd, 2)
			Else
				sDateEnd = ""
			End If
		ElseIf Request("Mode") = "Actual" Then
			If IsDate(sActualStart) Then
				sDateStart = FormatDateTime(sActualStart, 2)
			Else
				sDateStart = ""
			End If
			If IsDate(sActualEnd) Then
				sDateEnd = FormatDateTime(sActualEnd, 2)
			Else
				sDateEnd = ""
			End If
		End If
		
		Dim bIsMilestone
		If UCase(rs("milestone_yn")) = "Y" Then
			bIsMilestone = True
		Else
			bIsMilestone = False
		End If

		Response.Write "<TD valign=middle class='cell'><font size=1 class='text'>" & rs("item_description") & "</font></TD>"
		Response.Write "<TD valign=top class='cell' align=center bgcolor=cornsilk "
		If bIsMilestone Then Response.Write "ColSpan=2" 
		Response.Write " >"
		Response.Write "<INPUT type=hidden name=ScheduleDataID id=ScheduleDataID value='" & sScheduleDataID & "'>"
		Response.Write "<INPUT type=hidden name=ScheduleDefinitionDataID" & sScheduleDataID & " id=ScheduleDefinitionDataID" & sScheduleDataID & " value='" & rs("schedule_definition_data_id") & "'>"
		Response.Write "<INPUT type=hidden name=Milestone" & sScheduleDataID & " id=Milestone" & sScheduleDataID & " value='" & bIsMilestone & "'>"
		Response.Write "<INPUT type=hidden name=MilestoneName" & sScheduleDataID & " id=MilestoneName" & sScheduleDataID & " value='" & rs("item_description") & "'>"
		Response.Write "<INPUT type=hidden name=PorDateStart" & sScheduleDataID & " id=PorDateStart" & sScheduleDataID & " value='" & rs("por_start_dt") & "'>"
		Response.Write "<INPUT type=hidden name=OldDateStart" & sScheduleDataID & " id=OldDateStart" & sScheduleDataID & " value='" & sDateStart & "'>"
		Response.Write "<INPUT type=hidden name=ItemNotes" & sScheduleDataID & " id=ItemNotes" & sScheduleDataID & " value='" & sItemNotes & "'>"
		Response.Write "<INPUT type='text' size=10 id=DateStart" & sScheduleDataID & "  name=DateStart" & sScheduleDataID & " value='" & sDateStart & "' class='dateselection'><!--//&nbsp;<a href=""javascript: cmdDate_onclick('DateStart" & sScheduleDataID & "')""><img ID='picTarget' SRC='../mobilese/today/images/calendar.gif' alt='Choose Date' border='0' WIDTH='26' HEIGHT='21'></a>//--></TD>"

		If NOT bIsMilestone Then
			Response.Write "<TD valign=top class='cell' bgcolor=cornsilk align=center>"
			Response.Write "<INPUT type=hidden name=PorDateEnd" & sScheduleDataID & " id=PorDateEnd" & sScheduleDataID & " value='" & rs("por_end_dt") & "'>"
			Response.Write "<INPUT type=hidden name=OldDateEnd" & sScheduleDataID & " id=OldDateEnd" & sScheduleDataID & " value='" & sDateEnd & "'>"
			Response.Write "<INPUT type='text' size=10 id=DateEnd" & sScheduleDataID & " name=DateEnd" & sScheduleDataID & " value='" & sDateEnd & "' class='dateselection'><!--//&nbsp;<a href=""javascript: cmdDate_onclick('DateEnd" & sScheduleDataID & "')""><img ID='picTarget' SRC='../mobilese/today/images/calendar.gif' alt='Choose Date' border='0' WIDTH='26' HEIGHT='21'></a>//--></TD>"
		End If
		Response.Write "<TD><INPUT type='text' size=20 id=ChangeNote" & sScheduleDataID & " name=ChangeNote" & sScheduleDataID & " value='" & sChangeNotes &"' /></TD>"
		Response.Write "</TR>" & vbcrlf
		
		rs.MoveNext
	Loop
	Response.Write "</TABLE>"
	Response.Write "<TABLE Width='100%'><TR><TD vAlign=bottom align=right colspan=3><INPUT class='buttonHover2' type='button' value='Cancel' id=button1 name=button1 onclick='cancel_onClick();'><INPUT class='buttonHover1' type=submit value=Save id=submit name=submit></TD></TR></Table>"
	Response.Write "<INPUT type=hidden name=mode id=mode value=" & Request("Mode") & ">"
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing

End Sub

Sub Save()

	Dim dw
	Dim cn
	Dim cmd
	Dim iRowsChanged
	Dim bFoundErrors
	Dim item
	Dim dtProjectedStart
	Dim dtProjectedEnd
	Dim dtActualStart
	Dim dtActualEnd
	Dim bNotifyOfPddLock
	Dim sChangeNote
	Dim sItemNotes
    Dim obj

	If Not (Request.Form("Mode") = "Projected" Or Request.Form("Mode") = "Actual") Then
		Response.Write "<H3>Invalid Mode Selected</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If
	
	
	Set dw = New DataWrapper
	set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans
	bFoundErrors = False
	
	For Each item In Request.Form("ScheduleDataID")

		dtProjectedStart = Null
		dtProjectedEnd = Null
		dtActualStart = Null
		dtActualEnd = Null
		sChangeNote = ""
		bNotifyOfPddLock = False
		bFoundErrors = False

		If Request.Form("DateStart" & item) <> Request.Form("OldDateStart" & item) _
			Or Request.Form("DateEnd" & item) <> Request.Form("OldDateEnd" & item) Then
			
			If Request.Form("Mode") = "Projected" Then
				dtProjectedStart = Request.Form("DateStart" & item)
				dtProjectedEnd = Request.Form("DateEnd" & item)
				If Request("Milestone" & item) Then dtProjectedEnd = dtProjectedStart
			ElseIf Request("Mode") = "Actual" Then
				dtActualStart = Request.Form("DateStart" & item)
				dtActualEnd = Request.Form("DateEnd" & item)
				If Request("Milestone" & item) Then dtActualEnd = dtActualStart
			End If
			sChangeNote = Request.Form("ChangeNote" & item)
			sItemNotes = Request.Form("ItemNotes" & item)

            Set obj = New Schedule
            CALL obj.UpdateMilestone(cn, Request.Form("PorDateStart" & item), item, Trim(sUserFullName), m_UserEmail, _
                sItemNotes, sChangeNote, Request.Form("OldDateStart" & item), Request("OldDateEnd" & item), dtProjectedStart, dtProjectedEnd, _
                dtActualStart, dtActualEnd, NULL, NULL, Request.Form("ScheduleDefinitionDataID" & item), Request.Form("selItemOwner"))

    	End If

	Next

	If bFoundErrors Then
		cn.RollbackTrans
		Response.Write "<H3>Error Saving Changes</H3><H4>Transaction Rolled Back</H4>"
		Response.End
	Else
		cn.CommitTrans
		Response.Write "<INPUT type=hidden name=Success id=Success value='true'>"
	End If

	set cmd = nothing
	Set cn = Nothing
	Set dw = Nothing

End Sub
%>
<DOCTYPE html>
<HTML>
<HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<TITLE>Schedule Batch Update</TITLE>
<!-- #include file="../includes/bundleConfig.inc" -->
<style>
    .buttonHover1:hover{
	    background:#A6F4FF;
	    border:1px solid #26A0DA;
            height:25px;
            width:53px;
    }
    .buttonHover2:hover{
	    background:#A6F4FF;
	    border:1px solid #26A0DA;
        height:25px;
        width:64px;
    }
</style>
</HEAD>
<SCRIPT LANGUAGE=javascript type="text/javascript">
<!--
function cmdDate_onclick(FieldID) {
	/*var strID;
	var oldValue = window.fMain.elements(FieldID).value;
	strID = window.showModalDialog("../mobilese/today/caldraw1.asp",FieldID,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/ 
}

function cancel_onClick()
{
    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    }
    else {
        if (window.location != window.parent.location) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.close();
        }
    }
}

function onLoad()
{
	if(window.fMain.Success)
	{
	    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
	        // For Closing current popup
	        parent.window.parent.closeExternalPopup();
	        // For Reload PulsarPlusPmView Tab
	        parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);
	    }
	    else {
	        if (window.location != window.parent.location) {
	            parent.window.parent.modalDialog.cancel(true);
	        } else {
	            window.returnValue = 1;
	            window.close();
	        }
	    }
	}

    //Add modal dialog code to body tag: ---
	modalDialog.load();

    //Add datepicker to date fields
	load_datePicker();
}

function verify()
{
    if (document.fMain.mode.value == "Projected"){
        for(i=0; i<document.fMain.elements.length; i++)
        {
            var sOldName = document.fMain.elements[i].name;
            var sOldValue = document.fMain.elements[i].value;
            if (sOldName.substring(0,3) == "Old"){
                var sNewName = sOldName.substring(3);
                var sNewValue = document.fMain.elements[sNewName].value;
                var sPorName = "Por" + sNewName;
                var sPorValue = document.fMain.elements[sPorName].value;
                if (sNewName.substring(0, 7) == "DateEnd")
                    var sChangeName = "ChangeNote" + sNewName.substring(7);
                else
                    var sChangeName = "ChangeNote" + sNewName.substring(9);
                            
                var sChangeValue = document.fMain.elements[sChangeName].value;
                    
                if ((sPorValue != "") && (sOldValue != sNewValue) && (sChangeValue == "")) {
                    alert("Change Notes are required for all changes.");
                    document.fMain.elements[sChangeName].focus();
                    return false;
                }
            }
        }
    }
}
//-->
</SCRIPT>

<BODY onload="onLoad();">
<FORM action="milestonebatchupdate.asp" method="post" id="fMain" name="fMain" onsubmit="return verify();">
<input type=hidden name="PVID" id="PVID" value=<%= Request("PVID")%>>

<table border=0 width=100% cellpadding=10><tr><td>
<H3>Schedule Batch Update</H3>
<% Call Main() %>
</td></tr></table>
</FORM>
    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
</BODY>

</HTML>

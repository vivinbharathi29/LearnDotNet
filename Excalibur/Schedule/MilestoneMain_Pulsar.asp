<%@ Language=VBScript %>
<%Option Explicit%>

<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Update Product Milestone</title>
<!-- #include file="../includes/bundleConfig.inc" -->
<SCRIPT id=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	AddCat.txtName.focus();
}

function cmdDate_ScheduleResult(FieldID, oldValue) {
    var strID = $('#' + FieldID + '').val();
    var CheckBackDate = true;

    if(FieldID.indexOf("Actual") != -1){
        CheckBackDate = false;
    }

    if (typeof (strID) == "undefined") {
        return;
    } else {
        window.frmMilestone.elements(FieldID).value = strID;
    }
}

function bodyOnLoad()
{
	with (window.frmMilestone)
	{
		if (hidMilestone.value.toUpperCase() == "TRUE")
		{
			porend.style.visibility = "hidden";	
			targetend.style.visibility = "hidden";
			actualend.style.visibility = "hidden";
		}
		
        if (hidProjectedStartDt.value == "" || $("#txtActualStartDt").val() != "")
		{
		    txtActualStartDt.disabled = false;
		    txtActualEndDt.disabled = false;            
		    //hrefActualEndDtPicker.style.display = "none";
		    //hrefActualStartDtPicker.style.display = "none";
        }
	}

    //Validate form
	var sAction = $("#hidAction").val();
	if (sAction === 'View') {
	    disableFormElements("frmMilestone");
	    $("#PageTitle").text("View Schedule");
	}

    //Add datepicker
	load_datePicker();
}
//-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY bgcolor=ivory onload="bodyOnLoad()">
<%

Sub FillPhaseList(CurrentPhase)
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListSchedulePhases")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("schedule_phase_id") & """"
		If Trim(rs("schedule_phase_id")) = Trim(CurrentPhase) Then Response.Write " SELECTED "
		Response.Write ">" & rs("phase_name") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub

Sub FillOwnerList(CurrentOwner)
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "select item_description, schedule_definition_data_id from schedule_definition_data WHERE GenericOwner = 1")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
    If CurrentOwner = "" Then
        Response.Write "<OPTION value="""">--- Select Owner ---</OPTION>"
    End If

	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("schedule_definition_data_id") & """"
		If Trim(rs("schedule_definition_data_id")) = Trim(CurrentOwner) Then Response.Write " SELECTED "
		Response.Write ">" & rs("item_description") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub

'return: 0: non-custom schedule item, 
'        1: custom schedule item, 
'       -1: custom schedule item with blank owner (owner used to be a non-required field, please see the issue 3099)
Function GetGenericOwner(scheduleDefinitionDataID)
    If IsNull(scheduleDefinitionDataID) Then
        GetGenericOwner = -1    
    Else
	Dim dw, cn, cmd, rs, intGenericOwner
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "usp_SelectGenericOwner " & clng(scheduleDefinitionDataID))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
    
	intGenericOwner =  rs("GenericOwner")
    
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	GetGenericOwner = intGenericOwner
    End If
End Function


	dim cn
	dim rs 
	Dim strSQL
	dim strPORStart
	dim strPlannedStart
	dim strActualStart
	dim strPOREnd
	dim strPlannedEnd
	dim strActualEnd
	dim strComments
	dim strHistory
	dim blnValidID
	dim strMilestone
	dim strPhase
	dim strOwner
	Dim strDefinition
	Dim bIsMilestone
	Dim bShowOnReports
	Dim strScheduleDefinitionDataID
	Dim sProductVersionID
	Dim sProductVersion
	Dim sProgram
	Dim sScheduleID
	Dim strHistoryEntry
	Dim intGenericOwner
	
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	strMilestone = ""
	strComments = ""
	strPORStart = ""
	strPOREnd = ""
	strPlannedStart = ""
	strPlannedEnd = ""
	strActualStart = ""
	strActualEnd = ""
	strPhase = ""
	strOwner = ""
	strScheduleDefinitionDataID = ""
	
	
	strSQL = "usp_SelectScheduleData " & clng(request("ID")) 
    
    if request("ScheduleID") <> "" then
        strSQL = strSQL & "," & clng(request("ScheduleID"))
    end if

	rs.Open strSQL,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		blnValidID = false
	else
		blnValidID = true
		strMilestone = rs("item_description") & ""
		strComments = rs("item_notes") & ""
		strPORStart = rs("por_start_dt") & ""
		strPOREnd = rs("por_end_dt") & ""
		strPlannedStart = rs("projected_start_dt") & ""
		strPlannedEnd = rs("projected_end_dt") & ""
		strActualStart = rs("actual_start_dt") & ""
		strActualEnd = rs("actual_end_dt") & ""
		strPhase = rs("schedule_phase_id") & ""
		strOwner = rs("schedule_definition_data_id") & ""
		strDefinition = server.HTMLEncode(rs("item_definition")&"")
		strDefinition = Replace(strDefinition, vbCrLf, "<BR>")
		strScheduleDefinitionDataID = rs("schedule_definition_data_id")
		sProductVersionID = rs("productversionid")
		sProgram = rs("family_name") & " " & rs("schedule_name")
		sScheduleID = rs("schedule_id") & ""
		intGenericOwner = GetGenericOwner(strScheduleDefinitionDataID)
		
		If UCase(rs("show_on_reports_yn")) = "Y" Then
			bShowOnReports = True
		Else
			bShowOnReports = False
		End If
		
		If UCase(rs("milestone_yn")) = "Y" Then
			bIsMilestone = True
		Else
			bIsMilestone = False
		End If
	end if
	rs.Close
	
    strHistory = ""

	strSQL = "usp_SelectScheduleItemHistory " & clng(request("ID"))
	rs.Open strSQL,cn,adOpenStatic
	If Not (rs.EOF And rs.BOF) Then
	    Do Until rs.EOF
	        strHistoryEntry = ""
	        
	        If rs("old_projected_start_dt") & "" <> rs("new_projected_start_dt") & "" Then
	            If rs("old_projected_start_dt") & "" = "" Then
	                strHistoryEntry = strHistoryEntry & "Projected Start Date set to " & rs("new_projected_start_dt") & "." & vbCrLf
	            Else
	                strHistoryEntry = strHistoryEntry & "Projected Start Date Changed from " & rs("old_projected_start_dt") & " to " & rs("new_projected_start_dt") & "." & vbCrLf
                End If
	        End If
	        If rs("old_projected_end_dt") & "" <> rs("new_projected_end_dt") & "" And UCase(rs("milestone_yn")) = "N" Then
	            If rs("old_projected_end_dt") & "" = "" Then
	                strHistoryEntry = strHistoryEntry & "Projected End Date set to " & rs("new_projected_end_dt") & "." & vbCrLf
	            Else
	                strHistoryEntry = strHistoryEntry & "Projected End Date Changed from " & rs("old_projected_end_dt") & " to " & rs("new_projected_end_dt") & "." & vbCrLf
	            End If
	        End If
	        If rs("old_actual_start_dt") & "" <> rs("new_actual_start_dt") & "" Then
	            If rs("old_actual_start_dt") & "" = "" Then
	                strHistoryEntry = strHistoryEntry & "Actual Start Date set to " & rs("new_actual_start_dt") & "." & vbCrLf
	            Else
	                strHistoryEntry = strHistoryEntry & "Actual Start Date Changed from " & rs("old_actual_start_dt") & " to " & rs("new_actual_start_dt") & "." & vbCrLf
	            End If
	        End If
	        If rs("old_actual_end_dt") & "" <> rs("new_actual_end_dt") & "" And UCase(rs("milestone_yn")) = "N" Then
	            If rs("old_actual_end_dt") & "" = "" Then
	                strHistoryEntry = strHistoryEntry & "Actual End Date set to " & rs("new_actual_end_dt") & "." & vbCrLf
	            Else
	                strHistoryEntry = strHistoryEntry & "Actual End Date Changed from " & rs("old_actual_end_dt") & " to " & rs("new_actual_end_dt") & "." & vbCrLf
	            End If
	        End If

	        If strHistoryEntry <> "" Then
	            strHistoryEntry = rs("last_upd_date") & " - " & rs("last_upd_user") & vbCrLf & strHistoryEntry
	            if rs("notes") & "" <> "" Then  strHistoryEntry = strHistoryEntry & "Change Notes: " & rs("notes") & "" & vbcrlf
	            strHistoryEntry = strHistoryEntry & "----------------------------------------" & vbcrlf 
                strHistory = strHistory & strHistoryEntry
            End If

	        rs.MoveNext
	    Loop
	End If

	if not blnValidID then
		Response.Write "<BR><font size=2>Unable to find the selected Milestone</font>"
	else
		Response.Write "<h3 id=""PageTitle"">Update Milestone</h3>"
	%>
	<form id="frmMilestone" method="post" action="MilestoneSave.asp">
	<input type="hidden" id="hidMilestone" name="hidMilestone" value="<%=bIsMilestone%>" />
	<input type="hidden" id="hidScheduleDefinitionDataID" name="hidScheduleDefinitionDataID" value="<%= strScheduleDefinitionDataID%>" />
	<input type="hidden" id="hidScheduleDataID" name="hidScheduleDataID" value="<%=Request("ID")%>" />
	<input type="hidden" id="PVID" name="PVID" value="<%=sProductVersionID%>" />
	<input type="hidden" id="hidPorStartDt" name="hidPorStartDt" value="<%=strPorStart%>" />
	<input type="hidden" id="hidPorEndDt" name="hidPorEndDt" value="<%=strPorEnd%>" />
	<input type="hidden" id="hidProgram" name="hidProgram" value="<%=sProgram%>" />
	<input type="hidden" id="hidMilestoneName" name="hidMilestoneName" value="<%=strMilestone%>" />
	<input type="hidden" id="hidScheduleID" name="hidScheduleID" value="<%=sScheduleID%>" />
    <input type="hidden" id="hidAction" value="<%= request.querystring("action")%>">
    <input type="hidden" id="txtID" value="<%=sProductVersionID%>" />
	<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%= request.querystring("pulsarplusDivId")%>">
	<table id="tabUpdate" width="100%" bgcolor="cornsilk" border="1" cellspacing="0" cellpadding="2" bordercolor="tan">
		<tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Phase:</td>
			<td colspan=2>
				<select id="selItemPhase" name="selItemPhase">
					<% Call FillPhaseList(strPhase) %>
				</select>
			</td></tr>
		<tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold; width:150px;">Schedule Item:</td>
			<td style="width:100%; font-size:x-small; font-family:verdana;" colspan="2"><%=strMilestone%></td></tr>
        <%if intGenericOwner <> 0 then %>		
		<tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold; width:150px;">Owner:&nbsp;<font color="red" size="1">*</font></td>
			<td colspan=2>
				<select id="selItemOwner" name="selItemOwner">
					<% Call FillOwnerList(strOwner) %>
				</select>
			</td></tr>
        <% end if %>
		<tr>
			<td>&nbsp;</td>
			<td><strong>Start Date</strong></td>
			<td><strong>End Date</strong></td></tr>
	    <tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">POR:</td>
			<td><%=strPORStart%>&nbsp;</td>
			<td><span id="porend"><%=strPOREnd%>&nbsp;</span></td></tr>
	    <tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Current Date:</td>
			<td>
				<input type="hidden" id="hidProjectedStartDt" name="hidProjectedStartDt" value="<%=strPlannedStart%>" />
				<input type="text" id="txtProjectedStartDt" name="txtProjectedStartDt" value="<%=strPlannedStart%>" class="dateselection-validate" autocomplete="off" />
			</td>
			<td>
				<span id="targetend">
				<input type="hidden" id="hidProjectedEndDt" name="hidProjectedEndDt" value="<%=strPlannedEnd%>" />
				<input type="text" id="txtProjectedEndDt" name="txtProjectedEndDt" value="<%=strPlannedEnd%>" class="dateselection-validate" autocomplete="off" />
                </span></td></tr>
	    <tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Actual Date:</td>
			<td>
				<input type="hidden" id="hidActualStartDt" name="hidActualStartDt" value="<%=strActualStart%>" />
				<input type="text" id="txtActualStartDt" name="txtActualStartDt" value="<%=strActualStart%>" class="dateselection-validate" autocomplete="off" />
			</td>
			<td>
				<span id="actualend">
				<input type="hidden" id="hidActualEndDt" name="hidActualEndDt" value="<%=strActualEnd%>" /> 
				<input type="text" id="txtActualEndDt" name="txtActualEndDt" value="<%=strActualEnd%>" class="dateselection-validate" autocomplete="off" /> 
				</span></td></tr>
	    <tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Comments:</td>
			<td colspan="2"><textarea style="WIDTH: 100%; HEIGHT: 60px" rows="3" cols="51" id="txtComments" name="txtComments"><%=strComments%></textarea></td></tr>
	    <tr>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Change Notes:<!--<span style="font-size:xx-small; color:Red;">*</span>--></td>
			<td colspan="2"><textarea style="WIDTH: 100%; HEIGHT: 60px" rows="3" cols="51" id="txtItemNotes" name="txtItemNotes"></textarea></td></tr>
		<tr>
		    <td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Change History:</td>
		    <td colspan="2"><textarea style="width:100%;height:120px; background-color:beige;" readonly="readonly" rows="12" cols="51" id="txtItemHistory"><%=strHistory%></textarea></td></tr>
	    <tr <%If Trim(strDefinition) = "" Then Response.Write "style=""display:none"""%>>
			<td style="white-space:nowrap; vertical-align:top; font-weight:bold;">Definition:</td>
			<td colspan="2"><%=strDefinition%>&nbsp;</td></tr>
	    <tr>
			<td nowrap valign="top"><b>Show On Printed Reports:</b></td>
			<td colspan=2><input type="checkbox" id=cbxShowOnReports name=cbxShowOnReports <%If bShowOnReports Then Response.Write "CHECKED"%>></td></tr>
	</table>
	</form>

<%
	end if
set rs = nothing
set cn = nothing

%>
</BODY>
</HTML>

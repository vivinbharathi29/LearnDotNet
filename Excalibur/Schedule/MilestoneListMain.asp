<%@ Language=VBScript %>
<%
	
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
	  
%>

<!-- #include file = "../includes/noaccess.inc" -->

<DOCTYPE html>
<HTML>
<HEAD>
<title>Milestone List</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../includes/bundleConfig.inc" -->
<SCRIPT type="text/javascript" LANGUAGE=javascript>
<!--

function chkMilestone_onclick( RowID ) {

	if ( window.frmSchedule.chkTag[RowID].checked )
		window.frmSchedule.chkTag[RowID].checked = false;
	else
		window.frmSchedule.chkTag[RowID].checked = true;
		
	return true;
}

function chkPhase_onclick( obj )
{
	var i;
	var max = window.frmSchedule["chkSelected" + obj].length
	var rowID;
	for (i = 0; i < max; i++)
	{
		if (window.frmSchedule["chkSelected" + obj][i].checked != window.frmSchedule["phase" + obj].checked)
		{
			window.frmSchedule["chkSelected" + obj][i].checked = window.frmSchedule["phase" + obj].checked;
			rowID = window.frmSchedule["chkSelected" + obj][i].parentElement.parentElement.id;
			chkMilestone_onclick(rowID);
		}
	}
}

function addItem_onClick( PVID, ScheduleID )
{
    modalDialog.open({ dialogTitle: 'Add Schedule Item', dialogURL: 'Schedule.asp?ProdVID=' + PVID + '&ScheduleID=' + ScheduleID + '', dialogHeight: 550, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
	/*var strID;
	strID = window.showModalDialog("Schedule.asp?ProdVID=" + PVID + "&ScheduleID=" + ScheduleID, "", "dialogWidth:545px;dialogHeight:475px;maximize:Yes;edge: Sunken;center:Yes; help: No;resizable: Yes;status: Yes")
	if (typeof(strID) != "undefined")
		window.location.href = window.location.href;*/
}
//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs     
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();
}

//*****************************************************************
//Description:  Close Modal Dialog Window
//Function:     CancelModalDialog();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367      
//*****************************************************************
function CancelModalDialog() {
    //close child dialog window
    modalDialog.cancel();
}
//-->
</SCRIPT>
</HEAD>
<BODY onload="window_onload();" bgcolor="ivory">

<form ID=frmSchedule action=MilestoneListSave.asp method=post>
<%	

	dim cn 
	dim rs
	dim i
	dim strLastPhase


if request("ScheduleID") = "" then
	Response.Write "<BR><font size=2 face=verdana>Not enough information to display this page</font>"
else
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.open "spGetProductVersionName " & clng(request("PVID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product</font>"
		rs.Close
	else
		Response.Write "<table width='100%'><tr><td><font size=4 face=verdana><b>Select " & rs("name") & " Schedule Items</b></font></td><td align='right'><a href='javascript:addItem_onClick(" & Request("PVID") & ", " & Request("ScheduleID") & ");'><font face='verdana' size=1>Add Custom Item</font></a></td></tr></table><BR><BR>"
		rs.Close
		rs.Open "usp_SelectScheduleData NULL," & clng(Request("ScheduleID")) & ",NULL,NULL,NULL",cn,adOpenForwardOnly
		Response.Write "<Table width=100% border=0 cellpadding=1 cellspacing=0>" & _
			"<TR bgcolor=cornsilk><TD>&nbsp;</TD><TD><font size=1 face=verdana><b>Item Description</b></font></TD>" & _
			"<TD><font size=1 face=verdana><b>Type&nbsp;&nbsp;</b></font></TD>" & _
			"<TD><font size=1 face=verdana><b>Required&nbsp;&nbsp;</b></font></TD></tr>"
		i=0
		strLastPhase = ""
		do while not rs.EOF 
			if strLastPhase <> rs("item_phase") & "" then
				Response.Write "<TR bgcolor=MediumAquamarine><TD style=""BORDER-TOP: gray thin solid""><INPUT type=checkbox style=""width:16;height:16;"" id=phase" & rs("schedule_phase_id") & " onclick=""return chkPhase_onclick(" & rs("schedule_phase_id") & ")""></TD><TD colspan=3 style=""BORDER-TOP: gray thin solid""><font size=2 face=verdana><b>" & rs("phase_name") & "</b></font></TD></TR>"
				strLastPhase = rs("item_phase") & ""
			end if
			if Ucase(rs("active_yn")) = "Y" then
				Response.Write "<TR bgcolor=lightsteelblue ID=" & i & "><TD style=""BORDER-TOP: gray thin solid"">"
				If UCase(rs("required_yn")) = "Y" Then
					Response.Write "<INPUT type=checkbox style=""width:16;height:16;"" disabled checked>"
				Else
					Response.Write "<INPUT value=""" & rs("schedule_data_id") & """ checked  style=""width:16;height:16;"" class=" & i & " type=""checkbox"" id=chkSelected" & rs("schedule_phase_id") & " name=chkSelected onclick=""return chkMilestone_onclick(" & i & ")"">"
				End If
				Response.Write "<INPUT value=""" & rs("schedule_data_id") & """ style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
			else
				Response.Write "<TR bgcolor=Ivory ID=" & i & "><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("schedule_data_id") & """ style=""width:16;height:16;"" class=" & i & " type=""checkbox"" id=chkSelected" & rs("schedule_phase_id") & " name=chkSelected onclick=""return chkMilestone_onclick(" & i & ")""><INPUT value=""" & rs("schedule_data_id") & """ style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
			end if
			response.write "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("item_description") & "</font></TD>"
			response.write "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" 
				If UCase(rs("milestone_yn")) = "Y" Then
					Response.Write "Milestone"
				Else
					Response.Write "Task"
				End If
			Response.Write"&nbsp;</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>"
				If UCase(rs("required_yn")) = "Y" Then
					Response.Write "Yes"
				Else
					Response.Write "-"
				End If
			Response.Write "&nbsp;</font></TD></TR>"
			i=i+1
			rs.MoveNext
			
		loop
		rs.Close
		Response.Write "</table>"
	  
	end if
	set rs = nothing
	cn.Close
	set cn = nothing
end if


  %>
    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
<INPUT type="hidden" id=PVID name=PVID value="<%= request("PVID")%>">
</form>
</BODY>
</HTML>

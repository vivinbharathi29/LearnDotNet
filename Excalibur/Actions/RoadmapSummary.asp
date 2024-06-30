<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}
</STYLE>
<BODY>
<font size=2 face=verdana>SI Tools Roadmap Summary</font><BR><BR>
<TABLE style="BORDER-LEFT-STYLE: none;BORDER-Right-STYLE: none" border=1 width=100% cellspacing=0 cellpadding=2 bordercolor=Gray>
<%
	dim cm
	dim rs
	dim strLastOwner
	dim OwnerCount
	dim OwnerTotal
	dim ColorArray
	dim strTargetNotes
	ColorArray = split ("lavender,ivory,powderblue,cornsilk",",")
'	ColorArray = split ("lavender, green, blue, ivory, cornsilk, powderblue",",")
	
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")

	strLastOwner = ""
	OwnerCount = 0
	OwnerTotal = 0
	rs.open "spListRoadmapSummaryByGroup 439",cn,adOpenForwardOnly
	do while not rs.EOF
		if strLastOwner <> trim(rs("Owner")) then
			OwnerCount = 0
			OwnerTotal = OwnerTotal + 1
			strLastOwner = trim(rs("Owner") & "")
		end if
		OwnerCount=OwnerCount+1
		if trim(rs("TimeFrame") & "") = "" then
			strTimeframe = "TBD"
		elseif isdate(rs("Timeframe")) then
			if year(cdate(rs("Timeframe"))) > 1600 then
				strTimeframe = formatdatetime(rs("Timeframe") & "",vbshortdate)
			else
				strTimeframe = rs("Timeframe") & ""
			end if
		else
			strTimeframe = rs("Timeframe") & ""
		end if

		if trim(rs("OriginalTimeFrame") & "") = "" then
			strOriginalTimeframe = strTimeframe
		elseif isdate(rs("OriginalTimeframe"))  then
			if year(cdate(rs("OriginalTimeframe"))) > 1600 then
				strOriginalTimeframe = formatdatetime(rs("OriginalTimeframe") & "",vbshortdate)
			else
				strOriginalTimeframe = rs("OriginalTimeframe") & ""
			end if
		else
			strOriginalTimeframe = rs("OriginalTimeframe") & ""
		end if

		if trim(rs("timeframenotes") & "") = "" then
			strTargetNotes = "N/A"
		else
			strTargetNotes = rs("timeframenotes") & ""
		end if
		
		rs2.open "spListRoadmapTasks " & rs("ID") & ",0",cn,adOpenForwardOnly
		strTasks = ""
		TaskCount = 0
		TaskCompletedCount = 0
		do while not rs2.EOF
			TaskCount = TaskCount + 1
			if rs2("Status") = "Closed" then
				TaskCompletedCount = TaskCompletedCount + 1
			else
				strTasks = strTasks & "<TR><TD valign=top width=10><font face=wingdings>" & chr(159) & "</font></TD><TD>" & rs2("Summary") & "</TD></TR>"
			end if
			rs2.MoveNext
		loop
		rs2.close	
		strStatus = ""
		if TaskCount = 0 then
			strStatus = "Under Investigation"
		elseif TaskCount = TaskCompletedCount then
			strStatus = "<b>Complete</b>"
		else
			strStatus = "Completed " & TaskCompletedCount & " of " & TaskCount
			if TaskCount = 1 then
				strStatus = strStatus & " task."
			else
				strStatus = strStatus & " tasks."
			end if   
			
			if TaskCount <> 0 then
				strStatus = strStatus & " (" & round((TaskCompletedCount/TaskCount) * 100) & "%)"
			end if
		end if
		
		if strTasks <> "" then
			strTasks = "<Table border=0>" & strTasks & "</Table>"
		else
			strTasks = "&nbsp;"
		end if	


		rs2.open "spListRoadmapIssues " & rs("ID") & ",1",cn,adOpenForwardOnly
		strIssues = ""
		do while not rs2.EOF
			strIssues = strIssues & "<TR><TD valign=top width=10><font face=wingdings>" & chr(159) & "</font></TD><TD>" & rs2("Summary") & "</TD></TR>"
			rs2.MoveNext
		loop
		rs2.close	
		
		if strIssues <> "" then
			strIssues = "<Table border=0>" & strIssues & "</Table>"
		end if	

		if OwnerCount < 4 then
			if OwnerTotal -1 > ubound(colorarray) then
				OwnerTotal = 1
			end if
			Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD nowrap><b>Summary: </b></TD><TD colspan=3 nowrap>" & rs("Summary") & "</td></TR>"
			Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD nowrap><b>Owner: </b></TD><TD nowrap>" & rs("Owner") & "</td><td nowrap><b>Current Target: </td><td>" & strTimeframe & "</TD></tr>"
			Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD nowrap><b>Project: </b></TD><TD nowrap>" & rs("DOTSName") & "</td><td nowrap><b>Original Target: </td><td>" & strOriginalTimeframe & "</TD></tr>"
			Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD valign=top nowrap><b>Status: </b></TD><TD nowrap>" & strStatus & "&nbsp;</td><td valign=top nowrap><b>Reason Target Changed: </td><td>" & strTargetNotes & "&nbsp;</TD></tr>"
			if strIssues <> "" then
				Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD width=50% colspan=2 valign=top><b>Remaining Tasks: </b><BR>" & strTasks & "</td><td width=50% colspan=2 valign=top nowrap><b>Issues/Risks: <BR>" & strIssues & "&nbsp;</TD></tr>"
			else
				Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD colspan=4 valign=top><b>Remaining Tasks: </b><BR>" & strTasks & "</td></tr>"
			end if
			Response.Write "<TR bgcolor=" & colorarray(OwnerTotal - 1) & "><TD colspan=4 valign=top><b>Description:</b><BR>" & replace(rs("Details") & "",vbcrlf,"<BR>") & "&nbsp;</td></tr>"
			Response.Write "<TR bgcolor=white><TD style=""BORDER-TOP-STYLE: none;BORDER-BOTTOM-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-Right-STYLE: none"" colspan=4>&nbsp;</td></tr>"
		end if
		rs.MoveNext
	loop
	rs.Close
	
	
	set rs = nothing
	set rs2 = nothing
	cn.Close
	set cn=nothing
%>
</TABLE>
</BODY>
</HTML>

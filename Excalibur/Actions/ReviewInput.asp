<%@ Language=VBScript %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
td
{
    font-family: Verdana;
    font-size:x-small;
}

    a:visited
    {
        color: blue
    }
    a:hover
    {
        color: red
    }
    a
    {
        color: blue
    }
</style>
</HEAD>
<BODY>
<font face=verdana size=2>
<center>
<%if request("Report") = "1" then %>
    Important 
<%end if%>
Tasks I Closed this Review Cycle<br>
<font size=1 face=verdana>10/16/2012-10/15/2013</font><br /><br />
</center>

<b>Display: </b>
<%if request("Report") = "1" then%>
Important Tasks , <a href="ReviewInput.asp">All Tasks</a> , <a href="ReviewInput.asp?Report=2">Support Tickets</a><br /><br />
<%elseif request("Report") = "2" then%>
<a href="ReviewInput.asp?Report=1">Important Tasks</a> , <a href="ReviewInput.asp">All Tasks</a> , Support Tickets<br /><br />
<%else%>
<a href="ReviewInput.asp?Report=1">Important Tasks</a> , All Tasks , <a href="ReviewInput.asp?Report=2">Support Tickets</a><br /><br />
<%end if%>
<%
	dim cn
	dim rs
	dim strFirstDate
	dim sectionCount
	dim TotalCount
	dim SectionTasks
	dim TotalTasks
    dim strResolution
    dim ResolutionArray
    dim ImportantColor

    if request("Report") = "1" then
        ImportantColor = "black"
    else
        ImportantColor = "maroon"
    end if

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("Email") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	'if currentuserid = 31 then
	'	currentuserid = 694
	'end if

if request("Report") <> "2" then

	if CurrentUserID = 0 then
		Response.Write "Unable to find user in Excalibur."
	else	
		    strSQL = "Select i.ActualDate, r.Summary as Milestone, v.dotsname as product, Type, i.summary, i.onstatusreport, i.description, i.resolution " & _
				     "from deliverableissues i with (NOLOCK), productversion v with (NOLOCK), productfamily f with (NOLOCK), ActionRoadmap r with (NOLOCK) " & _
				     "where i.productversionid = v.id " & _
				     "and f.id = v.productfamilyid " & _
				     "and r.id = i.actionroadmapid " & _
				     "and i.ownerid = " & clng(CurrentUSerID) & " " & _
				     "and v.typeid= 2 " & _
				     "and i.type=2 " & _
                     "and i.summary not like 'Cloned%' " & _
				     "and (ActualDate between '10/16/2012' and '10/15/2013') " & _
				     "and i.status in (2,4,5) " 
	    if request("Report") = "1" then
            strSQl = strsql & "and i.onstatusreport = 1 "
            strSQl = strsql & "Order By r.summary" 'r.summary
        else
            strSQl = strsql & "Order By r.summary, i.onstatusreport desc" 
        end if

        rs.Open strSQL,cn,adOpenForwardOnly
		strLastMilestone = ""
		strFirstDate = Now
		SectionCount = 0
		TotalCount=0
		SectionTasks=0
	    TotalTasks=0

		do while not rs.EOF

		  if lcase(rs("Resolution") & "") <> "dup" and  left(lcase(rs("Resolution") & ""),5) <> "dup -" and lcase(rs("Resolution") & "") <> "obsolete" and lcase(rs("Resolution") & "") <> "nfn" and lcase(rs("Resolution") & "") <> "previously implemented" and lcase(rs("Resolution") & "") <> "duplicate"  and lcase(rs("Resolution") & "") <> "cancelled" and lcase(rs("Resolution") & "") <> "no needed anymore" and lcase(rs("Resolution") & "") <> "no longer needed"  and lcase(rs("Resolution") & "") <> "cancel"  and lcase(rs("Resolution") & "") <> "canceled" and lcase(rs("Resolution") & "") <> "dup." and left(lcase(rs("Resolution") & ""),6) <> "dup of"  and left(lcase(rs("Resolution") & ""),11) <> "duplicate -"  and left(lcase(rs("Summary") & ""),11) <> "duplicate -" then
			if strLastMilestone <> rs("Milestone") then
				if sectionCount <> 0 then
					TotalTasks = TotalTasks + SectionTasks
					Response.write "</ul></td></tr></table><font size=2 face=verdana color=green>Action Items Closed In This Section: " & SectionCOunt & "</font><BR><BR><BR><BR>"
					'Response.write "<font size=2 face=verdana color=red>Sub-Tasks Closed In This Section: " & SectionTasks & "</font><BR><BR><BR><BR>"
				end if
				Response.Write "<table width=""100%"" border=0 cellspacing=0 cellpadding=4><tr bgcolor=gainsboro><td><b>" & rs("Milestone") & "</b><BR></td></tr><tr><td>"
				SectionCount = 0
				SectionTasks = 0
			end if
			strLastMilestone = rs("Milestone") 
			SectionCOunt = SectionCount + 1
			TotalCount=TotalCount + 1
			if datediff("d",strFirstDate,rs("ActualDate")) < 0 then
				strFirstDate = rs("ActualDate")
			end if
            response.write "<ul>"
			if rs("OnstatusReport") then
				Response.Write "<li><font color=" & ImportantColor & ">" & rs("Summary") & "</font></li>"		
			else
				Response.Write "<li>" & rs("Summary") & "</li>"		
			end if

			if rs("Resolution") & "" <> "" then
                ResolutionArray = split(rs("Resolution") & "",vbcrlf)
                strResolution = ""
                for i = 0 to ubound(ResolutionArray)
                    if trim(ResolutionArray(i)) <> "" then
                        if i = 0 then
                            strResolution = "<li>" & resolutionarray(i) & "</li>"
                        else
                            strResolution = strResolution & "<li>" & resolutionarray(i) & "</li>"
                        end if   
                    end if
                next
				'strResolution = replace(replace(rs("Resolution"),vbcrlf,"</li><li>"),vbcrlf & vbcrlf,vbcrlf)
                if rs("OnstatusReport") = 0 then
					Response.write "<ul>" & strResolution & "</ul>"
				else
					Response.write "<font color=" & ImportantColor & "><ul>" & strResolution & "</ul></font>"
				end if
			end if
            response.write "</ul>"
    		SectionTasks = SectionTasks + CountSubTasks(rs("Resolution")& "")
'			Response.Write "<BR>"
		  end if
			rs.MoveNext
		loop
		rs.Close

				if sectionCount <> 0 then
        			TotalTasks = TotalTasks + SectionTasks
					Response.write "<font size=2 face=verdana color=green>Action Items Closed In This Section: " & SectionCOunt & "</font><BR>"
					'Response.write "<font size=2 face=verdana color=red>Sub-Tasks Closed In This Section: " & SectionTasks & "</font><BR><BR>"
				end if

	
	end if


	
	function CountSubTasks(strResolution)
	    dim TaskCount
	    dim TaskArray
	    dim strTask
	    TaskCount =0
    	taskarray = split(strResolution,vbcrlf)
    	for each strtask in TaskArray
    	    if trim(strTask) <> "" then
    	        TaskCount = TaskCount + 1
    	    end if
    	next
    	if TaskCount = 0 then
    	    TaskCount = 1
    	end if
	    CountSubTasks = TaskCount
	end function
%>
<BR>
<font size=2 face=verdana color=blue><HR><BR>
Total Action Items Closed This Period: <%=TotalCount%><BR>
<!--Total Sub-Tasks Closed This Period: <%=TotalTasks%><BR>-->
<!--First Action Item Closed On: <%=strFirstDate%>-->
</font>
<%
else
    rs.open "Select * from supportissues with (NOLOCK) where statusid=2 and ownerid = " & clng(currentuserid),cn
    if not (rs.eof and rs.bof) then
        Response.Write "<BR><BR><hr><b>Tickets Closed</b><br><br>"
    end if
    do while not rs.EOF
        response.write "-" & rs("Summary") & "<br>"
        if trim(rs("Resolution") & "") <> "" then
             response.write "--" & rs("Resolution") & "<br>"
        end if
        response.write "<BR>"
        rs.MoveNext
    loop
    rs.close

	set rs = nothing
	cn.close
	set cn = nothing
end if

 %>

</font>
</BODY>
</HTML>

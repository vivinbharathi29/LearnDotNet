<html>

<link href="../../style/Excalibur.css" type="text/css" rel="stylesheet">

<body bgColor="white">
<%

	dim cn
	dim RowCount
	dim strID
	dim strDCRCreated
	strID = " " & request("ID")
%>

<center>
	<font face=verdana size=4><b>Change Request <%=strID%> History</b></font>
	<font face=verdana size=2><br><br><%=now()%><BR><BR></font>
</center>
	
<%

    Function GetFriendlyName(FieldName)
    Dim newName
    Select Case FieldName
           Case ChangeDt
                newName = "Change Date"
           Case Else
                newName = FieldName
    End Select
    GetFriendlyName = newName
    End Function

    set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
	
    dim ApprovalCount
	  set rs = server.CreateObject("ADODB.Recordset")
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovals"
		
		Set p = cm.CreateParameter("@ActionID", 3, &H0001)
		p.Value = clng(request("ID"))
		cm.Parameters.Append p
	
		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spListApprovals " & IssueID,cn,adOpenForwardOnly
		'ApprovalCount = 0
%>		
		
<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%  
    'ar = rs.GetRows()
    'Response.Write "<TR bgcolor=cornsilk><TD><center><font size=1 face=verdana>Change Request Created " & ar(8,0) & "</font></center></td></tr>"
%>	
</TABLE>
<% 
Response.Write "<BR>"
%>
<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%
    set rs = server.CreateObject("ADODB.Recordset")
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovals"
		
		Set p = cm.CreateParameter("@ActionID", 3, &H0001)
		p.Value = clng(request("ID"))
		cm.Parameters.Append p
	
		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
        i=0
		do while not rs.EOF
		    If i=0 Then
		        'Response.Write "<TR bgcolor=cornsilk><TD><center><font size=1 face=verdana>Change Request Created " & rs("Created") & "</font></center></td></tr>"
		        Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>Approver</b></font></td><TD><font size=1 face=verdana><b>Date Approved</b></font></td><TD><font size=1 face=verdana><b>Status</b></font></td><TD><font size=1 face=verdana><b>Comments</b></font></td></tr>"
		    End If
           
		    strStatusText = rs("Status")
		    select case strStatusText
			case "1"
			    strStatusText = "Approval Requested"
			case "2"
			    strStatusText = "Approved"
			case "3"
			    strStatusText = "Disapproved"
			case "4"
			    strStatusText = "Cancelled"
			case "5"
			    strStatusText = "Not Applicable"
			end select
			
			Response.Write "<TR><TD>" & rs("Approver") & "&nbsp;</TD><TD>" & rs("Updated") & "&nbsp;</TD><TD>" & strStatusText & "&nbsp;</TD><TD>" & rs("Comments") & "&nbsp;</TD></TR>" 
            i=i+1
		    rs.MoveNext
		loop			
		rs.Close
%>
</TABLE>
<%
	Response.Write "<BR>"
%>
<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%
	set rs = server.CreateObject("ADODB.Recordset")
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 1
	cm.CommandText = "SELECT ChangeDt,(select Name from ActionStatus where id = dh.Status) as Status,(select FullName from UserInfo where UserId = dh.OwnerID) as Owner,dh.OwnerID,ActualDate,(CASE isnumeric(LastUpdUser) when 1 then (select FullName from UserInfo where UserId=LastUpdUser) ELSE LastUpdUser END )as LastUpdUser FROM DeliverableIssuesHistory dh WHERE id = " & clng(request("ID")) & " ORDER BY ChangeDt DESC" 
	
	rs.CursorType = adOpenStatic
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
    
    RowCount=0	    
    
    if Not (rs.EOF and rs.BOF) Then
        Response.Write "<TR><TD><b>Changed Date:</b></TD><TD><b> Status </b></TD><TD><b> Owner </b></TD><TD><b> Actual Date </b></TD><TD><b> Updated By </b></TD></TR>"
        do while not rs.EOF
                    Response.Write "<TR><TD>" & rs("ChangeDt") & "</TD><TD>" & rs("Status") & "</TD><TD>" & rs("Owner") & " &nbsp;</TD><TD>" & rs("ActualDate") & " &nbsp; </TD><TD>" & rs("LastUpdUser") & "</TD></TR>"
                  
                    RowCount = RowCount + 1

		    rs.MoveNext
		loop			

    else
        Response.Write "<TR><TD><center><font size=2 face=verdana>There Is No History Available For This Change Request</font></center></td></tr>"
    end if

  rs.close
	cn.Close
	set rs = nothing
	set cn = nothing
%>
</TABLE>
<%
	if RowCount > 0 then
		Response.Write "<BR><BR>Changes Displayed: " & RowCount
	end if
%>

</body>
</html>
 

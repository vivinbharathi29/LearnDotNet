<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (txtID.value == "")
        frmMain.VersionID.focus();
	
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
}
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
   dim strDeliverable 
   dim blnValidState

    strState = ""

    if request("ID") = "" then
        response.write "<h5>Deliverable Version - Rollback Workflow</h5>"

        response.write "Please enter the Excalibur ID number of the Deliverable you want to rollback."
        response.write "<BR><form id=frmMain method=get action=""RollbackVersionWorkflow.asp"">Excalibur ID: <input id=""VersionID"" name=""ID"" type=""text"" value=""""></form>"
    else

        if request("UpdateOK") <> "1" then
            response.write "<font color=red><b>Preview Only.  Nothing in Excalibur was updated.</b></font><br><br>"
        end if
        response.write "<b>Deliverable Version - Rollback Workflow.</b></font><br><br>"
        response.write "<font size=1 face=verdana> - Moves a deliverable back one step in the workflow.</b></font><br>"
        response.write "<font size=1 face=verdana> - No notifications are sent.</b></font><br><br>"
        dim cn
	    dim rs 
	    dim strSQL

        strSQl = "Select ID, DeliverableName as name, Version, Revision, Pass, Status, Location " & _
                 "from deliverableversion with (NOLOCK) " & _
                 "where id= " & clng(request("ID"))
	    set cn = server.CreateObject("ADODB.Connection")
	    set rs = server.CreateObject("ADODB.Recordset")

	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open

        rs.open  strSQl, cn,adOpenStatic
        strDeliverable = ""
        if rs.eof and rs.bof then
            response.write "Unable to find the selected deliverable"
        else
            strState = trim(rs("Status") & "")
            
            strDeliverable = rs("Name") & " " & rs("Version")
            if trim(rs("Revision") & "") <> "" then
                strDeliverable = strDeliverable & "," & rs("Revision")
            end if
            if trim(rs("Pass") & "") <> "" then
                strDeliverable = strDeliverable & "," & rs("Pass")
            end if
            
            if strState <> "2" and strState <> "3" then
                response.write "This deliverable is not is a state that can be rolled back.  Please contact Dave Whorton for assistance."
            else
                if request("UpdateOK") <> "1" then
                    response.write "<a href=""RollbackVersionWorkflow.asp?UpdateOK=1&ID=" & request("ID") & """>Update Excalibur now</a><BR><BR>"
	            end if
                response.write "<b>" & strDeliverable & "<b>"
	            response.write "<table bgcolor=ivory  width=""100%"" border=1 bordercolor=gainsboro cellpadding=2 cellspacing=0>"
	            response.write "<tr bgcolor=beige><td><b>Table</td><td><b>Field</td><td><b>Current&nbsp;Value</td><td><b>New&nbsp;Value</td></tr>"
            
                rs.Close
                strSQl = "Select ID,Milestone, Actual, StatusID " & _
                         "from deliverableschedule with (NOLOCK) " & _
                         "where deliverableversionid= " & clng(request("ID")) & " " & _
                         "and statusid in (1,3) " & _
                         "order by id desc"
                rs.open strSQl, cn,adOpenStatic
                do while not rs.EOF
                    if trim(rs("StatusID") & "") = "1" then
                        response.write "<BR>" & rs("ID") & " In Progress"
                    elseif trim(rs("StatusID") & "") = "3" then
                        response.write "<BR>" & rs("ID") & " Done"
                    end if
                    rs.MoveNext
                loop        
            
            end if
            
        end if
        rs.Close
        set rs = nothing

        cn.Close
        set cn = nothing
	    response.write "</table>"

    end if

%>
<input id="txtID" name="txtID" type="hidden" value="<%=request("ID")%>">
</BODY>
</HTML>

<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
A:link
{
    COLOR: Blue;
}
A:visited
{
    COLOR: Blue;
}
A:hover
{
    COLOR: red;
}  
</STYLE>

</HEAD>


<BODY>
<%
    dim strSQL, cn, rs

    if request("RootID") = "" or request("StatusID") = "" then
        response.write "Unable to find the specified report"
    else
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open
	    set rs = server.CreateObject("ADODB.recordset")

        if request("StatusID") = "7" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (2) and s.status_Name = 'Open'  " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.ExcaliburID in (Select id from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.Division_ID=6 " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "P2 Backlog"
        elseif request("StatusID") = "1" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.Status_Name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.ExcaliburID in (Select id from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _ 
                     "and s.Division_ID=6 " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "P1 Backlog"
        elseif request("StatusID") = "2" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.status_name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.state_name in ('Cannot Duplicate','Cannot Duplicate – Disagree','Duplicate – Disagree','Fix Failed','Need Info','New*/Reopen','No Fix Needed – Disagree','Transfer Requested','Under Investigation','Will Not Fix - Disagree') " & _
                     "and s.excaliburID in (Select id from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.Division_ID=6 " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "Investivating - P1"
        elseif request("StatusID") = "5" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.status_name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.state_name not in ('Cannot Duplicate','Cannot Duplicate – Disagree','Duplicate – Disagree','Fix Failed','Need Info','New*/Reopen','No Fix Needed – Disagree','Transfer Requested','Under Investigation','Will Not Fix - Disagree','Understood/Problem Identified','Fix in Progress','Fix in Progress - Waiting on Vendor') " & _
                     "and s.Division_ID=6 " & _
                     "and s.ExcaliburID in (Select ID from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "Retest - P1"
        elseif request("StatusID") = "4" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.status_name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.state_name in ('Fix in Progress','Fix in Progress - Waiting on Vendor') " & _
                     "and s.Division_ID=6 " & _
                     "and s.ExcaliburID in (Select ID from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "Fix In Progress - P1"
        elseif request("StatusID") = "3" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.status_name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and s.state_name in ('Understood/Problem Identified') " & _
                     "and s.Division_ID=6 " & _
                     "and s.ExcaliburID in (Select ID from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "Identified - P1"
        elseif request("StatusID") = "6" then
            strSQl = "Select s.Observation_ID as ObservationID, s.Comp_Part_Name AS  Component, s.Platform_Cycle_Version as Product, s.State_Name as State, o.Short_Description as Summary, s.report_date  " & _
                     "from  HOUSIREPORT01.DataWarehouse.dbo.SI_Snapshot_Observation s  with (NOLOCK), HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK) " & _
                     "where s.priority_name in (0,1) and s.status_name = 'Open' " & _
                     "and o.Observation_ID = s.Observation_ID " & _
                     "and coalesce(s.gating_name,'') = 'Web Release' " & _
                     "and s.Division_ID=6 " & _
                     "and s.ExcaliburID in (Select ID from deliverableversion with (NOLOCK) where deliverablerootid = " & clng(request("RootID")) & ") " & _
                     "and s.report_date between getDate()-1 and getDate()"
            strReportName = "Web - P1"
        else
            strSQL = ""
        end if

        if strSQL = "" then
            response.write "Unable to find the specified report"
        else
            blnHeaderWritten = false
            rs.open strSQL,cn
            do while not rs.eof
                if not blnHeaderWritten then
                    Response.write "<font size=2 face=verdana><b>" & rs("Component") & " - " & strReportname & "</b><br><br></font>"
                    response.write "<font size=1>Report Date: " & rs("Report_Date") & "</font><br><br>"
                    response.write "<table bgcolor=ivory cellpadding=2  border=1 bordercolor=gainsboro cellspacing=0 width=""100%"">"
                    response.write "<tr style=""background-color: beige""><td><b>OTS&nbsp;ID</b></td><td><b>Product</b></td><td><b>State</b></td><td><b>Summary</b></td></tr>"
                    blnHeaderWritten = true 
                end if
                response.write "<tr>"
                response.write "<td valign=top><a target=_blank href=""../search/ots/report.asp?txtReportSections=1&txtObservationID=" & rs("observationid") & """>" & rs("observationid") & "</a></td>"
                response.write "<td valign=top nowrap>" & rs("product") & "</td>"
                response.write "<td valign=top nowrap>" & rs("State") & "</td>"
                response.write "<td>" & rs("Summary") & "</td>"
                response.write "</tr>"
                rs.movenext
            loop
            rs.close    
            if blnHeaderWritten then
                response.write "</table>"
            end if
        end if
    

        set rs = nothing
        cn.Close
        set cn = nothing
    end if
%>
</BODY>
</HTML>





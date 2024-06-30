<%@ Language="VBScript" %>

	<%
	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
    else	
        Response.Buffer = True
        Response.ExpiresAbsolute = Now() - 1
        Response.Expires = 0
        Response.CacheControl = "no-cache"
    end if	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->


<html>
<head>
    <%if request("Title")=""then%>
    <title>Commodity Yearly Reports</title>
    <%else%>
    <title><%=request("Title")%></title>
    <%end if%>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
        
    //-->
</SCRIPT>
</head>

<STYLE>
    td{
        FONT-FAMILY: Verdana;   
        FONT-SIZE: x-small;
    }
    A:visited
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }    
</STYLE>

<body bgcolor=ivory>
<%
    if request("lstCategories") = "" or request("txtStartDate") = "" or request("txtEndDate")=""then
        response.write "Not enough information supplied to create this report."
    else
%>

<table width=100% cellpadding=1 cellspacing=0 border=1>
<TR>
        <TD><b>ID</b></TD>
        <TD><b>Deliverable&nbsp;Name</b></TD>
        <TD><b>Part&nbsp;Number</b></TD>
        <TD><b>QComplete&nbsp;Product</b></TD>
        <TD><b>Workflow&nbsp;Complete</b></TD>
        <TD><b>First&nbsp;QComplete</b></TD>
</TR>
<%
    dim cn, rs, strSQL,i
    dim strCategory
    dim CategoryArray
    dim strCategoryIDList
    
    CategoryArray = split(request("lstCategories") & "",",")
    
    strCategoryIDList=""

    for each strCategory in CategoryArray
        strCategoryIDList = strCategoryIDList & "," & trim(clng(strcategory))
    next
    if strCategoryIDList <> "" then
        strCategoryIDList = mid(strCategoryIDList,2)
    end if
    
    
    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.CommandTimeout = 120
	cn.IsolationLevel=256
    cn.Open

    set rs = server.CreateObject("ADODB.recordset")
    strSQl = "Select v.id , v.deliverablename, v.partnumber, s.actual, q.product, q.FirstProductQComplete, datediff(d,actual,FirstProductQComplete) " & _
             "from deliverableversion v, deliverableschedule s, deliverableroot r, deliverablecategory c,  " & _
							"(Select v.id as VersionID, max(s.id) as WorkflowID " & _
							"from deliverableversion v, deliverableroot r, deliverableschedule s " & _
							"where v.deliverablerootid = r.id " & _
							"and r.typeid =1 " & _
							"and s.deliverableversionid = v.id " & _
							"group by v.id) ms, " & _
							"( " & _
							"Select l.deliverableversionid, p.dotsname as product, l.updated as FirstProductQComplete " & _
							"from productversion p, actionlog l, " & _
								"( " & _
								"Select l.DeliverableVersionID, min(l.id) as ID " & _
								"from actionlog l " & _
								"where actionid=21 " & _
								"and toID=5 " & _
								"and updated between '" & cdate(request("txtStartDate")) & "' and '" & dateadd("d",1,cdate(request("txtEndDate"))) & "' " & _
								"group by l.DeliverableVersionID " & _
								") q " & _
							"where l.id = q.id " & _
							"and p.id = l.productversionid " & _
							"and l.deliverableversionid = q.deliverableversionid	 " & _
							") q " & _
            "where v.location = 'Workflow Complete' " & _
            "and c.id in (" & strCategoryIDList & ") " & _
            "and s.id = ms.workflowid " & _
            "and s.deliverableversionid = v.id " & _
            "and c.id = r.categoryid " & _
            "and r.id = v.deliverablerootid " & _
            "and actual between '" & cdate(request("txtStartDate")) & "' and '" & dateadd("d",1,cdate(request("txtEndDate"))) & "' " & _
            "and ms.versionid = v.id " & _
            "and v.filename not like 'HFCN%' " & _
            "and q.deliverableversionid = v.id " & _
            "order by datediff(d,actual,FirstProductQComplete)"

    
    rs.open strSQL,cn
    do while not rs.eof
            response.write "<tr>"
            response.write "<td>" & rs("ID") & "</td>"
            response.write "<td>" & rs("DeliverableName") & "</td>"
            response.write "<td>" & rs("PartNumber") & "</td>"
            response.write "<td>" & rs("Product") & "</td>"
            response.write "<td>" & formatdatetime(rs("Actual"),vbshortdate) & "</td>"
            response.write "<td>" & formatdatetime(rs("FirstProductQComplete"),vbshortdate) & "</td>"
            response.write "</tr>"

        rs.movenext
    loop
    rs.close


%>


</table>

<%
    set rs = nothing
    cn.close
    set cn = nothing

    end if
%>

</body>
</html>




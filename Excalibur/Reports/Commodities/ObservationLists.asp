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
        <TD><b>OTSID</b></TD>
        <TD><b>Developer</b></TD>
        <TD><b>Category</b></TD>
        <TD><b>PartNumber</b></TD>
        <TD><b>Priority</b></TD>
        <TD><b>Summary</b></TD>
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
    'OTS Before    
    strSQl =   "Select o.observationid, o.developer, v.category, v.partnumber, o.priority, o.shortdescription, (case when o.dateCreated < datediff(s,'1/1/1970',v.qcompletedate) then 1 end) as Before,(case when o.dateCreated > datediff(s,'1/1/1970',v.qcompletedate) then 1 end) as After  " & _
                "from ots_observation_tracking o, " & _
                "( " & _
                "Select c.name as Category, v.partnumber, v.OTSPartNumber, p.dotsname, max(l.Updated) as QCompleteDate " & _
                "from deliverableversion v, productversion p, deliverablecategory c, deliverableroot r, ( " & _
                "            Select DeliverableVersionID,details, productversionid, Updated " & _
							"from actionlog " & _
							"where actionid=21 " & _
							"and toID=5 " & _
							") l " & _
	            "where c.id = r.categoryid " & _
	            "and v.deliverablerootid= r.id " & _
                "and c.id in (" & strCategoryIDList & ") " & _
	            "and l.productversionid <> 100 " & _
	            "and l.deliverableversionid = v.id " & _
	            "and p.id = l.productversionid " & _
	            "group by c.name, v.otspartnumber,v.partnumber, p.dotsname " & _
	            "having max(l.Updated) between '" & cdate(request("txtStartDate")) & "' and '" & cdate(request("txtEndDate")) & "' " & _
	            ") v " & _
                "where  o.partnumber = v.otspartnumber " & _
                "and (o.platform=v.dotsname or o.systemboard=v.dotsname) " & _
                "and v.category <> 'Input Device' " & _
                "and dateadd(s,o.dateCreated,'1/1/1970') between '" & cdate(request("txtStartDate")) & "' and v.qcompletedate  " & _
                "order by v.category, v.partnumber"
    
    response.write strSQL
    response.flush
    rs.open strSQL,cn,adOpenStatic
    do while not rs.eof
        response.write "<tr>"
        response.write "<td>" & rs("Category") & "</td>"
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

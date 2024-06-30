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
    <%if request("Report") = "1" or request("Report") = "2" then%>
        <TD><b>Category</b></TD>
        <TD><b>All Products</b></TD>
        <TD><b>Commercial</b></TD>
        <TD><b>Consumer</b></TD>
        <TD><b>Week</b></TD>
        <TD><b>Year</b></TD>
    <%else%>
        <TD><b>Category</b></TD>
        <TD><b>Versions</b></TD>
        <TD><b>Week</b></TD>
        <TD><b>Year</b></TD>
    <%end if%>
</TR>
<%
    dim cn, rs, strSQL,i
    dim LastWeek
    dim LastYear
    dim CategoriesID
    dim CategoriesName
    dim CategoriesValue
    dim strCategory
    dim strIDList
    dim Cells
    dim CellValue
    
    LastWeek = 0
    LastYear = 0
    
    strIDList = ""
    
    CategoriesID = split(request("lstCategories"),",")
    CategoriesName = split(request("lstCategories"),",")
    CategoriesValue = split(request("lstCategories"),",")
    
    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
    cn.Open

    set rs = server.CreateObject("ADODB.recordset")
    
    i=0
    for each strCategory in CategoriesID
        strSQl = "spGetDeliverableCategoryName " & clng(strCategory)
        rs.open strSQl, cn
        if rs.eof and rs.bof then
            CategoriesName(i) = ""
        else
            CategoriesName(i) = rs("Name")
        end if
        rs.close
        if request("Report") = "3" then
            CategoriesValue(i) = "0"
        else
            CategoriesValue(i) = "0|0|0"
        end if
        strIDList = strIDList & "," & clng(strCategory)
        i=i+1
    next
    if strIDList <> "" then
        strIDList = mid(strIDList,2)
    end if
    if request("Report") = "1" then
        strSQL = "Select c.name as Category, c.id as Categoryid, count(1) as VersionProductCount, count(CASE WHEN p.devcenter<> 2 THEN 1 END) as CommercialProducts,count(CASE WHEN p.devcenter = 2 THEN 1 END) as ConsumerProducts,datepart(""ww"",l.updated ) as ReportWeek,datepart(""yyyy"",l.updated ) as ReportYear " & _
                 "from actionlog l, deliverableversion v, deliverableroot r, deliverablecategory c, productversion p " & _
                "where r.id = v.deliverablerootid " & _
                "and l.deliverableversionid = v.id " & _
                "and l.productversionid = p.id " & _
                "and actionid= 19 " & _
                "and r.categoryid in (" & strIDList & ") " & _
                "and l.updated between '" & cdate(request("txtStartDate")) & "' and '" & dateadd("d",1,cdate(request("txtEndDate"))) & "' " & _
                "and c.id = r.categoryid " & _
                "group by c.name,c.id,datepart(""ww"",l.updated), datepart(""yyyy"",l.updated ) " & _
                "order by ReportYear, Reportweek, category"
    elseif request("Report") = "2" then
        strSQL = "Select c.name as Category, c.id as Categoryid, count(1) as VersionProductCount, count(CASE WHEN p.devcenter<> 2 THEN 1 END) as CommercialProducts,count(CASE WHEN p.devcenter = 2 THEN 1 END) as ConsumerProducts,datepart(""ww"",l.updated ) as ReportWeek,datepart(""yyyy"",l.updated ) as ReportYear " & _
                 "from actionlog l, deliverableversion v, deliverableroot r, deliverablecategory c, productversion p " & _
                "where r.id = v.deliverablerootid " & _
                "and l.deliverableversionid = v.id " & _
                "and l.productversionid = p.id " & _
                "and actionid= 21 " & _
                "and toID=5 " & _
                "and r.categoryid in (" & strIDList & ") " & _
                "and l.updated between '" & cdate(request("txtStartDate")) & "' and '" & dateadd("d",1,cdate(request("txtEndDate"))) & "' " & _
                "and c.id = r.categoryid " & _
                "group by c.name,c.id,datepart(""ww"",l.updated), datepart(""yyyy"",l.updated ) " & _
                "order by ReportYear, Reportweek, category"
    elseif request("Report") = "3" then
        strSQL = "Select count(1) as VersionProductCount, c.name as Category, c.id as categoryid, datepart(""ww"",v.created) as ReportWeek, datepart(""yyyy"",v.created) as ReportYear " & _
                 "from deliverableversion v, deliverableroot r, deliverablecategory c " & _
                 "where v.created between '" & cdate(request("txtStartDate")) & "' and '" & dateadd("d",1,cdate(request("txtEndDate"))) & "' " & _
                 "and v.deliverablerootid = r.id " & _
                 "and c.id = r.categoryid " & _
                 "and r.categoryid in (" & strIDList & ") " & _
                 "group by datepart(""ww"",v.created), datepart(""yyyy"",v.created), c.name, c.id " & _
                 "order by datepart(""yyyy"",v.created), datepart(""ww"",v.created), c.name "
    end if
    rs.open strSQL,cn
    do while not rs.eof
        if (trim(lastweek) <> trim(rs("ReportWeek")) or trim(lastyear) <> trim(rs("ReportYear"))) and trim(lastyear) <> 0 then
            for i = 0 to ubound(CategoriesValue)
                if request("chkZeros") = "" or (CategoriesValue(i) <> "0|0|0" and categoriesValue(i) <>"0")then
                    Cells = split(CategoriesValue(i),"|")
                    response.write "<TR><TD>" &  CategoriesName(i) & "</TD>"
                    for each CellValue in cells
                        response.write "<TD>" &  CellValue & "</TD>"
                    next
                    response.write "<TD>" &  LastWeek & "</TD>"
                    response.write "<TD>" &  LastYear & "</TD>"
                    response.write "</TR>"
                end if
                if request("Report") = "3" then
                    CategoriesValue(i) = "0"
                else
                    CategoriesValue(i) = "0|0|0"
                end if
            next

        end if
        lastweek = rs("ReportWeek")
        lastyear = rs("ReportYear") 
        for i = 0 to ubound(CategoriesID)
            if trim(CategoriesID(i)) = trim(rs("CategoryID")) then
                if request("Report") = "3" then
                    CategoriesValue(i) = rs("VersionProductCount")
                else
                    CategoriesValue(i) = rs("VersionProductCount") & "|" & rs("CommercialProducts") & "|" & rs("ConsumerProducts")
                end if
                exit for
            end if
        next
        rs.movenext
    loop
    rs.close

    for i = 0 to ubound(CategoriesValue)
        if request("chkZeros") = "" or (CategoriesValue(i) <> "0|0|0" and CategoriesValue(i) <> "0" )then
            Cells = split(CategoriesValue(i),"|")
            response.write "<TR><TD>" &  CategoriesName(i) & "</TD>"
            for each CellValue in cells
                response.write "<TD>" &  CellValue & "</TD>"
            next
            response.write "<TD>" &  LastWeek & "</TD>"
            response.write "<TD>" &  LastYear & "</TD>"
            response.write "</TR>"
        end if
        if request("Report") = "3" then
            CategoriesValue(i) = "0"
        else
            CategoriesValue(i) = "0|0|0"
        end if
   next


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


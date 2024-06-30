<%@ Language=VBScript %>
<!-- #include file = "../includes/noaccess.inc" -->
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

<BODY>
<h5>System BIOS Observations</h5>
<font size=1 face=verdana>10/1/2010 - 9/30/2011</font><br><br>
<%
    
	dim cn, rs, strSQL
    dim strLastProduct
    dim MonthArray
    dim Total1,Total2,Total3,Total4

    MonthArray = split("xx,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")

    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    strSQL = "SELECT count(1) as Total,  count(case when priority=0 or priority=1  then 1 end) as P1, count(case when priority=2 then 1 end) as P2,count(case when priority=3 or priority=4 or priority=5 then 1 end) as P345, PrimaryProduct as Product, datepart(mm,dateopened) as Month, datepart(yyyy,DateOpened) as [year] " & _
            "FROM HOUSIREPORT01.SIO.dbo.SI_Observation_Report (NOLOCK) " & _
            "where Subsystem = 'System BIOS' " & _
            "and DivisionID=6 " & _
            "and dateopened between '10/1/2010' and '10/1/2011' " & _
            "group by PrimaryProduct, datepart(yyyy,dateOpened), datepart(mm,dateOpened) " & _
            "order by PrimaryProduct, datepart(yyyy,dateOpened), datepart(mm,dateOpened) " 
    rs.open strSQL,cn,adOpenStatic
    strLastProduct = ""
    response.write "<table width=300>"
    Total1 = 0
    Total2 = 0
    Total3 = 0
    Total4 = 0

    do while not rs.EOF
        if strLastProduct <> trim(rs("Product") & "") then
            if strLastProduct <> "" then
                response.write "<tr bgcolor=darkseagreen><td>Total:</td><td align=center>" & Total1 & "</td><td align=center>" & Total2 & "</td><td align=center>" & Total3 & "</td><td align=center>" & Total4 & "</td></tr>"
                response.write "<tr><td colspan=5>&nbsp;</td></tr>"
                Total1 = 0
                Total2 = 0
                Total3 = 0
                Total4 = 0
            end if
            response.write "<tr bgcolor=lightsteelblue><td colspan=5>" & trim(rs("Product") & "") & "</td></tr>"
            response.write "<tr bgcolor=gainsboro><td>Month</td><td align=center>P1</td><td align=center>P2</td><td align=center>P3/4/5</td><td align=center>Total</td></tr>"
        end if
        strLastProduct = trim(rs("Product") & "")
        response.write "<tr>"
        response.write "<td>" & MonthArray(trim(rs("Month") & "")) & " " & rs("Year") & "</td>"
        response.write "<td align=center>" & rs("P1") & "" & "</td>"
        response.write "<td align=center>" & rs("P2") & "" & "</td>"
        response.write "<td align=center>" & rs("P345") & "" & "</td>"
        response.write "<td align=center>" & rs("Total") & "" & "</td>"
        response.write "</tr>"
        Total1 = Total1 + rs("P1")
        Total2 = Total2 + rs("P2")
        Total3 = Total3 + rs("P345")
        Total4 = Total4 + rs("Total")
        rs.MoveNext
    loop
    rs.Close

    if strLastProduct <> "" then
        response.write "<tr bgcolor=darkseagreen><td>Total:</td><td align=center>" & Total1 & "</td><td align=center>" & Total2 & "</td><td align=center>" & Total3 & "</td><td align=center>" & Total4 & "</td></tr>"
        response.write "<tr><td colspan=5>&nbsp;</td></tr>"

        strSQL = "SELECT  count(1) as Total,  count(case when priority=0 or priority=1  then 1 end) as P1, count(case when priority=2 then 1 end) as P2,count(case when priority=3 or priority=4 or priority=5 then 1 end) as P345, datepart(mm,dateopened) as Month, datepart(yyyy,DateOpened) as [year] " & _
                    "FROM HOUSIREPORT01.SIO.dbo.SI_Observation_Report (NOLOCK) " & _
                    "where Subsystem = 'System BIOS'  " & _
                    "and DivisionID=6 " & _
                    "and dateopened between '10/1/2010' and '10/1/2011'  " & _
                    "group by datepart(yyyy,dateOpened), datepart(mm,dateOpened) " & _
                    "order by datepart(yyyy,dateOpened), datepart(mm,dateOpened)"
        rs.open strSQL,cn,adOpenStatic
        if not (rs.eof and rs.bof) then
            response.write "<tr bgcolor=lightsteelblue><td colspan=5>All Products</td></tr>"
            response.write "<tr bgcolor=gainsboro><td>Month</td><td align=center>P1</td><td align=center>P2</td><td align=center>P3/4/5</td><td align=center>Total</td></tr>"
            Total1 = 0
            Total2 = 0
            Total3 = 0
            Total4 = 0
            do while not rs.EOF
                response.write "<tr>"
                response.write "<td>" & MonthArray(trim(rs("Month") & "")) & " " & rs("Year") & "</td>"
                response.write "<td align=center>" & rs("P1") & "" & "</td>"
                response.write "<td align=center>" & rs("P2") & "" & "</td>"
                response.write "<td align=center>" & rs("P345") & "" & "</td>"
                response.write "<td align=center>" & rs("Total") & "" & "</td>"
                response.write "</tr>"
                Total1 = Total1 + rs("P1")
                Total2 = Total2 + rs("P2")
                Total3 = Total3 + rs("P345")
                Total4 = Total4 + rs("Total")
            rs.MoveNext
            loop
            response.write "<tr bgcolor=darkseagreen><td>Total:</td><td align=center>" & Total1 & "</td><td align=center>" & Total2 & "</td><td align=center>" & Total3 & "</td><td align=center>" & Total4 & "</td></tr>"
            rs.Close
        end if
    end if
    response.write "</table>"

    set rs = nothing
    cn.Close
    set cn = nothing
%>
</BODY>
</HTML>





<%@ Language=VBScript %>

<html>
<head>
    <title></title>
</head>
<body>

<%

	dim cn
	dim rs
	dim strSQL
	dim strAllVersions
	strRowBorderColor = "LavenderBlush"
	strAllVersions = ""

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")


    rs.open "Select r.id as RoHSID, i.partnumber, i.name " & _
            "from " & _
	        "( " & _
	        "Select partnumber,cast(month(mydate) as varchar(2)) + '/' + cast(day(mydate) as varchar(2)) + '/' + right(year(mydate),2) as Name " & _
        	"from datawarehouse.dbo.sheet1$ with (NOLOCK) " & _
	        ") i,prs.dbo.rohs r " & _
	        "where  r.name = i.name ",cn
	do while not rs.eof
	    response.write rs("Partnumber") & "[" & rs("RoHSID") & "]<br>"
    	set rs2 = server.CreateObject("ADODB.recordset")
	    rs2.open "Select ID from deliverableversion with (NOLOCK) where partnumber = '" & rs("Partnumber") & "'",cn
        if rs2.eof and rs2.bof then
            response.write "Error: No version found with this part number.<BR>"    
        else
            do while not rs2.eof
                response.write "update deliverableversion set ROHSID=" & rs("RoHSID")& " where id=" & rs2("ID") & "<BR>"   
                'cn.execute "update deliverableversion set ROHSID=" & rs("RoHSID")& " where id=" & rs2("ID")
                strAllVersions = strAllVersions & "," & rs2("ID")
                rs2.movenext
            loop
            response.write "<BR>"
        end if
        rs2.close
	    set rs2 = nothing
	    rs.movenext
	loop
	rs.close
	set rs = nothing
	cn.close
	set cn = nothing
	
response.write strAllVersions
%>

</body>
</html>

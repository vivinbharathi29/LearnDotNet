<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<html>
<head>

<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<title>IRS Components</title>
<script type="text/javascript" id="clientEventHandlersJS" language="javascript">
<!--

//-->
</script>
<style>
    td
    {
         font-family: Verdana;
         font-size: xx-small;
    }
    body
    {
         font-family: Verdana;
         font-size: x-small;
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
</head>

<body bgcolor="white">

<%
	dim cn
	dim rs

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    if Request("ID") = "" then
        response.write "Product ID not specified."
    else
    	rs.Open "spGetProductVersion_Pulsar " & clng(Request("ID")),cn,adOpenForwardOnly

        if rs.eof and rs.bof then
            response.write "Requested product not found."
            rs.close
        else
            response.write "<font face=""verdana"" size=2><b>IRS Component List - " & rs("DotsName") & "</b><br /></font>"
	        rs.close
            response.write "<div style=""margin-top:6px;margin-bottom:6px""><font face=""verdana"" size=2>Display: "
            if request("OS") = "Win7" then
                response.write "<a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & """>All</a> , Win7 , <a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & "&OS=Win8"">Win8</a>"
            elseif request("OS") = "Win8" then
                response.write "<a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & """>All</a> , <a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & "&OS=Win7"">Win7</a> , Win8"
            else
                response.write "All , <a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & "&OS=Win7"">Win7</a> , <a href=""IRSComponentList.asp?ID=" & clng(Request("ID")) & "&OS=Win8"">Win8</a>"
            end if
            response.write "</font></div>"

            dim strProductOSFamilyColumns
	        dim ProductOSFamilyArray
            
            dim strDeliverableOSFamilyList
            strDeliverableOSFamilyList = ""
    
            set rs = server.CreateObject("ADODB.recordset")
            rs.open "spListProductOSFamiliesPreinstalled " & clng(Request("ID")) & ",2",cn,adOpenStatic
	        strProductOSFamilyColumns = ""
	        do while not rs.eof
	            if rs("Name") <> "FD" then
	                strProductOSFamilyColumns = strProductOSFamilyColumns & "," & rs("name")
	            end if
	            rs.movenext
	        loop
	        rs.close
	        if strProductOSFamilyColumns <> "" then
	            strProductOSFamilyColumns = mid(strProductOSFamilyColumns,2)
	        end if
	        ProductOSFamilyArray = split(strProductOSFamilyColumns,",")


        	rs.Open " spListProductComponentsInIRS " & clng(Request("ID")),cn,adOpenForwardOnly
	        if rs.EOF and not rs.BOF then
                response.write "No components found for the selected product."
            else
%>
	            <table id="tabParts" width="100%" bgcolor="white" border="1" cellspacing="0" cellpadding="2" style="border-color: Gray">
		        <tr>
	    	        <td style=" background-color:gainsboro; white-space:nowrap"><b>IRS Part Number</b></td>
                	<%
	                    for i = 0 to ubound(ProductOSFamilyArray)
	                        response.write 	"<td width=""10"" style=""background-color:gainsboro;white-space:nowrap""><b>" & ProductOSFamilyArray(i) & "&nbsp;&nbsp;</b></td>"
	                    next
	                %>

	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Excalibur ID</b></td>
	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Name</b></td>
	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Version</b></td>
	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Rev</b></td>
   	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Pass</b></td>
   	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Target Notes</b></td>
   	    	        <td style="background-color:gainsboro;white-space:nowrap"><b>Image Summary</b></td>
		        </tr>
                <% 
                do while not rs.eof

                    strSQL = trim(rs("images") & "")
                    if strSQl = "" then
                        strDeliverableOSFamilyList = strProductOSFamilyColumns
                    else
                        if instr(strSQL, ":")> 0 then
                            strSQl = left(strSQl,instr(strSQL, ":")-1)
                        end if
                        if right(trim(strSQl),1) = "," then
                            strSql = strSQl & "0"
                        end if
                        strSQL = "Select distinct f.shortname as Name, f.id " & _
                                    "from images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK), osfamily f with (NOLOCK) " & _
                                    "where d.id = i.imagedefinitionid " & _
	                                "and f.id = osfamilyid " & _
	                                "and d.osid = o.id " & _
                                    "and i.id in (" & strSQL & ")"
                        set rs2 = server.CreateObject("ADODB.recordset")

	                    rs2.open strSQL,cn,adOpenStatic
	                    if (rs2.eof and rs2.bof) then
    	                    strDeliverableOSFamilyList = ""
                        else

                            strDeliverableOSFamilyList = ""
                            do while not rs2.eof
                                strDeliverableOSFamilyList = strDeliverableOSFamilyList & "," & rs2("Name")
                                rs2.movenext
                            loop
                        end if
	                    rs2.close
	                    set rs2 = nothing
                        if strDeliverableOSFamilyList <> "" then
                            strDeliverableOSFamilyList = mid(strDeliverableOSFamilyList,2)
                        end if
	                end if
	                strDeliverableOSFamilyList = "," & replace(strDeliverableOSFamilyList," ","") & ","


                  
                   if instr(lcase(strDeliverableOSFamilyList),"," & lcase(trim( request("OS") ))) > 0  or trim(request("OS")) = ""then
                        response.write "<tr>"
                    else
                        response.write "<tr style=""display:none"">"
                    end if
                  
                  
                    response.write "<td>" & rs("PartNumber") & "</td>"
            	    for i = 0 to ubound(ProductOSFamilyArray)
            	        if instr(lcase(strDeliverableOSFamilyList),"," & lcase(trim(ProductOSFamilyArray(i)))) > 0 then
	                        response.write 	"<td align=""center"">X</td>"
	                    else
	                        response.write 	"<td>&nbsp;</td>"
                        end if
	                next

                    response.write "<td>" & rs("ID") & "</td>"
                    response.write "<td>" & rs("Name") & "</td>"
                    response.write "<td>" & rs("Version") & "</td>"
                    response.write "<td>" & rs("Revision") & "&nbsp;</td>"
                    response.write "<td>" & rs("Pass") & "&nbsp;</td>"
                    response.write "<td>" & rs("TargetNotes") & "&nbsp;</td>"
                    response.write "<td>" & rs("ImageSummary") & "&nbsp;</td>"
                    response.write "</tr>"
                    rs.movenext
                loop
                rs.close
                %>
    	        </table>

<%
            end if
        end if
    end if

	set rs=nothing
	cn.Close
	set cn=nothing
%>

</BODY>
</HTML>

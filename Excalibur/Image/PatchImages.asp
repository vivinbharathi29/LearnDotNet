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

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
    padding: 2px 2px 2px 2px;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
tr.header {
   background-color:Beige;   
   font-weight:bold  
  }
  table
  {
      background-color: Ivory;
  }
h1
{
    FONT-FAMILY: Verdana;
    FONT-SIZE:small;
    font-weight: bold;
}

h2
{
    FONT-FAMILY: Verdana;
    FONT-SIZE:x-small;
    font-weight: bold;
}
</STYLE>


</HEAD>


<BODY>

<%
    if trim(request("ProdID")) = "" or trim(request("DelID")) = "" then
        response.write "Unable to find the requested information."
    elseif not (isnumeric(request("ProdID")) or isnumeric(request("DelID")) ) then
        response.write "Unable to find the requested information."
    else
        dim cn, rs
        dim ProdID
        dim DelID
        ProdID = clng(request("ProdID"))
        DelID = clng(request("DelID"))

	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open
	    set rs = server.CreateObject("ADODB.recordset")

        dim strProductName
        dim strDeliverable

        rs.open "spGetProductDeliverable " & ProdID & "," & DelID,cn
        if rs.eof and rs.bof then
            response.write "Unable to find the requested information."
            rs.close
        else
            response.write "<h1>Patch Image List</h1><h2>" & rs("Deliverable") & " on " & rs("product") & "</h2>"
            rs.close
            rs.open "spListPatchImages " & ProdID & "," & DelID,cn
            if not (rs.eof and rs.bof) then
                response.write "<table cellpadding=2 cellspacing=0 border=1 bordercolor=gainsboro>"
                response.write "<tr class=header>"
                response.write "<td>SKU</td>"
                response.write "<td>Priority</td>"
                response.write "<td>Region</td>"
                response.write "<td>Code</td>"
                response.write "<td>Brand</td>"
                response.write "<td>OS</td>"
                response.write "<td>Apps&nbsp;Bundle</td>"
                response.write "<td>Image</td>"
                response.write "</tr>"
            end if
            do while not rs.eof
                response.write "<tr>"
                response.write "<td>" & rs("SkuNumber") & "</td>"
                response.write "<td>" & rs("Priority") & "</td>"
                response.write "<td>" & rs("Region") & "</td>"
                response.write "<td>" & rs("OptionConfig") & "</td>"
                response.write "<td>" & rs("Brand") & "</td>"
                response.write "<td>" & rs("OS") & "</td>"
                response.write "<td>" & rs("SWType") & "</td>"
                response.write "<td>" & rs("ImageType") & "</td>"
                response.write "</tr>"
                rs.movenext
            loop
            if not (rs.eof and rs.bof) then
                response.write "</table>"
            else
                response.write "No images specified for this patch."
            end if
            rs.close
    
        end if

        set rs = nothing
        cn.Close
        set cn = nothing
    end if
%>
</BODY>
</HTML>





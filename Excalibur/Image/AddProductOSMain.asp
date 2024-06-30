<%@ Language="VBScript" %>

<html>
<head>
</head>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<body bgcolor=ivory>
    <form id=frmMain method=post action="AddProductOSSave.asp">

<% 
    dim strProductID
    dim strProductName
    
    strProductID = request("ProductID")
    if not isnumeric(strProductID ) then
        response.write "Unable to process your request."
    else
      	set cn = server.CreateObject("ADODB.Connection")
    	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open
	    set rs = server.CreateObject("ADODB.recordset")

        rs.open "spGetProductVersionName " & clng(strProductID), cn
        if rs.bof and rs.eof then
            strProductName = ""
        else
            strProductName = rs("Name") & ""
        end if
        rs.close
        if strProductName = "" then
            response.write "Unable to find the requested product information."
        else
            response.write "<h3>" & strProductname & " - Add Supported OS for Preinstall</h3>"
            
%>
	<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr ID=AddOSRow>
		<td width=10 nowrap valign=top><b>Add&nbsp;OS&nbsp;Support:</b></td>
		<td>
 			<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 400px; BACKGROUND-COLOR: white" id=DIV2>
			<TABLE width=100% ID=TableOS>
				<THEAD><tr  style="position:relative;top:expression(document.getElementById('DIV2').scrollTop-2);"><TD bgcolor=#c9ddff width=10 style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;</TD><TD width=282 style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"  bgcolor=#c9ddff>&nbsp;OS Name</TD></tr></THEAD>
                <TBODY>


<%            
            rs.open "spListProductOSAll " & clng(strProductID),cn
            do while not rs.eof
                if isnull(rs("Preinstall")) or not rs("Preinstall") then
					Response.Write "<TR>"
					Response.Write "<TD><input style=""width:16px;height:16px"" id=""chkOS"" name=""chkOS"" type=""checkbox"" value=""" & rs("ID") & """>&nbsp;&nbsp;</TD>"
					Response.Write "<TD width=""100%"">" & rs("shortName") & "</TD>"
					Response.Write "</TR>"
                end if
                rs.movenext
            loop
            rs.close
%>
            </TBODY>
            </TABLE>
           </DIV>
       </td></tr></table>     
       
<%            
        end if
        cn.close
        set cn = nothing
    
    end if


%>
        <input style="display:none" id="txtProductID" name="txtProductID" type="text" value="<%=strProductID%>">
        <input style="display:none" id="txtProductName" name="txtProductName" type="text" value="<%=strProductName%>">
       </form>

</body>
</html>


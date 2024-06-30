<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<style>
    td
    {
    FONT-Family: Verdana;
    FONT-Size:xx-small
    }
    th
    {
    FONT-Family: Verdana;
    FONT-Size:xx-small
    }
</style>
<BODY>
<font face=verdana size=2>

<b>Virtual Products in Sudden Impact</b><br><br>

<font size=1 face=verdana>Items in red are in Excalibur but have not been verifed to be in Sudden Impact yet.</font><br><br>

<%

Server.ScriptTimeout = 5400

	dim cn
	dim rs 
	dim strSQL
	dim strUpdate
	dim strFontColor

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")

    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.IsolationLevel=256
	cn.Open


    strSQl = "spListOTSVirtualProducts"
    
    rs.open strSQL,cn
    if not (rs.eof and rs.bof) then
        response.write "<table cellspacing=0 cellpadding=2 border=1 width=""100%"">"
        response.write "<tr bgcolor=beige>"
        response.write "<th>ID</th>"
        response.write "<th>Primary&nbsp;Product</th>"
        response.write "<th>Product&nbsp;Version</th>"
        response.write "<th>ODM</th>"
        response.write "<th>PM&nbsp;Email</th>"
        response.write "<th>Test&nbsp;Lead&nbsp;Email</th>"
        response.write "</tr>"
    end if
    do while not rs.eof
        if isnull(rs("SIID")) then
            strFontColor = "red"
        else
            strFontColor = "black"
        end if
        response.write "<tr bgcolor=ivory>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("ID") & "</font></td>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("PrimaryProduct") & "</font></td>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("ProductVersion") & "</font></td>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("ODM") & "&nbsp;</font></td>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("PMEmail") & "</font></td>"
        response.write "<td><font color=""" & strFontColor & """>" & rs("TestLeadEmail") & "</font></td>"
        response.write "</tr>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
        response.write "</table>"
    end if
    rs.close
    
    set rs = nothing
    cn.close
    Set cn = nothing
%>
</font>
</BODY>
</HTML>

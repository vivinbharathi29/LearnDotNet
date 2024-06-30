<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
		window.opener='X';
		window.open('','_parent','')
		window.close();
}


//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font face=verdana size=2>
<%

Server.ScriptTimeout = 5400

	dim cn
	dim strID
	dim rs 
	dim strSQL
	dim blnFound
	dim ProdID
	dim ProdName

	

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")

    rs.open "spGetProducts4Web -1",cn
    
    do while not rs.eof
'        if (rs("Active") = 1 or rs("Sustaining") = 1) and rs("ID") <> 100 then
        if rs("ProductStatusID") < 5 and rs("ID") <> 100 then
            cn.execute "spUpdateSuddenImpactProduct2 " & clng(rs("ID"))
            response.Write rs("ID") & ": " & rs("Name") & " " & rs("Version") & "<BR>"
            response.Flush
        end if
        rs.movenext
        
    loop
    rs.close
    
    set rs = nothing
    cn.close
    set cn = nothing
    
	Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	oMessage.From = "max.yu@hp.com"
	oMessage.To= "max.yu@hp.com"
	oMessage.Subject = "Products Updated in Sudden Impact" 

    oMessage.HTMLBody = "<font face=Arial size=2 color=black>All product properties have been updated in Sudden Impact.</font>"
		
	oMessage.Send 
	Set oMessage = Nothing 	    
%>
</font>
</BODY>
</HTML>

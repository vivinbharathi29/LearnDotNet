<%@ Language=VBScript %>
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
<STYLE>
body{
	Font-Size:xx-small;
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
    dim cn
    dim strMissingSiteList

    strMissingSiteList = ""

    'Check the Houston Server    
    set cn = server.CreateObject("ADODB.Connection")

    on error resume next
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=houbnbcvr01.auth.hpicorp.net;"    
    cn.open
    if cn.errors.count > 0 then
	strMissingSiteList = strMissingSiteList  & ",Houston (houbnbcvr01.auth.hpicorp.net)"
    end if
    cn.close
    set cn = nothing


    'Check the TDC Server    
    set cn = server.CreateObject("ADODB.Connection")

    on error resume next
    cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=16.159.144.23;"    
    cn.open
    if cn.errors.count > 0 then
	strMissingSiteList = strMissingSiteList  & ",TDC Primary (16.159.144.23)"'"16.159.144.23"'"tpopsgcvr3.auth.hpicorp.net
    end if
    cn.close
    set cn = nothing

    'Check the Brazil Server    
   ' set cn = server.CreateObject("ADODB.Connection")

  '  on error resume next
  '  cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;User ID=viewer;Initial Catalog=conveyor;Data Source=BRAHPQCVR01.auth.hpicorp.net;"     
   ' cn.open
   ' if cn.errors.count > 0 then
	'strMissingSiteList = strMissingSiteList  & ",Brazil (BRAHPQCVR01.auth.hpicorp.net)"
   ' end if
   ' cn.close
   ' set cn = nothing


    'Check the Singapore Server    
'    set cn = server.CreateObject("ADODB.Connection")

 '   on error resume next
  '  cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=sgpacccvr01.auth.hpicorp.net;" 
   ' cn.open
    'if cn.errors.count > 0 then
	'strMissingSiteList = strMissingSiteList  & ",Singapore (sgpacccvr01.auth.hpicorp.net)"
    'end if
    'cn.close
    'set cn = nothing
    
    
    'Check the CDC Server    
'    set cn = server.CreateObject("ADODB.Connection")

 '   on error resume next
  '  cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=shgcdccvr01.auth.hpicorp.net;" 
   ' cn.open
   ' if cn.errors.count > 0 then
'	strMissingSiteList = strMissingSiteList  & ",CDC (shgcdccvr01.auth.hpicorp.net)"
'    end if
'    cn.close
'    set cn = nothing


    if strMissingSiteList ="" then
        response.write "All Sites Found"
    else
	strMissingSiteList  = mid(strMissingSiteList ,2)
	'Email
	Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	oMessage.From = "max.yu@hp.com"
	oMessage.To= "max.yu@hp.com;james.moon@hp.com"
	oMessage.Subject = "Conveyor Servers Not Found" 
	oMessage.HTMLBody = "Unable to access the following Conveyor Servers: <BR><BR>" &  strMissingSiteList 
	oMessage.Send 
	Set oMessage = Nothing 	
    end if
    
  

%>


</BODY>
</HTML>



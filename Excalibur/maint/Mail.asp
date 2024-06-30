<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>

<%
	Set oMessage = New EmailWrapper
	
	oMessage.From = "max.yu@hp.com"
	
	oMessage.To= "max.yu@hp.com"
	oMessage.Subject = "Test"
						
	oMessage.HTMLBody =  "<Body><font size=2 face=verdana>Test</font></body>"
	oMessage.DSNOptions = cdoDSNFailure

	'AddAttachment Server.MapPath("..") & "\images\excalibur.jpg" ', "excalibur.gif", cdoRefTypeId

	oMessage.Send 
	Set oMessage = Nothing 

%>
</BODY>
</HTML>

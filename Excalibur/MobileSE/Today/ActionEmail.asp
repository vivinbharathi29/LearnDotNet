<%@ Language=VBScript %>
<!-- #include file="../../includes/EmailQueue.asp" -->
<html>
<head>
<title></title>
<meta  name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS"  language="javascript" type="text/javascript">
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		if (txtSuccess.value == "1")
			window.parent.close();
}
//-->
</script>
</head>
<body onload="return window_onload()">

<%
   
	if request("txtFrom") = "" or request("txtTo") = "" then
		Response.Write "<BR><font size=2 face=verdana>Not enough information supplied to send an email.</font><BR>"
		strSuccess = "0"
	else
		Set oMessage = New EmailQueue 
		
		oMessage.From = request("txtFrom")
		oMessage.To= request("txtTo")
		if request("txtCC") <> "" then
			oMessage.cc= request("txtCC")
		end if
		oMessage.Subject = request("txtSubject") 
		
		oMessage.HTMLBody = "<font face=verdana size=2 color=black>" & request("txtNotes") & "</font><BR><BR>" & request("txtEmailBody")
	
		oMessage.Send 
		Set oMessage = Nothing 	
		strSuccess = "1"
	end if
%>
<input id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>" />

</body>
</html>

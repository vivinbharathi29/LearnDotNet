<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		if (txtSuccess.value == "1")
			window.parent.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	if request("txtFrom") = "" or request("txtTo") = "" then
		Response.Write "<BR><font size=2 face=verdana>Not enough information supplied to send an email.</font><BR>"
		strSuccess = "0"
	else
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		
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
%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

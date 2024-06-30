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
		window.close();	}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%

			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
			oMessage.From = "max.yu@hp.com"
			oMessage.To= "max.yu@hp.com" 'absolutsystemteam@hp.com;
			'oMessage.CC= "pulsar.support@hp.com;releaseteam@hp.com"
			oMessage.Subject = "Email Test - Please Ignore" 
			'oMessage.HTMLBody = strBody
			oMessage.HTMLBody = "Test" 
			oMessage.Send 
			Set oMessage = Nothing 	

%>
</BODY>
</HTML>

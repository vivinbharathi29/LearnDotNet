<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = txtList.value;
	window.parent.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
Sending values to previous window...
<INPUT type="hidden" id=txtList name=txtList value="<%=replace(request("lstObservations")," ","")%>">
</BODY>
</HTML>

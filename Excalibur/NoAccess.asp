<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtLevel.value=="1")
		{
		window.parent.document.write ("<center><font size=2 face=verdana><BR><BR><BR>You do not have access to the requested page</font></center>");
		/*if (window.parent.frames("LowerWindow").cmdOK !=null)
			window.parent.frames("LowerWindow").cmdOK.disabled=true;
		if (window.parent.frames("LowerWindow").cmdNext !=null)
			window.parent.frames("LowerWindow").cmdNext.disabled=true;
		*/}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<font size=2 face=verdana>You do not have access to the requested page.<BR><BR>
<a href="javascript: window.history.back(2);">Back</a>
</font>
<INPUT type="hidden" id=txtLevel name=txtLevel value="<%=request("Level")%>">
</BODY>
</HTML>

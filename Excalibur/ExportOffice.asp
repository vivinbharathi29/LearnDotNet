<%@ Language=VBScript %>

<%
	if request("txtFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("txtFormat")= 2 then
		Response.ContentType = "application/msword"
	end if
		

%>

<!-- #include file = "includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
}

//-->
</SCRIPT>
</HEAD>
<BODY ID=OutputArea LANGUAGE=javascript onload="return window_onload()">
<%
response.write request("txtBody")
%>

</BODY>
</HTML>

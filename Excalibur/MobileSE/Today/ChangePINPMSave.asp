<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel(true);
    } else {
        window.returnValue = "1";
        window.close();
    }
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	Response.Write "Impersonate: " & request("cboEmployee")
	Response.Write "My ID: " & request("txtEmployeeID")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	cn.execute "usp_TodayPage_UpdatePINPMImpersonate " & clng(request("txtEmployeeID")) & "," & clng(request("cboEmployee"))

	cn.Close
	set cn = nothing

%>

</BODY>
</HTML>

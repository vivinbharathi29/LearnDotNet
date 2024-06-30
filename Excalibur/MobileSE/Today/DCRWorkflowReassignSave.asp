<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (IsFromPulsarPlus()) {
            window.parent.parent.parent.popupCallBack(1);
            ClosePulsarPlusPopup();
        }
        else {
            window.returnValue = "1";
            window.close();
        }
    }
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	cn.execute "usp_UpdateDCRWorkflowReassign " & clng(request("txtHistoryID")) & "," & clng(request("cboEmployee"))
	
	cn.Close
	set cn = nothing

%>

</BODY>
</HTML>

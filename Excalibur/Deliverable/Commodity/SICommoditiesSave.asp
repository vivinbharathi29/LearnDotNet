<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        var pulsarplusDivId = document.getElementById("pulsarplusDivId");
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value != "0") {
                if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
                    // For Closing current popup if Called from pulsarplus
                    parent.window.parent.closeExternalPopup();
                }
                else {
                    window.returnValue = txtSuccess.value;
                    window.parent.close();
                }
            }
            else
                document.write("<BR><font size=2 face=verdana>Unable to save status.</font>");
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to save status.</font>");
    }

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim UpdateArray
	dim strUpdate
	dim ValueArray
	dim cn
	dim RowsUpdated
	dim blnFailed
	dim strSuccess
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	UpdateArray = split(request("txtUpdates"),",")
	blnFailed = false
	cn.BeginTrans
	for each strUpdate in UpdateArray
		if trim(strUpdate) <> "" then
			ValueArray = split(strUpdate,"_")
			if ubound(ValueArray)=2 then
				if trim(ValueArray(2)) <> "0" then
					Response.write "spUpdateSICommodityCounts " & ValueArray(0) & "," & ValueArray(1) & "," & ValueArray(2)
					cn.Execute	"spUpdateSICommodityCounts " & ValueArray(0) & "," & ValueArray(1) & "," & eval(ValueArray(2)),RowsUpdated
					if RowsUpdated <> 1 then
						blnFailed = true
						exit for
					end if
					Response.write "<BR>"
				end if
			end if
		end if	
	next
	
	if blnFailed then
		cn.RollbackTrans
		Response.Write "Failed"
		strSuccess = "0"
	else
		cn.CommitTrans
		Response.Write "Success"
		strSuccess = "1"
	end if

	
	
%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</BODY>
</HTML>

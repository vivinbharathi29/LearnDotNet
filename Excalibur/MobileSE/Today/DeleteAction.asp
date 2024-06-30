<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = txtValidID.value;
	if (txtValidID.value == "1")
		window.close();
}


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%
	dim strID
	strID = request("xxI1Iu4uT9Tg6gR2R")
	if strID <> "" then
		response.write "<BR><font face=verdana size=2>&nbsp;&nbsp;&nbsp;Deleting " & strID & ". Please wait..."

		dim cn
		dim strConnect
		dim rowseffected
		strConnect = Session("PDPIMS_ConnectionString") 
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = strConnect
		cn.CommandTimeout = 60
		cn.IsolationLevel=256
		cn.Open
		
		cn.BeginTrans
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spRemoveAction"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = strID
		cm.Parameters.Append p

		cm.Execute rowseffected
		Set cm=nothing

	
	
	'	cn.Execute "spRemoveAction " & strID,rowseffected
	
		if cn.Errors.count or rowseffected > 10 then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
		
		cn.Close
		set cn = nothing
		response.write "<INPUT type=""hidden"" id=txtValidID name=txtValidId value=""1"">"
	else
		response.write "<BR><font face=verdana size=2>&nbsp;&nbsp;&nbsp;Invalid ID"
		response.write "<INPUT type=""hidden"" id=txtValidID name=txtValidId value=""0"">"	
	end if



%>

</BODY>
</HTML>

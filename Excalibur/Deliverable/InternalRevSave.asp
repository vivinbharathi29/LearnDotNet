<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value != "0")
			{
			window.returnValue = txtSuccess.value;
			window.parent.close();
			}
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update internal rev.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update internal rev.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim strSuccess
	dim cn
	dim cm
	dim RowsUpdated
	
	strSuccess = request("txtRev")
	

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	cn.BeginTrans
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spUpdateDeliverableInternalRev"
	
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(Request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Rev", 3, &H0001)
	if Request("txtRev") = "" then
		p.Value = null
	else
		p.Value = clng(Request("txtRev"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Team", 3, &H0001)
	if Request("txtTeam") = "" then
		p.Value = 0
	else
		p.Value = clng(Request("txtTeam"))
	end if
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@PreinstallPrepStatus", 3, &H0001)
	p.Value = clng(Request("txtPrepStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OSCode", 200, &H0001,5)
	p.Value = ucase(left(request("txtOSCode"),5))
	cm.Parameters.Append p

	cm.Execute RowsUpdated
	
	if RowsUpdated <> 1 then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.CommitTrans
	end if
		
	cn.Close
	set cn = nothing
		


%>


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">


</BODY>
</HTML>

<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value != "")
		{
		window.parent.returnValue = txtSuccess.value;
		window.parent.close();	
		}
		
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<font face=verdana size=2>Saving.  Please wait...</font>

<%
	dim cn
	dim cm
	dim strSuccess
	dim rowschanged
	
	set cn = server.CreateObject("ADODB.Connection")
	set cm = server.CreateObject("ADODB.Command")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '
	cn.Open

	cm.ActiveConnection = cn


	ErrorsFound = false
	
	cm.CommandText = "spUpdateDelRootHFCN"
	cm.CommandType =  &H0004
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("txtRootID")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Name", 200, &H0001, 120)
		p.value = left(Request("txtName"),120)
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@CategoryID", 3, &H0001)
		p.Value = request("cboCategory")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@DevManager", 3, &H0001)
		p.Value = request("cboPM")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Description", 200, &H0001, 2000)
		p.value = left(Request("txtDescription"),2000)
		cm.Parameters.Append p
	
		cm.Execute rowschanged
		
		
	
		if rowschanged <> 1 then
			Response.Write "<BR>Could not save changes."
			strSuccess = ""
		else
			strSuccess = request("txtRootID")
		end if

	set cm = nothing
	set cn=nothing

%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

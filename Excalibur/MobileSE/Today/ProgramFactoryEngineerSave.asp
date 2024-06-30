<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (txtSuccess.value != "-1") {
        var iframeName = parent.window.name;
        if (iframeName != '') {
            parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
        } else {
            window.returnValue = txtSuccess.value;
            window.close();
        }
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

Saving Program.&nbsp; Please Wait...

<%

	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateProductVersionFactoryEngineer"
	cm.CommandType =  &H0004

	if request("txtID") <> "" then
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value = clng(request("txtID"))
		cm.Parameters.Append p
	end if

	set p =  cm.CreateParameter("@SCFactoryEngineerID", 3, &H0001)
	p.value = clng(request("cboEngineer"))
	cm.Parameters.Append p

	cm.Execute RowsEffected
	Set cm=nothing

	if cn.Errors.Count > 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""-1"">"
		Response.Write "<font size=2 face=verdana><b>Unable to update this product.</b></font>"
		cn.RollbackTrans
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""" & clng(request("cboEngineer")) & """>"
		cn.CommitTrans
	end if

	set cn=nothing
%>

</BODY>
</HTML>

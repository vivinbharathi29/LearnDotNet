<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdClose_onclick() {
	window.returnValue = txtSuccess.value;
	window.opener='X';
	window.open('','_parent','')
	window.close();	
}

//-->
</SCRIPT>
</HEAD>
<BODY  bgcolor=ivory>
<%
	
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	dim strName
	dim strID
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open


	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	cm.CommandText = "spApproveNewODMUser"
	cm.CommandType =  &H0004

	set p =  cm.CreateParameter("@ID", 3, &H0001)
	p.value = clng(request("ID"))
	cm.Parameters.Append p

	cn.BeginTrans
	cm.Execute RowsEffected
	set cm = nothing
	
	if cn.Errors.Count > 1 or Rowseffected <> 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		Response.Write "<font size=2 face=verdana><b>&nbsp;&nbsp;&nbsp;Unable to approve this employee.</b></font>"
		cn.RollbackTrans
	else
		Response.Write "<font size=2 face=verdana><BR>&nbsp;&nbsp;&nbsp;Account Approved.<BR><BR></font>"
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
		cn.CommitTrans
	end if
	
		
	
	
%>

<Table width=100% cellpadding=0 cellspacing=0>
<tr><td width=20>&nbsp;</td>
<td>
<INPUT type="button" value="Close" id=cmdClose name=cmdClose LANGUAGE=javascript onclick="return cmdClose_onclick()">
</td>
</BODY>
</HTML>

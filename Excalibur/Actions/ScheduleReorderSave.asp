<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
			window.returnValue = 1;
			window.parent.close();
			}
//		else
//			document.write ("<BR><font size=2 face=verdana>Unable to update order.</font>");
		}
//	else
//		document.write ("<BR><font size=2 face=verdana>Unable to update order.</font>");

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
	dim cn
	dim i
	dim ItemArray
	dim blnFailed
	dim RowsUpdated

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	blnFailed = false
	cn.BeginTrans	
	ItemArray = split(request("txtNewOrder"),",")
		if trim(Itemarray(i)) <> "" then
			if RowsUpdated <> 1 then
				exit for
		end if
	next
	if blnFailed then
		cn.RollbackTrans
		strSuccess = ""
	else
		cn.CommitTrans
		strSuccess = "1"
	end if
	
	cn.Close
	set cn=nothing

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
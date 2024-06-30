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
<script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	    if (txtSuccess.value == "1") {
	        window.returnValue = "1";
	        if (IsFromPulsarPlus()) {
	            ClosePulsarPlusPopup();
	            window.parent.parent.parent.EOLMultiUpdateReloadCallback(1);
	        }
	        else {
	            window.parent.close();
	        }
	    }
	    else
	        document.write("<BR><font size=2 face=verdana>Unable to update deliverable Availability Information.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update deliverable Availability Information.</font>");
}


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<BR>&nbsp;&nbsp;&nbsp;<font size=2 face=verdana>Updating...</font>
<%
	dim strSuccess
	strSuccess = "1"

	dim cn
	dim cm
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans
		
		
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "spUpdateDeliverableEOL4Developer"	

	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = clng(request("ID"))
	cm.Parameters.Append p

	cm.Execute rowschanged

	set cm=nothing

		
	if  cn.Errors.count > 0 then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.CommitTrans
	end if
		
	cn.Close
	set cn = nothing
	
	
	
	
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=Request("app")%>">
</BODY>
</HTML>


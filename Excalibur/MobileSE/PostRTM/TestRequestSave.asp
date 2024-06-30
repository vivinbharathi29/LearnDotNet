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
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update the record because strSuccess 0.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update the record because strSuccess is not defined.</font>");
}

//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	dim cn
	dim cm
	dim rs
	dim strSuccess

	strSuccess = "1"
	
	

	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	cn.BeginTrans
	set cm = server.CreateObject("ADODB.Command")
	'&H0004 means stored procedure
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "spUpdatePostRTMDeliverableVersion"	

	Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
	p.Value = request("VersionID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Status", 3,  &H0001)
	p.Value = request("cboVersionStatus")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TargetDate", 135,  &H0001)
	if trim(request("txtTargetDate")) = "" then
		p.Value = null
	else
		p.Value = request("txtTargetDate")
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Comments", 200,  &H0001,80)
	if trim(request("txtComments")) = "" then
		p.Value = null
	else
		p.Value = left(request("txtComments"),80)
	end if
	cm.Parameters.Append p

	cm.Execute rowschanged
	if rowschanged <> 1 then
		strSuccess = "0"
	end if	
	Set cm=nothing
	Set p=nothing

	if cn.Errors.count > 0 then
		strSuccess = "0"
	end if	

	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if		
	
	cn.Close
	set rs = nothing
	set cn = nothing
	
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT id=txtTest name=txtTest value="<%=request("VersionID")%>">
</BODY>
</HTML>

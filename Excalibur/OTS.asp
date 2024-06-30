<%@ Language=VBScript %>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = txtSummary.value;
	window.close();

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" bgcolor=Beige>

<%
	Response.Write "<BR>&nbsp;&nbsp;Finding Observation in OTS.  Please wait..."
	'Create Database Connection
	on error resume next
	dim rs
	dim cn
	dim cm
	dim p
	dim strOutput
	dim OTSID
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	OTSID = request("ID")
	OTSID = Right("0000000" & request("ID"),7)

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetOTSSummary"

	Set p = cm.CreateParameter("@ID", 200, &H0001,7)
	p.Value = left(OTSID,7)
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	'rs.Open "spGetOTSSummary '" & OTSID & "'",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strOutput = ""
	else
		strOutput = OTSID & " - " & rs("Summary")
	end if
	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing


%>
<INPUT style="Display:" type="text" id=txtSummary name=txtSummary value="<%=replace(strOutput,"""","&quot;")%>">
</BODY>
</HTML>

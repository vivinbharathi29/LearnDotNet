<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.parent.returnValue = txtID.value;
	window.parent.close();

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" bgcolor=Beige>

<%
	Response.Write "<BR>&nbsp;&nbsp;Add new Vendor.  Please wait..."
	'Create Database Connection
	on error resume next
	dim cm
	dim cn
	dim p
	dim strOutput
	dim VendorID
	dim cnString
	dim rowschanged
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cnString =  Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.ConnectionString = cnString
	cn.Open

	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spAddVendor"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@Name", 200, &H0001, 50)
	p.value = left(request("Name"),50)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@NewID", 3,  &H0002)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SMTID", 3,  &H0002)
	cm.Parameters.Append p
	
	cm.Execute rowschanged

	strOutput = cm("@NewID")
	
	set cm = nothing
	cn.Close
	set cn = nothing


%>
<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=strOutput%>">
</BODY>
</HTML>

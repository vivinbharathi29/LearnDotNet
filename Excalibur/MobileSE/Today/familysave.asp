<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../includes/client/jquery.min.js"></script>
<script type="text/javascript" src="../../includes/client/json2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array();
	if (txtSuccess.value!="0")
	{
	    //window.alert(txtName.value);
	    OutArray[0] = txtID.value;
	    OutArray[1] = txtName.value;

	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        //save array value and return to parent page: ---
	        parent.window.parent.modalDialog.passArgument(JSON.stringify(OutArray), 'family_query_array');
	        parent.window.parent.cmdAddFamilyResults();
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.returnValue = OutArray;
	        window.close();
	    }
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

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
	cm.CommandText = "spAddNewProductFamily"
	cm.CommandType =  &H0004


	Set p = cm.CreateParameter("@Name", 200, &H0001, 64)
	strName = left(request("txtName"),64)
	p.Value = trim(strName)
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ID", 3, &H0002)
	cm.Parameters.Append p


	cm.Execute RowsEffected
	
	if cn.Errors.Count > 1 or Rowseffected <> 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		Response.Write "<font size=2 face=verdana><b>Unable to save this Product Family.</b></font>"
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""" & cm("@ID") & """>"
	end if
	
	strID = cm("@ID")

	set cm = nothing
	set cn = nothing
%>

Saving Product Family.&nbsp; Please Wait...
<INPUT type="hidden" id=txtName name=txtName value="<%=strName%>">
<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
</BODY>
</HTML>

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
	if (txtSuccess.value != "0") {
	    OutArray[0] = txtCoreTeamID.value;
	    OutArray[1] = txtName.value;

	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        //save array value and return to parent page: ---
	        parent.window.parent.modalDialog.passArgument(JSON.stringify(OutArray), 'role_query_array');
	        parent.window.parent.EditComponentResults();
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.returnValue = OutArray;
	        window.close();
	    }

	} else {
	    document.write("Unable to update this core team.");
	}
 }

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Updating Core Team.&nbsp; Please Wait...<br></font>

<%
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	dim strName
	dim strCoreTeamID
	dim strID
	dim strSuccess
	
	strSuccess = "0"
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	set cm = server.CreateObject("ADODB.command")

    'Modified By: 02/22/2016 JMalichi - Task 16730: Update all products assigned when generic component core team changes
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateOTSComponentCoreTeam"
	cm.CommandType =  &H0004

	set p =  cm.CreateParameter("@OTSComponentID", 3, &H0001)
	p.value = clng(request("txtID"))
	Response.write "<BR>" & clng(request("txtID"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@CoreTeamID", 3, &H0001)
	p.value = clng(request("cboCoreTeam"))
	Response.write "<BR>" & clng(request("cboCoreTeam"))
	cm.Parameters.Append p

	cm.Execute RowsEffected
	
	Response.write "spUpdateOTSComponentCoreTeam " & clng(request("txtID")) & "," & clng(request("cboCoreTeam"))
	if cn.Errors.Count > 1  then
		Response.Write "<font size=2 face=verdana><b>Unable to save this core team.</b></font>"
	'	cn.RollbackTrans
		strSuccess = "0"
        
		
		strName=""
	else
		strSuccess = "1"
        strCoreTeamID = clng(request("cboCoreTeam"))
        strName = request("txtName")
	'	cn.CommitTrans
	end if
	

	set rs = nothing
	set cm = nothing
	cn.Close
	set cn = nothing
		
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="text" id=txtCoreTeamID name=txtCoreTeamID value="<%=strCoreTeamID%>">
<INPUT type="text" id=txtName name=txtName value="<%=strName%>">

</BODY>
</HTML>

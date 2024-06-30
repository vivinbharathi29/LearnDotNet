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
	    OutArray[0] = txtOwnerID.value;
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
	}
	else {
	    document.write("Unable to update this owner.");
	}

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Updating Component Ownership.&nbsp; Please Wait...<br></font>

<%
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	dim strName
	dim strOwnerID
	dim strID
	dim strSuccess
	
	strSuccess = "0"
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open


	set rs = server.CreateObject("ADODB.recordset")
	set cm = server.CreateObject("ADODB.command")

	rs.Open "spGetEmployeeByID " & clng(request("cboEmployee")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strOwnerID = "0"
		strName = ""
	else
		strOwnerID = rs("ID")
		strName = longname(rs("Name"))
	end if
	rs.Close

	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateOTSComponentOwnership"
	cm.CommandType =  &H0004

	set p =  cm.CreateParameter("@ID", 3, &H0001)
	p.value = clng(request("txtID"))
	Response.write "<BR>" & clng(request("txtID"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@RoleID", 3, &H0001)
	p.value = clng(request("txtRoleID"))
	Response.write "<BR>" & clng(request("txtRoleID"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@OwnerID", 3, &H0001)
	p.value = clng(request("cboEmployee"))
	Response.write "<BR>" & clng(request("cboEmployee"))
	cm.Parameters.Append p

	cm.Execute RowsEffected
'	cn.Execute "spUpdateOTSComponentOwnerShip " & clng(request("txtID")) & "," & clng(request("txtRoleID")) & "," & clng(request("cboEmployee"))
	
	Response.write "spUpdateOTSComponentOwnerShip " & clng(request("txtID")) & "," & clng(request("txtRoleID")) & "," & clng(request("cboEmployee"))
	if cn.Errors.Count > 1  then
		Response.Write "<font size=2 face=verdana><b>Unable to save this employee.</b></font>"
	'	cn.RollbackTrans
		strSuccess = "0"
		strOwnerID = ""
		strName=""
	else
		strSuccess = "1"
	'	cn.CommitTrans
	end if
	

	set rs = nothing
	set cm = nothing
	cn.Close
	set cn = nothing
	
	
	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
		
	end function
	
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="text" id=txtOwnerID name=txtOwnerID value="<%=strOwnerID%>">
<INPUT type="text" id=txtName name=txtName value="<%=strName%>">

</BODY>
</HTML>

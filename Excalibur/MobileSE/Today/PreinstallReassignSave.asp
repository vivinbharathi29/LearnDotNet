<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value != "")
		{
		window.parent.returnValue = txtSuccess.value;
		window.parent.close();	
		}
		
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<font face=verdana size=2>Saving.  Please wait...</font>

<%
	dim cn
	dim cm
	dim strSuccess
	dim rowschanged
	
	set cn = server.CreateObject("ADODB.Connection")
	set cm = server.CreateObject("ADODB.Command")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open


	if instr(request("txtID"),":")> 0 then
		'Multi-Reassign selected
		dim PIArray
		PIArray = split(request("txtID"),",")
		
		cn.BeginTrans
		
		for i = 0 to ubound(PIArray)
			if instr(PIarray(i),":") = 0 then
				Response.Write "<BR>InvalidID<BR>"
				Response.write "<BR>" & request("PIAlerts") & "</BR>"
				strSuccess = "0"
				exit for
			else

				strProductID = trim(left(PIArray(i),instr(PIarray(i),":")-1))		
				strVersionID = trim(mid(PIArray(i),instr(PIarray(i),":")+ 1))

				set cm = server.CreateObject("ADODB.Command")
			
				cm.CommandText = "spUpdatePreinstallOwner2"
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
				
	
				Set p = cm.CreateParameter("@ProdID", 3, &H0001)
				p.Value = clng(strProductID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@DelID", 3, &H0001)
				p.Value = clng(strVersionID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OwnerID", 3, &H0001)
				p.Value = request("cboOwner")
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@UpdateGroup", 16, &H0001)
				p.Value = request("txtUpdateGroup")
				cm.Parameters.Append p
	
				cm.Execute rowschanged
				Set cm = Nothing
		
				if cn.Errors.count > 0 then
					strSuccess = "0"
					exit for
				else
					strSuccess = "1"
				end if
			end if
		next 

		if strSuccess="0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
	
	else
		'Single-Reassign selected
		cn.BeginTrans
		
		cm.CommandText = "spUpdatePreinstallOwner"
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("txtID")
		Response.Write "ID:" & request("txtID") & "<BR>"
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OwnerID", 3, &H0001)
		p.Value = request("cboOwner")
		Response.Write "Owner: " & request("cboOwner") & "<BR>"
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@UpdateGroup", 16, &H0001)
		p.Value = request("txtUpdateGroup")
		Response.Write "Group: " & request("txtUpdateGroup") & "<BR>"
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PartNumber", 200, &H0001,50)
		p.Value = left(request("txtPart"),50)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Urgent", 11, &H0001)
		if request("chkUrgent") = "on" then
			p.Value = 1
		else
			p.value = 0
		end if
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Notes", 200, &H0001,80)
		p.Value = left(request("txtNotes"),80)
		cm.Parameters.Append p
	
		cm.Execute rowschanged
				
		set cm = nothing
		
		if cn.Errors.count > 0 then
			Response.Write "<BR>Could not save changes."
			cn.RollbackTrans
			strSuccess = ""
		else
			strSuccess = request("txtName")
			cn.CommitTrans
		end if


	end if

	set cn=nothing
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

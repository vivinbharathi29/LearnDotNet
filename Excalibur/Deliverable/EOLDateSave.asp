<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
		    window.returnValue = txtRemove.value;
		    //close window
		    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
		        parent.window.parent.closeExternalPopup();
		        parent.window.parent.reloadFromPopUp(pulsarplusDivId);
		    }
			else if (IsFromPulsarPlus()) {
			    ClosePulsarPlusPopup();
			    window.parent.parent.parent.ComponentEndOfLifeDateExpiredReloadCallback(txtRemove.value);
			}
			else {
			    if (parent.window.parent) {
			        alert("Successfully update.\nPlease reload the report to see the change.");
			        parent.window.parent.ClosePopUp();
			    } else {
			        window.parent.close();
			    }
			}
		}
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update deliverable Availablity Information.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update deliverable Availablity Information.</font>");
}


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<%
	dim strRemove 
	dim strSuccess
	if request("chkEOL") = "" then
		Response.Write "<BR>Remove It - Active Not Checked"
		strRemove = "1"
	elseif isdate(request("txtEOLDate")) then
		if datediff("d",request("txtEOLDate"),Now())>0 and trim(request("txtTypeID"))<>"2" then
			Response.Write "<BR>Leave It - EOL Date is Expired"
			strRemove = "0"
		elseif datediff("d",request("txtEOLDate"),Now())>-90 and trim(request("txtTypeID"))="2" then
			Response.Write "<BR>Leave It - EOL Date is Expired"
			strRemove = "0"
		else
			Response.Write "<BR>Remove It - EOL Date has not expired yet."
			strRemove = "1"
		end if
	else
		Response.Write "<BR>Remove It - No EOL Date Specified and Not EOL."
		strRemove = "1"
	end if
	
	
	strSuccess = "1"

	dim cn
	dim cm
	dim blnErrors
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans
	blnErrors = false
		
		
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
	if trim(request("txtTypeID"))="2" then
		cm.CommandText = "spUpdateDeliverableServiceEOL"	
	else
		cm.CommandText = "spUpdateDeliverableEOL"	
	end if
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = clng(request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@EOLDate", 135,  &H0001)
	if request("txtEOLDate") = "" then
		p.Value = null
	else
		p.Value = cdate(request("txtEOLDate"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Active", 11,  &H0001)
	if trim(request("chkEOL")) = "" then
		p.Value = false
	else
		p.Value = true
	end if
	cm.Parameters.Append p

	cm.Execute rowschanged

	set cm=nothing

		
	if  rowschanged <> "-1" then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.CommitTrans
	end if
		
	cn.Close
	set cn = nothing
	
	
	
	
%>
<INPUT type="hidden" id=txtRemove name=txtRemove value="<%=strRemove%>">
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

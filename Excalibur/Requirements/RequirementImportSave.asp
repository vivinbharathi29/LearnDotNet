<%@ Language=VBScript %>
 <% Server.ScriptTimeout = 6000 %>

<!-- #include file = "../includes/noaccess.inc" -->
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
		    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
		        // For Reload PulsarPlusPmView Tab
		        parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

		        // For Closing current popup
		        parent.window.parent.closeExternalPopup();
		    }
		    else {
		        if (parent.window.parent.document.getElementById('modal_dialog')) {
		            //save value and return to parent page: ---
		            parent.window.parent.modalDialog.cancel(true);
		        } else {
		            window.returnValue = 1;
		            window.parent.close();
		        }
		    }
		}
//		else
//			document.write ("<BR><font size=2 face=verdana>Unable to import the requirement list.</font>");
	}
//	else
//		document.write ("<BR><font size=2 face=verdana>Unable to import the requirement list.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();">

<%
	dim i
	dim ReqArray
	dim cn
	dim p
	dim cm
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.CommandTimeout = 3000
	cn.Open
		
	cn.BeginTrans

	ReqArray = split(request("chkSelected"),",")
	for i = lbound(ReqArray) to ubound(ReqArray)
		if isnumeric(ReqArray(i)) then
			Response.Write "Copy " & ReqArray(i)  & " to Product " & request("txtID") & "<BR>"
			
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
		
			cm.CommandText = "spCopyRequirement"	
			cm.CommandTimeout = 3000

			Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
			p.Value = request("txtID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@RequirementID", 3,  &H0001)
			p.Value = ReqArray(i)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@DelCopy", 11,  &H0001)
			p.Value = true
			cm.Parameters.Append p
					
			cm.Execute rowschanged

			if cn.Errors.count > 0 then
				FoundErrors = true
				exit for
			end if
		
			set cm = nothing
			
		end if
	next
	
	cn.CommitTrans
	
	set p = nothing
	set cm = nothing
	set cn = nothing
	
	if FoundErrors then
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
	else
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
	end if
	
%>
     <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</BODY>
</HTML>

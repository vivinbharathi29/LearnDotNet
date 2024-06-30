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
	        window.parent.returnValue="1";
	        var pulsarplusDivId = '<%=Request("pulsarplusDivId")%>';
	        var profileId = '<%=Request("txtProfileID")%>';
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	            parent.window.parent.reloadProfileShare(profileId);
	        }
	        else {
	            window.parent.close();
	        }
	    }
	//	else
	//		document.write ("Unable to save profile sharing information.  An unexpected error occurred.");	
		}
//	else
//		{
//		document.write ("Unable to save profile sharing information.  An unexpected error occurred.");
//		}

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
	if trim(request("txtProfileID")) = "" or trim(request("txtAction")) = "" or trim(request("txtEmployeeID") ) = "" then
		Response.Write "Not enough information provide to display this page."	
	else

		set cn = server.CreateObject("ADODB.Connection")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open

		if trim(request("txtAction")) = "2" and trim(request("AddType")) <> "2" then
			cn.execute "spRemoveSharedProfile2 " & clng(request("txtProfileID")) & "," & clng(request("txtEmployeeID")) ,RowCount
		elseif trim(request("txtAction")) = "2" then
			cn.execute "spRemoveSharedProfileGroup " & clng(request("txtProfileID")) & "," & clng(request("txtEmployeeID")) ,RowCount
		elseif trim(request("AddType")) = "2" then
			cn.Execute "spUpdateSharedProfileGroup " & clng(request("txtProfileID")) & "," & clng(request("txtEmployeeID")) & "," & clng(request("optEditPermission")) & "," & clng(request("optDeletePermission")),RowCount
		else
			cn.Execute "spUpdateSharedProfile " & clng(request("txtProfileID")) & "," & clng(request("txtEmployeeID")) & "," & clng(request("optEditPermission")) & "," & clng(request("optDeletePermission")),RowCount
		end if
		
		if RowCount <> 1 then
			strSuccess = "0"
		else
			strSuccess = "1"
		end if		

		cn.Close
		set cn = nothing
	end if
%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

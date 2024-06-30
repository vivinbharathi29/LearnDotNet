<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp(pulsarplusDivId);                    
                }
                else {
                    if (window.parent.frames["UpperWindow"]) {
                        parent.window.parent.modalDialog.cancel(true);
                    } else {
                        window.parent.close();
                    }
                    /*window.returnValue = 1;
                    window.parent.close();*/
                }
            }
            else
                document.write("<BR><font size=2 face=verdana>Unable to import the image list.</font>");
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to import the image list.</font>");
    }

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">

<%=request("txtID")%><BR>

<%
    response.Flush
	dim strSuccess
	dim ImportArray
	dim i
	dim cn
	dim rs
	dim cm
	dim NewID
			
	strSuccess = "1"
	
	if trim(request("chkSelected")) <> "" then
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
	
'		set rs = server.CreateObject("ADODB.recordset")
	
		ImportArray = split(request("chkSelected"),",")
	
		cn.BeginTrans
		for i = lbound(ImportArray) to ubound(ImportArray)
	        response.write "Old:" & ImportArray(i) & "<br>"
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
		
			cm.CommandText = "spImportImageDefinitionFusion"	

			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = ImportArray(i)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
			p.Value = clng(request("txtID"))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@NewID", 3,  &H0002)
			cm.Parameters.Append p

			cm.Execute rowschanged

			NewID = cm("@NewID")
			set cm=nothing

			if cn.Errors.count > 0 then
				strSuccess = "0"
				exit for
			else
			
	            response.write "New:" & NewID & "<br><br>"
			
				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
		
				cm.CommandText = "spImportImages"	

				Set p = cm.CreateParameter("@OldID", 3,  &H0001)
				p.Value = ImportArray(i)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@NewID", 3,  &H0001)
				p.Value = NewID
				cm.Parameters.Append p

			
				cm.Execute rowschanged

				set cm = nothing

				if cn.Errors.count > 0 then
					strSuccess = "0"
					exit for
				end if
			
			
			end if
		
	
		next
	
		if strSuccess = "0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
	
'		set rs = nothing
		cn.Close
		set cn=nothing
	end if
	

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

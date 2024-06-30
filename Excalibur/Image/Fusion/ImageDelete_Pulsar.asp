<%@ Language=VBScript %>

<% Option Explicit%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (txtSuccess.value == "1") {
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.closeExternalPopup();
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                
            }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
                /*window.returnValue = 1;
                window.parent.close();*/
            }
        }
        else
            document.write("Unable to delete Image Definition.");
		
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">


<%

	dim cn
	dim cm
	dim p
	dim strSuccess
	dim rowschanged
	dim blnFailed
	
	if request("Auth") = "DeLeTeOk" and request("DelImageID") <> "" and request("txtDelUserID") <> "" then
		
		blnFailed = false

		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "usp_Image_DisableImageDefinition"		

		Set p = cm.CreateParameter("@p_intImageDefID", 3, &H0001)
		p.Value = request("DelImageID")
		cm.Parameters.Append p
	
        Set p = cm.CreateParameter("@p_intDCRID", 3,  &H0001)
		if request("DelDCRID") <> "" and isnumeric(request("DelDCRID")) then
			p.Value = clng(request("DelDCRID"))
		else
			p.Value = null
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@p_intUserID", 3,  &H0001)
		p.Value = clng(request("txtDelUserID"))
		cm.Parameters.Append p

        cm.Execute rowschanged
		Set cm=nothing

		if rowschanged <> 1 then
			blnFailed = true
		end if		

		if blnFailed then			
			strSuccess = "0"	
		else			
			strSuccess = "1"
		end if
	else
		strSuccess = "0"
	end if

	cn.Close
	set cn = nothing

	%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>
<script src="../Scripts/Pulsar2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (txtSuccess.value != "0") {
        if (isFromPulsar2()) {
            closePulsar2Popup(true);
        }
        else if (IsFromPulsarPlus()) {
            window.parent.parent.parent.ApplicationErrorCallback(txtSuccess.value);
            ClosePulsarPlusPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel(true);
            } else {
                window.returnValue = "1";
                window.parent.close();
            }
        }

    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Saving Cause.&nbsp; Please Wait...<br></font>

<%
	strSuccess = "0"
	if trim(request("txtID")) = "" then
		Response.Write "Unable to process the request because not enough information was supplied."
	else
		dim cn
		dim cm
		dim rowseffected
	
		
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
		cn.BeginTrans	
		
		if trim(request("txtCopyTo")) = "" then 'Update this and all duplicate errors

		    set cm = server.CreateObject("ADODB.command")
    		
		    cm.ActiveConnection = cn
		    cm.CommandText = "spUpdateAppError"
		    cm.CommandType =  &H0004

		    set p =  cm.CreateParameter("@ID", 3, &H0001)
		    p.value = clng(request("txtID"))
		    cm.Parameters.Append p
    	
		    Set p = cm.CreateParameter("@Cause", 200, &H0001, 2000)
		    p.Value = left(request("txtCause"),2000)
		    cm.Parameters.Append p
    	
		    cm.Execute RowsEffected
		    if cn.Errors.count >0 then
			    strSuccess = 0
			    cn.RollbackTrans
		    else	
			    strSuccess=1
			    cn.CommitTrans
		    end if
    	
		    set cm = nothing
        else 'Updated selected errors
            dim strUpdate
            dim UpdateArray
            dim strError
            
            strUpdate = request("txtID") & "," & request("txtCopyTo")
            UpdateArray = split(strUpdate,",")
            for each strError in UpdateArray
                if trim(strError) <>"" then
		            set cm = server.CreateObject("ADODB.command")
            		
		            cm.ActiveConnection = cn
		            cm.CommandText = "spUpdateAppError2"
		            cm.CommandType =  &H0004

		            set p =  cm.CreateParameter("@ID", 3, &H0001)
		            p.value = clng(strError)
		            cm.Parameters.Append p
            	
		            Set p = cm.CreateParameter("@Cause", 200, &H0001, 2000)
		            p.Value = left(request("txtCause"),2000)
		            cm.Parameters.Append p
            	
		            cm.Execute RowsEffected
		            set cm = nothing
                

		            if RowsEffected <> 1 then
			            strSuccess = 0
		                exit for
		            else	
			            strSuccess=1
		            end if
                
                end if
            next

		    if strSuccess = 0 then
			    cn.RollbackTrans
		    else	
			    cn.CommitTrans
		    end if
        
        end if
		set cn = nothing
	end if
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>


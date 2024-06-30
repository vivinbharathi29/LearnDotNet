<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value != "")
	{
	    if (IsFromPulsarPlus()) {
	        window.parent.parent.parent.popupCallBack(1);
	        ClosePulsarPlusPopup();
	    }
	    else {
	        if (parent.window.parent.document.getElementById('modal_dialog')) {
	            try {
	                parent.window.parent.CloseFeatureRequestRejectMVC();
	            } catch (err) {
	                parent.window.parent.modalDialog.cancel(true);
	            }

	        } else {
	            window.parent.returnValue = txtSuccess.value;
	            window.parent.close();
	        }
	    }
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<font face=verdana size=2>Saving.  Please wait...</font>

<% response.Flush
	dim cn
	dim cm
    dim rs
	dim strSuccess
	dim rowschanged

	
	set cn = server.CreateObject("ADODB.Connection")
	set cm = server.CreateObject("ADODB.Command")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open


	    'Get User
        set rs = server.CreateObject("ADODB.recordset")
	    dim CurrentDomain
        dim CurrentUser
	    dim CurrentUserID
	    CurrentUser = lcase(Session("LoggedInUser"))

	    if instr(currentuser,"\") > 0 then
		    CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		    Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	    end if

	    Set cm.ActiveConnection = cn
	    set rs = server.CreateObject("ADODB.recordset")

	    cm.CommandType = 4
	    cm.CommandText = "spGetUserInfo"
	

	    Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	    p.Value = Currentuser
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	    p.Value = CurrentDomain
	    cm.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 

	    set cm=nothing	
	

	    CurrentUserID = 0
	    if not (rs.EOF and rs.BOF) then
		    CurrentUserID = rs("ID")
	    end if
	    rs.Close

        'save
        set cm = server.CreateObject("ADODB.Command")

		cn.BeginTrans
		
		cm.CommandText = "spPULSAR_Today_RejectFeatureRequest"
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = clng(request("txtID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@WhyRejected", 200, &H0001,80)
		p.Value = left(request("txtReason"),80)
		cm.Parameters.Append p

    	Set p = cm.CreateParameter("@RejectedBy", 3, &H0001)
		p.Value = CurrentUserID
		cm.Parameters.Append p
	
		cm.Execute rowschanged
				
		set cm = nothing
		
		if cn.Errors.count > 0 then
			Response.Write "<BR>Could not save changes."
			cn.RollbackTrans
			strSuccess = ""
		else
			strSuccess = "1"
			cn.CommitTrans
		end if


    cn.close
	set cn=nothing
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

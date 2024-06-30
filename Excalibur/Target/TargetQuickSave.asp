<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<%if request("Type") = 1 then%>
<TITLE>Target Version</TITLE>
<%else%>
<TITLE>Reject Version</TITLE>
<%end if%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined"){
		    if (txtSuccess.value=="1"){
		        //close window
		        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
		            parent.window.parent.closeExternalPopup();
		            parent.window.parent.reloadFromPopUp(pulsarplusDivId);
		        }
		        else if (IsFromPulsarPlus()) {
		            window.parent.parent.parent.popupCallBack(1);
		            ClosePulsarPlusPopup();
		        }
		        else {
		            if (window.location != window.parent.location) {
		                parent.modalDialog.passArgument(1, 'target_save_status');
		                parent.modalDialog.cancel();
		            } else {
		                window.returnValue = 1;
		                window.close();
		            }
		        }
			    //window.returnValue = 1;
			    //window.close();
		    }
	}else {
	    document.write ("<font face=verdana size=2>Unable to update this deliverable.  An unexpected error occurred.</font>");
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<BR>
<table width=100%><TR><TD align=center>
<%if request("Type") = 1 or request("TargetType") = 1 then%>
<font face=verdana size =2>Targeting Version.  Please wait...</font>
<%else%>
<font face=verdana size =2>Rejecting Version.  Please wait...</font>
<%end if%>
</td></tr></table>
<%

	dim strVersionID
	dim strProductID
	dim strType
	dim strRejected
	dim cn
	dim rs
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
	dim cm
	dim TargetArray

	if (request("ProdID") = "" or request("VersionID") = "" or request("Type") = "") and request("PMAlerts") = "" then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
'		Response.Write "<font face=verdana size=1>ProductID:" & request("ProdID") & "<BR>"
'		Response.Write "DeliverableID:" & request("VersionID") & "<BR>"
'		Response.Write "Type:" & request("Type") & "<BR>"
'		Response.Write "Multi:" & request("PMAlert") & "<BR>"
'		Response.Write "MultiType:" & request("TargetType") & "<BR></font>"

	elseif request("PMAlerts") <> "" then


		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open


	'Get User
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
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
	
	if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") 
		end if
		rs.Close
	

	
		
		cn.BeginTrans
		
		TargetArray = split(request("PMAlerts"),",")
		
		for i = 0 to ubound(TargetArray)
			if instr(targetarray(i),":") = 0 then
				Response.Write "<BR>InvalidID<BR>"
				Response.write "<BR>" & request("PMAlerts") & "</BR>"
				strSuccess = "0"
				exit for
			else
				strProductID = trim(left(targetArray(i),instr(targetarray(i),":")-1))		
				strVersionID = trim(mid(targetArray(i),instr(targetarray(i),":")+ 1))
	
				set cm = server.CreateObject("ADODB.Command")
		
				cm.ActiveConnection = cn
				cm.CommandText = "spTargetDeliverableVersionWeb"
				cm.CommandType = &H0004
						
				Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
				p.Value = clng(strProductID)
				cm.Parameters.Append p
			
				Set p = cm.CreateParameter("@DeliverableVersionID",adInteger, &H0001)
				p.Value = clng(strVersionID)
				cm.Parameters.Append p
			    
				Set p = cm.CreateParameter("@TargetValue",adBoolean, &H0001)
				p.Value = clng(request("TargetType"))
				cm.Parameters.Append p
			    
				Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
				p.Value = CurrentUserID
				cm.Parameters.Append p
			
				Set p = cm.CreateParameter("@Rejected",adBoolean, &H0001)
				if request("TargetType") = "1" then
					p.Value = 0
				else
					p.Value = 1
				end if	
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@ResetOtherTargets",adBoolean, &H0001)
				if request("TargetType") = "1" then
					p.Value = 1
				else
					p.Value = 0
				end if
				cm.Parameters.Append p
					
				cm.Execute
				Set cm = Nothing
	
				if cn.Errors.count > 0 then
					strSuccess = "0"
					cn.RollbackTrans
					exit for
				else
					strSuccess = "1"
				end if
				Response.Write "<BR>DONE<BR>"
			end if
		next
		if strSuccess = "1" then
				cn.CommitTrans
		end if
					            
		set rs = nothing
		set cn = nothing


	else
	
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open


	'Get User
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
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
	
	if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") 
		end if
		rs.Close
	

		set cm = server.CreateObject("ADODB.Command")
	
		
		cn.BeginTrans
		
					            
	    cm.ActiveConnection = cn
	    cm.CommandText = "spTargetDeliverableVersionWeb"
	    cm.CommandType = &H0004
	       
	    Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
		p.Value = clng(request("ProdID"))
	    cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@DeliverableVersionID",adInteger, &H0001)
		p.Value = clng(request("VersionID"))
	    cm.Parameters.Append p
	    
	    Set p = cm.CreateParameter("@TargetValue",adBoolean, &H0001)
		p.Value = clng(request("Type"))
	    cm.Parameters.Append p
	    
	    Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
		p.Value = CurrentUserID
	    cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Rejected",adBoolean, &H0001)
		if request("Rejected") = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@ResetOtherTargets",adBoolean, &H0001)
		if request("Type") = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p
		    
	    cm.Execute
		Set cm = Nothing
		
		if cn.Errors.count > 0 then
			strSuccess = "0"
			cn.RollbackTrans
		else
			strSuccess = "1"
			cn.CommitTrans
		end if
	
		set rs = nothing
		set cn = nothing

end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

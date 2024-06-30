<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<%if trim(request("Status")) = "6" then%>
<TITLE>Cancel File Transfer</TITLE>
<%else%>
<TITLE>Retry File Transfer</TITLE>
<%end if%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
	{
		if (txtSuccess.value=="1")
		{
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.popupCallBack(1);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        window.returnValue = 1;
		        window.close();
		    }
		}
	}
	else
	{
	document.write ("<font face=verdana size=2>Unable to update this deliverable.  An unexpected error occurred.</font>");
	}


}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<BR>
<table width=100%><TR><TD align=center>
<%if trim(request("Status")) = "6" then%>
<font face=verdana size =2>Cancel File Transfer.  Please wait...</font>
<%else%>
<font face=verdana size =2>Restart File Transfer.  Please wait...</font>
<%end if%>
</td></tr></table>
<%

	dim cn
	dim rs
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
	dim cm

	if request("Status") = "" or request("ID") = "" then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
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
	

		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.Command")
		
		cm.ActiveConnection = cn
		cm.CommandText = "spFusion_COMPONENT_ComponentCopiedToIRS"
		cm.CommandType = &H0004
						
		Set p = cm.CreateParameter("@ID",adInteger, &H0001)
		p.Value = clng(request("ID"))
		cm.Parameters.Append p
			
	    Set p = cm.CreateParameter("@Status",adInteger, &H0001)
	    p.Value = clng(request("Status"))
	    cm.Parameters.Append p
			    
		cm.Execute
		Set cm = Nothing
	
		if cn.Errors.count > 0 then
			strSuccess = "0"
			cn.RollbackTrans
		else
			strSuccess = "1"
		end if
		Response.Write "<BR>DONE<BR>"

	    if strSuccess = "1" then
    	    cn.CommitTrans
	    end if
					            
	    set rs = nothing
	    set cn = nothing
    end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

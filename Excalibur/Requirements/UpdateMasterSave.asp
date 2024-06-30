<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array;
	
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value=="1")
			{
			OutArray[0] = txtNewID.value;
			OutArray[1] = txtName.value;
			window.parent.returnValue = OutArray;
			window.parent.close();
			}
		else
			document.write ("<BR><BR>Unable to add requirement.  An unexpected error occurred.");
		
		}
	else
		{
		document.write ("<BR><BR>Unable to add requirement.  An unexpected error occurred.");
		}
}


//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim cn
	dim cm
	dim p
	dim strSuccess
	strSuccess = "0"
  
	if request("txtName") =  "" then
		FoundErrors = true
	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
	
		cn.BeginTrans

		FoundErrors = false	
	
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn	
		cm.CommandText = "spAddNewRequirement"	
	
		Set p = cm.CreateParameter("@ReqName", 200,  &H0001,255)
		p.Value = request("txtName")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@CategoryID", 3,  &H0001)
		p.Value = 127 'Default to "General"
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SpecTemplate", 201, &H0001, 2147483647)
		p.value = "describe"
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@ProcessRequirement", 11,  &H0001)
		p.Value = 0 'Default to "No"
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
		cm.Parameters.Append p

		
		cm.Execute rowschanged

		if rowschanged <> 1 then
			FoundErrors = true
		else
			strNewID = cm("@NewID")
		end if
		
		if FoundErrors then
			cn.RollbackTrans
		else
			cn.CommitTrans
			strSuccess = "1"
			on error resume next
			dim oMessage
			dim CurrentUser
			dim CurrentUserEmail
			
			
			'Get User
			dim CurrentDomain
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
				CurrentUserEmail = rs("Email") & ""
			end if			
			
			if CurrentUSerEmail = "" then
				CurrentUserEmail = "max.yu@hp.com"
			end if
						
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")		
		
			oMessage.From = CurrentUserEmail
			oMessage.To= "max.yu@hp.com"
			oMessage.Subject = "New Marketing Requirement Added" 
		
			oMessage.HTMLBody = "<font size=2 face=verdana>" & request("txtName") & " has been added to the master list of marketing requirements.</font>"
		
			oMessage.Send 
			Set oMessage = Nothing 			
		
		
		end if

				
		set cm = nothing
		set cn = nothing
		
		
		
		
		
	end if
	
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="text" id=txtNewID name=txtNewID value="<%=strNewID%>">
<INPUT type="text" id=txtName name=txtName value="<%=request("txtName")%>">
</BODY>
</HTML>

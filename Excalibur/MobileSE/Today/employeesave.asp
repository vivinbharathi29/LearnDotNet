<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        var OutArray = new Array();
	if (txtSuccess.value!="0")
		{
            //window.alert(txtName.value);
            OutArray[0] = txtID.value;
            OutArray[1] = txtName.value;
            OutArray[2] = txtPartnerID.value;
            window.returnValue = OutArray;
            window.parent.opener = 'X';
            window.parent.open('', '_parent', '')
            window.parent.location.replace("/pulsarplus/today");
            //window.parent.close();
            //window.close();
        }
    }

    //-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Saving Employee.&nbsp; Please Wait...<br></font>

<%
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	dim strName
	dim strID
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open


	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	if request("txtID") = "" then
		cm.CommandText = "spAddEmployeeWeb"
	else
		cm.CommandText = "spUpdateEmployee"
	end if
	cm.CommandType =  &H0004

	if request("txtID") <> "" then
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value = clng(request("txtID"))
		cm.Parameters.Append p
	end if
	
	Set p = cm.CreateParameter("@Name", 200, &H0001, 80)
	if trim(request("txtPartnerName")) <> "" and instr(lcase(request("txtFirstName"))," (" & lcase(request("txtPartnerName")) & ")")=0 then
		strName = left(trim(request("txtLastName")) & ", " & trim(request("txtFirstName")) & " (" & trim(request("txtPartnerName")) & ")",80)
	else
		strName = left(trim(request("txtLastName")) & ", " & trim(request("txtFirstName")),80)
	end if
	p.Value = strName
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FirstName", 200, &H0001, 30)
	p.Value = left(trim(request("txtFirstName")),30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@LastName", 200, &H0001, 30)
	p.Value = left(trim(request("txtLastName")),30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Phone", 200, &H0001, 30)
	p.Value = left(trim(request("txtPhone")),30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Email", 200, &H0001, 80)
	p.Value = left(trim(request("txtEmail")),80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@NTName", 200, &H0001, 80)
	strNTName = lcase(trim(request("txtNTName")))
	do while instr(strNTName,"\") > 0
		strNTName = mid(strNTName,instr(strNTName,"\")+1)
	loop
	p.Value = left(strNTName,80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	if isnumeric(request("cboDomain")) then
		p.value = "houhpqexcal03"
	else
		p.value = trim(request("cboDomain"))
	end if
'	if instr(lcase(request("txtNTName")),"\") > 0 then
'		p.Value =  left(left(lcase(request("txtNTName")),instr(lcase(request("txtNTName")),"\")-1),30)
'	else
'		p.Value = ""
'	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@WorkgroupID", 3, &H0001)
	p.value = clng(request("cboGroup"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@DivisionID", 3, &H0001)
	p.value = clng(request("cboDivision"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PartnerID", 3, &H0001)
	p.value = clng(trim(request("txtPartnerID")))
	cm.Parameters.Append p 

   	if request("txtID") = "" then

        Set p = cm.CreateParameter("@User", 200, &H0001, 30)
	    p.value = "Excalibur"
	    cm.Parameters.Append p 

    	set p =  cm.CreateParameter("@FTPAccessRequested", 3, &H0001)
	    if trim(request("cboVPN")) = "1" and trim(request("cboFTP")) = "1" then
	        p.value = 1
	    else
	        p.value = 0
	    end if
	    cm.Parameters.Append p

		set p =  cm.CreateParameter("@ID", 3, &H0002)
		cm.Parameters.Append p
	end if

	cn.BeginTrans
	cm.Execute RowsEffected
	if cn.Errors.Count > 1 or Rowseffected <> 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		Response.Write "<font size=2 face=verdana><b>Unable to save this employee.</b></font>"
		cn.RollbackTrans
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""" & cm("@ID") & """>"
		cn.CommitTrans
	end if
	
	if request("txtID") = "" then
		strID = cm("@ID")
	else
		strID = request("txtID")
	end if



	'Update ODM User NTName if needed
	if isnumeric(strID) and clng(request("txtPartnerID")) <> 1 and trim(lcase(request("txtNTName"))) = left(trim(lcase(request("txtEmail"))),30) then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn

		cm.CommandType = 4
		cm.CommandText = "spUpdateODmUserNTName"
	

		Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
		p.Value = clng(strID)
		cm.Parameters.Append p
        
		cm.Execute 

		set cm=nothing	

    end if	
	
	if request("txtID") = "" then
		'get Current User Info
		Dim CurrentUser	
		Dim CurrentUserID
		dim CurrentUserEmail
	
		CurrentUserID = 0
		CurrentUserEmail = ""

		'Get User
		dim CurrentDomain
		CurrentUser = lcase(Session("LoggedInUser"))
	
		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

		set cm = server.CreateObject("ADODB.Command")
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

		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 

		set cm=nothing	
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID")
			CurrentUserEmail = rs("Email") & ""
		end if
		rs.Close

		'Set the Sponser as the Manager for ODM users
		if CurrentUserID <> 0 and isnumeric(strID) and clng(request("txtPartnerID")) <> 1 then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn

			cm.CommandType = 4
			cm.CommandText = "spUpdateEmployeeODMSponser"
	

			Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
			p.Value = clng(strID)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@SponserID", 3, &H0001)
			p.Value = clng(CurrentUserID)
			cm.Parameters.Append p

			cm.Execute 

			set cm=nothing	
		end if
	
	
		'Send Email
	
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
	
		if CurrentUserEmail <> "" then
			oMessage.From = CurrentUserEmail
		else
			oMessage.From = "max.yu@hp.com"
		end if
		oMessage.To= "max.yu@hp.com"
	
		dim strUserNotes
		if trim(request("txtNotes")) <> "" then
			strUserNotes = "<BR><BR>Notes: " & server.HTMLEncode(request("txtNotes"))
		else
			strUserNotes = ""
		end if
		if trim(request("cboVPN")) = "1" then
			oMessage.Subject = "Employee Record Added - Partner Login account requested"
			oMessage.HTMLBody =  "<Body><font size=2 face=verdana>A new employee account [ " & strName & " ] has been added by " & request("txtCurrentUser") & ".<BR><BR><b><font color=red>A Partner Login account is requested.</font></b>" & strUserNotes & "<BR><BR><a href=""http://16.81.19.70/mobilese/today/employee.asp?ID=" & strID & """>View In Excalibur</a></body>"
		else
			oMessage.Subject = "Employee Record Added - No Partner Login account required"
			oMessage.HTMLBody =  "<Body><font size=2 face=verdana>A new employee account [ " & strName & " ] has been added by " & request("txtCurrentUser") & "." & strUserNotes & "<BR><BR><a href=""http://16.81.19.70/mobilese/today/employee.asp?ID=" & strID & """>View In Excalibur</a></body>"
		end if				
		oMessage.DSNOptions = cdoDSNFailure

		'oMessage.Send 
		Set oMessage = Nothing 
	end if	
	
	set cm = nothing
	set cn = nothing

%>

<INPUT type="hidden" id=txtName name=txtName value="<%=strName%>">
<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
<INPUT type="hidden" id=txtPartnerID name=txtPartnerID value="<%=trim(request("txtPartnerID"))%>">

</BODY>
</HTML>

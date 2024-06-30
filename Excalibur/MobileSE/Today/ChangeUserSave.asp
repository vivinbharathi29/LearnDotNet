<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = "1";
	window.close();
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	dim CurrentDomain
	dim CurrentUser
		dim strEmployees
		dim ActualUserID
        dim rs
        dim blnActualUserPulsarAdmin

	CurrentUser = lcase(Session("LoggedInUser"))
     
	    set rs = server.CreateObject("ADODB.recordset")

        if instr(currentuser,"\") > 0 then
		    CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		    Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	    end if

	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open

        set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
	    cm.CommandType = 4
	    cm.CommandText = "spGetEmployeeImpersonateID"
		
	    Set p = cm.CreateParameter("@NTName", 200, &H0001, 80)
	    p.Value = Currentuser
	    cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	    p.Value = CurrentDomain
	    cm.Parameters.Append p
	
	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set	rs = cm.Execute 
	
	    set cm=nothing	
		
	    if not (rs.EOF and rs.BOF) then
            ActualUserID = rs("EmployeeID")
        else
            ActualUserID = ""
        end if
        rs.close

        dim blnSupportAdmin
        rs.open "spSupportIsAdminSelect " & clng(CurrentUserID),cn
	    'Response.write("UserName: " + rs("Name"))
        if rs.eof and rs.bof then
            blnSupportAdmin = false
        else
            blnSupportAdmin = true
        end if
        rs.close

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
	    Set	rs = cm.Execute 
	
	    set cm=nothing	
		
	    if not (rs.EOF and rs.BOF) then

            if CLng(rs("ID")) <> CLng(ActualUserID) or Cint(rs("PulsarSystemAdmin")) = 1 or blnSupportAdmin then
                blnActualUserPulsarAdmin = True
            else
                blnActualUserPulsarAdmin = False
            end if
        else
            blnActualUserPulsarAdmin = false
        end if
        rs.close

    rs.open "spSupportIsAdminSelect " & clng(CurrentUserID),cn
	'Response.write("UserName: " + rs("Name"))
    if rs.eof and rs.bof then
        blnSupportAdmin = false
    else
        blnSupportAdmin = true
    end if
    rs.close

    if blnActualUserPulsarAdmin = false and blnSupportAdmin = false then
		Response.Write "You do not have access to make this change."
	else
		'Response.Write "Impersonate: " & request("cboEmployee")
		'Response.Write "My ID: " & request("txtEmployeeID")
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open

		cn.execute "spUpdateEmployeeImpersonate " & clng(request("txtEmployeeID")) & "," & clng(request("cboEmployee"))
	
		cn.Close
		set cn = nothing
	end if

%>

</BODY>
</HTML>

<%@  language="VBScript" %>
<% Option Explicit %>
<%


dim cn
dim cm
dim arrReleaseID
dim ReleaseIDs
dim item
dim p
dim rowschanged
dim rs
dim CurrentUserID
dim strSEPM
dim CurrentDomain
dim Currentuser

CurrentUser = lcase(Session("LoggedInUser"))
  	
if instr(CurrentUser,"\") > 0 then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	Currentuser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
end if

set cn = server.CreateObject("ADODB.Connection")
cn.ConnectionString = Session("PDPIMS_ConnectionString")
cn.Open

set rs = server.CreateObject("ADODB.Recordset") 	
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
	CurrentUserID = rs("ID") & ""
end if
rs.Close

set rs = server.CreateObject("ADODB.Recordset") 
set cm = server.CreateObject("ADODB.Command")
Set cm.ActiveConnection = cn
cm.CommandType = 4
cm.CommandText = "spGetProductVersion_Pulsar"
		
Set p = cm.CreateParameter("@ID", 3, &H0001)
p.Value = request("ProductVersionID")
cm.Parameters.Append p

rs.CursorType = adOpenForwardOnly
rs.LockType=AdLockReadOnly
Set rs = cm.Execute 
Set cm=nothing

strSEPM = rs("SEPMID") & ""
rs.Close

if strSEPM = CurrentUserID then
    ReleaseIDs = ""
    arrReleaseID = split(request("ReleaseIDs"), ",")

    For Each item In arrReleaseID
        if (item <> "") then
            if ReleaseIDs = "" then
                ReleaseIDs = item
            else
                ReleaseIDs = ReleaseIDs & "," & item
            end if
        end if
    Next    
	
    cn.BeginTrans

    set cm = server.CreateObject("ADODB.Command")
    cm.CommandType =  &H0004
    cm.ActiveConnection = cn
		
    cm.CommandText = "usp_ProductVersion_AssignRelease"	

    Set p = cm.CreateParameter("@ID", 3,  &H0001)
    p.Value = clng(request("ProductVersionID"))
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@ReleaseIDs", 200,  &H0001, 128)
    p.Value = ReleaseIDs
    cm.Parameters.Append p

    cm.Execute rowschanged
    set cm=nothing

    if cn.Errors.count > 0 then
	    response.Write "0"
	    cn.RollbackTrans
    else
	    cn.CommitTrans
    end if
else
    response.Write "Only SEPM can edit release"
end if

cn.Close
set cn = nothing
%>
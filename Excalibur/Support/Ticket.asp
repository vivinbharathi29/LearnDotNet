<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<TITLE>Support Ticket</TITLE>

</HEAD>

<%
	dim cn, rs, cm
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    dim blnEdit
    blnEdit = false

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID

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
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = ""
	end if
	
	rs.Close

    rs.open "spSupportIsAdminSelect " & clng(CurrentUserID),cn
    if rs.eof and rs.bof then
        blnEdit = false
    else
        blnEdit = true
    end if

    cn.Close
    set rs = nothing
    set cn = nothing

    if blnEdit then
%>


        <FRAMESET ROWS="*,57" ID=TopWindow >
	        <FRAME noresize ID="MainWindow" Name="MainWindow" SRC="TicketMain.asp?ID=<%=request("ID")%>">
	        <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="TicketButtons.asp">
        </FRAMESET>
    <%else %>
        <FRAMESET ROWS="*" ID=TopWindow >
	        <FRAME noresize ID="MainWindow" Name="MainWindow" SRC="TicketPreview.asp?ID=<%=request("ID")%>">
        </FRAMESET>

    <%end if%>

</HTML>
<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<% 
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim TypeID : TypeID = regEx.Replace(Request.QueryString("Type"), "")
    Dim ProdID : ProdID = regEx.Replace(Request.QueryString("ProdID"), "")
    Dim IssueID : IssueID = regEx.Replace(Request.QueryString("ID"), "")
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")
    
	dim cn
	dim rs
	dim CurrentUser
	dim CurrentUserPartner

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

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
	
	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../../NoAccess.asp?Level=0"
	else
		CurrentUserPartner = rs("PartnerID")
	end if 
	rs.Close


	dim strTypeID
	
	strTypeID=""
	if IssueID <> "" then 

		dim strProdID
		dim strActionType
		dim strPartnerID
		
		rs.Open "spGetActionProductType " &  clng(IssueID) ,cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then 		
			strTypeID = rs("TypeID") & ""
			strProdID = rs("ProdID") & ""
			strActionType = rs("ActionType") & ""
			strPartnerID = trim(rs("PartnerID") & "")
		end if
	    rs.Close
	end if

	'Verify Access is OK
	'if trim(CurrentUserPartner) <> "1" then
	'	if trim(strPartnerID) <> trim(CurrentUserPartner) then
	'		Response.Redirect "../../NoAccess.asp?Level=0"
	'	end if
	'end if


	dim ConfigErrors
	
	Set MyBrow=Server.CreateObject("MSWC.BrowserType")
		
	ConfigErrors = ""
	if lcase(MyBrow.browser) <> "ie" or clng(left(MyBrow.version,1)) < 4 then
		ConfigErrors = ConfigErrors & "Internet Explorer 4.0 or greater is required<br>"		
	end if
	
	if not MyBrow.frames then
		ConfigErrors = ConfigErrors & "Frames must be enabled<br>"
	end if
	if not MyBrow.tables then
		ConfigErrors = ConfigErrors & "Table support is required.<br>"
	end if
	if not MyBrow.cookies then
		ConfigErrors = ConfigErrors & "Cookie support is required.<br>"
	end if
	if not MyBrow.javascript then
		ConfigErrors = ConfigErrors & "Javascript support is required.<br>"
	end if
		
	dim strType
	rs.Open "spGetActionType " &  clng(TypeID) ,cn,adOpenForwardOnly
    If Not rs.EOF Then
        strType = rs("Name")
    End If
    rs.Close

			if IssueID <> "" then %>
				<TITLE><%=strType%> Properties</TITLE>
			<% else%>
				<%	if CategoryID = "1" then%>
				<TITLE>Add New SKU Change Request</TITLE>
				<%else%>
				<TITLE>Add New <%=strType%></TITLE>
				<%end if%>
			<% end if%>
			<HEAD>

			</HEAD>
			<FRAMESET ROWS="*,55" ID=TopWindow>
				<FRAME ID="UpperWindow" Name="UpperWindow" SRC="actionmain.asp?ProdID=<%=ProdID%>&CAT=<%=CategoryID%>&ID=<%=IssueID%>&Type=<%=TypeID%>">
					<FRAME ID="LowerWindow" Name="LowerWindow" SRC="actionbuttons.asp?ProdID=<%=ProdID%>&CAT=<%=CategoryID%>&ID=<%=IssueID%>&Type=<%=TypeID%>" scrolling="no">

				
			</FRAMESET>
   	set rs = nothing
	cn.Close
	set cn = nothing

%>
	
</HTML>
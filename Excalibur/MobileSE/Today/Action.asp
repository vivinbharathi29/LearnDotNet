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
	Dim Layout : Layout = Request.QueryString("layout")

    'If Type 7 (ECR) Redirect to ECR Action page.
    If IsNumeric(TypeID) And TypeID <> "" Then
        If CLng(Trim(TypeID)) = 7 Then Response.Redirect "../../Service/Action.asp?" & Request.QueryString
    End If
    
    If Trim(TypeID) = "" And Trim(IssueID) = "" Then
        Response.Write "<h1>Insufficient Information</h1><h2>Please launch this page from within the Pulsar application</h2><h3><a href=""/excalibur.asp"">Click Here to go to Pulsar</a></h3>"
        response.End
    End If
    
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
		rs.close
	
	end if

	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strPartnerID) <> trim(CurrentUserPartner) then
			Response.Redirect "../../NoAccess.asp?Level=0"
		end if
	end if


	dim ConfigErrors
			
	if trim(strActionType) = "6" then
		Response.redirect "../../TestManager/TestRequest.asp?ActionID=" & IssueID 
	elseif trim(strTypeID) = "2" or clng(TypeID) = 2 then 'Action Item or Tools Project
		Response.redirect "../../actions/action.asp?ID=" & IssueID & "&Type=" & strActionType & "&ProdID=" & strProdID & "&Working=0"
	elseif 1=1 then 'ConfigErrors = "" then

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

			    <FRAME ID="UpperWindow" frameborder="0" Name="UpperWindow" SRC="actionmain.asp?ProdID=<%=ProdID%>&CAT=<%=CategoryID%>&ID=<%=IssueID%>&Type=<%=TypeID%>&Layout=<%=Layout%>">
			    <FRAME ID="LowerWindow" frameborder="0" Name="LowerWindow" SRC="actionbuttons.asp?ProdID=<%=ProdID%>&CAT=<%=CategoryID%>&ID=<%=IssueID%>&Type=<%=TypeID%>&Layout=<%=Layout%>" scrolling=no>

			</FRAMESET>
	   <%'end if
	else%>
		<HEAD>
		</HEAD>
		<BODY>
			<h4>Your browser configuration is not compatible with this application.</h4>
			<font face = verdana size=2>Please correct the following configuration errors and try again:<BR><BR>
			<%="<font face=verdana color=red size=2>" & ConfigErrors & "</font>"%>
			<BR><BR></font>
			<a href="mailto:max.yu@hp.com?Subject=Browser Compatibility">Send Email to Application Administrator</a><br>
		</BODY>
<%
    end if
   	set rs = nothing
	cn.Close
	set cn = nothing
%>
	
</HTML>
<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%=request("optChoose")%><BR>
<%=request("cboRoot")%><BR>
<%=request("txtDeliverableID")%><BR>
<%=request("txtproductID")%><BR>


<%

	dim cn
	dim rs
	dim strPRID
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	
	
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
		CurrentUserID = rs("ID")
	end if
	if CurrentUserID = "" then
		CurrentUserID = 0
	end if
	rs.Close   		
	
	
	
	
	
	rs.Open "Select pr.id, from prodReq_DelRoot pdr with (NOLOCK), product_requirement pr with (NOLOCK) where pr.id = pdr.productrequirementid and pr.productId = " & clng(request("txtproductID")) & " and pdr.deliverablerootid = " & clng(request("txtDeliverableID")) ,cn,adOpenForwardOnly

	do while not rs.EOF
	'Remove Roots

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn	
		cm.CommandText = "spRemoveDelRootFromProductReq"	
			
		Set p = cm.CreateParameter("@PRID", 3,  &H0001)
		p.Value = rs("id")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = clng(request("txtproductID"))
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@RootID", 3,  &H0001)
		p.Value = clng(request("txtDeliverableID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
		p.Value = clng(CurrentUserID)
		cm.Parameters.Append p

		'cm.Execute rowschanged
		Response.Write "<br><br>"
		if cn.Errors.count > 0 then
			FoundErrors = true
		end if
					
		set cm = nothing
		rs.MoveNext
	loop
	set rs=nothing
	cn.Close
	set cn = nothing
%>
</BODY>
</HTML>

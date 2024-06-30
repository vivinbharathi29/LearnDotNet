<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">



<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function PrintScreen(){
	window.focus();
	window.print();
}

//-->
</script>

</head>
<body bgcolor="ivory" LANGUAGE=javascript>

<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">


<%
	'Create Database Connection
	dim cn 
	dim rs
	dim cm
	dim p
	dim strSQL
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserPartner
	dim strChanges
	dim strVersion
	dim HiddenCount
	
	HiddenCount=0
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.ConnectionTimeout = 60
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

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUSerID = rs("ID")
		CurrentUserPartner = trim(rs("PartnerID") & "")
	end if
	
	rs.Close
	
	
	'rs.Open "spGetOTSNote '70000'",cn,adOpenForwardOnly
	'rs.Close

	rs.Open "spGetVersionProperties4Web " & clng(request("ID")),cn,adOpenForwardOnly
	
	if rs.EOF and rs.BOF then
		strChanges = "none specified"
	elseif trim(rs("Changes") & "")  = "" then
		strChanges = "none specified"
	else
		strChanges = replace(rs("Changes") & "" ,vbcrlf,"<BR>")
	end if
	
	strVersion = rs("Version") & ""
	if rs("Revision") & "" <> "" then
		strVersion = strVersion &  "," &  rs("Revision") 
	end if
	if rs("Pass") & "" <> "" then
		strVersion = strVersion & "," &  rs("Pass") 
	end if
	Response.write "<h3>" & rs("Name") & " " & strVersion & " Changes</h3>"
	
	rs.Close
	%>
    
</TABLE>
<form id="frmFT" method="post" action="FTSave.asp">
<input style="Display:none" type="text" id="ID" name="ID" value="<%=request("ID")%>">
<input style="Display:none" type="text" id="RootID" name="RootID" value="<%=request("RootID")%>">

<font size=2 face=verdana><b>Observations Fixed in this Release</b></font><br>
<table ID="tabRoot" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">

<%

	dim strProductName
	on error resume next
	rs.Open "spGetOTSByDelVersion " & clng(request("ID")),cn,adOpenForwardOnly
	on error goto 0
	if cn.Errors.count> 0 then
		Response.Write "<font size=2 color=red and family=verdana><b>OTS is unavailable.</b></font>"
		cn.Errors.clear
	else
	
	if rs.EOF and rs.BOF then
		Response.Write "<tr><td colspan=2 style=""Width:100%""><font size=2face=verdana>none specified</font></td></tr>"
	else
		do while not rs.EOF
			'if currentuserpartner <> "1" then
				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.Open "spGetProductVersionByName '" & rs("platform") & "'",cn,adOpenForwardOnly
				if rs2.EOF and rs2.BOF then
					strProductName=rs("platform") & ""
				else
					if currentuserpartner = 1 or (trim(currentuserpartner) = trim(rs2("PartnerID"))) then
						strProductName=rs("platform") & ""
					else
						strProductName= "ID:" & rs2("ID") & " (Product Name Hidden)"
					end if
				end if
				
				set rs2=nothing
			'end if
		
%>


	<tr>
		<td  valign=top width="150" nowrap><b>Observation:</b></td>
		<td  valign=top ><a target=OTS href="../mobilese/today/setupots.asp?USERID=<%=CurrentUSerID%>&OTSID=<%=rs("OTSNumber")%>"><%=rs("OTSNumber") & ""%></a><%= " - " & rs("ShortDescription")%></td>
	</TR>
	<tr>
		<td valign=top width="150" nowrap><b>Product:</b></td>
		<td  valign=top style="Width:100%"><%= strproductName %></td>
	</tr>
	<tr>
		<td valign=top width="150" nowrap><b>Reproduce:</b></td>
		<td  valign=top style="Width:100%"><%=replace(rs("StepsToReproduce"),vblf,"<BR>")%></td>
	</tr>
	
	<%
				rs.MoveNext
			loop	
		end if
		rs.Close
	end if

	%>
    
</TABLE>
<BR><BR>
<font size=2 face=verdana><b>Other Changes Made in this Release</b></font><br>
<table ID="tabRoot" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td  valign=top width="100%"><%=strChanges %></td>
	</TR>
    
</TABLE>
</select>	

</body>
</html>
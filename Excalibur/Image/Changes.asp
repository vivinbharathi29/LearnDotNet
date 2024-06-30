<%@ Language=VBScript %>

<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<html>

<STYLE>

.Details TABLE
{
    BORDER-RIGHT-STYLE: thin;
    BORDER-TOP-STYLE: thin;
    BORDER-LEFT-STYLE: thin;
    BORDER-BOTTOM-STYLE: thin;
}
.Details THEAD
{
    BACKGROUND-COLOR: gainsboro;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
	FONT-WEIGHT: bold;
}
.Details TBody
{
    BACKGROUND-COLOR: white;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
}
.Details TD
{
    BORDER-RIGHT-STYLE: thin;
    BORDER-TOP-STYLE: thin;
    BORDER-LEFT-STYLE: thin;
    BORDER-BOTTOM-STYLE: thin;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
}



</STYLE>

<body bgColor=white>
<%

	dim cn
	dim cm
	dim p
	dim rs
	dim strError
	dim strProdName
	dim strPOR
	dim CurrentUser
	dim CurrentUserPartner

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	if request("ID") = "" or not isnumeric(request("ID")) then
		strError = "No Product Specified."
	else
	
	
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
			Response.Redirect "../NoAccess.asp?Level=0"
		else
			CurrentUserPartner = rs("PartnerID")
		end if 
		rs.Close		
	
	
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersionName"
		
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p
	
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing


		strError = ""
		'rs.Open "spGetProductVersionName " & request("ID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strError = "Could not find the selected product."
		else
			strProdName = rs("Name") & ""
			strPOR = rs("PDDReleased") & ""

			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					set rs = nothing
					set cn=nothing
				
					Response.Redirect "../NoAccess.asp?Level=0"
				end if
			end if
			
		end if
		rs.Close

	end if
	
	if trim(strPOR) = "" then
		strError = "No changes are logged before Product POR."
	else
		strPOR = formatdatetime(strPOR ,vbshortdate)	
	end if
	if strError <> "" then
		Response.Write "<font size=2 face=verdana><b>" & strError & "</b></font>"
	else
%>


<center><font face=verdana size=4><b><%=strProdName%> Localization Changes</b></font>
	<font face=verdana size=2><br><br>Changes since POR (<%=strPOR%>)<BR><BR></font>
</center>
<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListLocalizationChanges"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

'	rs.Open "spListLocalizationChanges " & request("ID"),cn,adOpenForwardOnly
	do while not rs.eof
		if trim(rs("SKUNUmber")& "") <> "" then
			Response.Write "<TR BGCOLOR=cornsilk><TD><STRONG>SKU:</STRONG> " & rs("SKUNUmber") & "</TD>"
		else
			Response.Write "<TR BGCOLOR=cornsilk><TD><STRONG>SKU:</STRONG> Not Specified (ID=" & rs("ID") & ")</TD>"
		end if
		Response.Write "<TD><STRONG>By:</STRONG> " & rs("Employee") & "</TD>"
		Response.Write "<TD><STRONG>Updated:</STRONG> " & rs("ModDate") & "</TD>"
		'Response.Write "<TD><STRONG>DCR ID:</STRONG> " & rs("DCRID") & " - " & rs("Summary") &  "</TD></TR>"
		Response.Write "<TR><TD colspan=3 STYLE=""PADDING-LEFT: 10px"">"
		
		if trim(rs("DCRID") & "") <> "" then 
			Response.Write "<BR><STRONG>DCR ID:</STRONG> " & rs("DCRID") & " - " & rs("Summary") & "<BR>"
		end if
		if rs("Details") & "" = "Added" then
			Response.Write "<BR><b>SKU Added</b><br><br>"
		elseif rs("Details") & "" = "Deleted" then
			Response.Write "<BR><b>SKU Deleted</b><br><br>"
		else
			Response.Write "<BR><TABLE border=1 width=400 class=""Details""><THEAD><TR><TD width=200><b>Field</b></TD><TD width=100><b>Old&nbsp;Value</b></TD><TD width=100><b>New&nbsp;Value</b></TD></tr></THEAD><TBODY>"
			Response.Write rs("Details") & "</TBODY></TABLE><BR>"
		end if		
		rs.MoveNext
	loop
	rs.close

	Response.Write "</td></tr>"
	
	
	end if

	cn.Close
	set rs = nothing
	set cn = nothing
%>
</TABLE>


</body>
</html>
 

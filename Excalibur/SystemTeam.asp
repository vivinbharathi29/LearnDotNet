<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <title>Full System Team Roster</title>
    <STYLE>
    TD{
	    font-family:Verdana;
	    font-size:x-small;
	    background-color:ivory;
	    BORDER-TOP: lightgrey 1px solid;	

    }
    TH{
	    font-family:Verdana;
	    font-size:x-small;
	    background-color:beige;
	    TEXT-ALIGN:left;
	    BORDER-TOP: lightgrey 1px solid;	
    }
    </STYLE>
</HEAD>

<BODY>


<%
	dim cn
	dim rs
	dim strRow
	dim strDevCenter
	set cn = server.CreateObject("ADODB.Connection")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
    Dim IsPulsarProduct : IsPulsarProduct = 0

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

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	
	if (rs.EOF and rs.BOF) then
		set rs = nothing
       	set cn=nothing
       	Response.Redirect "NoAccess.asp?Level=0"
    else
        CurrentUserPartner = rs("PartnerID")
        CurrentUserID = rs("ID")
    end if 
    rs.Close
	
	
	
	
	dim strProductName
	dim strProductPartner 

	rs.Open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
	if (rs.EOF and rs.BOF) then
		strProductName = ""
		strProductPartner = ""
		strDevCenter = ""
        IsPulsarProduct = 0
	else
		strProductName = rs("Name") & ""
		strProductPartner = trim(rs("PartnerID") & "")
		strDevCenter = trim(rs("DevCenter") & "")
        IsPulsarProduct = rs("FusionRequirement")
	end if
	rs.Close
	
	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strProductPartner) <> trim(CurrentUserPartner) then
			set rs = nothing
			set cn=nothing
			
			Response.Redirect "NoAccess.asp?Level=0"
		end if
	end if
	
	
	
	Response.Write "<font size=2 face=verdana><b>" & strProductName & " System Team Roster</b></font><BR><BR>"
    if IsPulsarProduct then
        rs.Open "spListSystemTeam " & clng(request("ID")) & ",0,1",cn,adOpenForwardOnly
    else
    	rs.Open "spListSystemTeam " & clng(request("ID")),cn,adOpenForwardOnly
	end if
    strRow=""
	dim Team
    dim RoleName : RoleName = ""
	Team = 1
	do while not rs.EOF
		if Team = 1 and rs("PrimaryTeam") = 0 then
			team = 0
			strRows = strRows & "<TR><TH colspan=3>Extended System Team</TH></TR>"
		end if
        if ((rs("Role") <> "SMB Marketing" and rs("Role") <> "Consumer Marketing" and IsPulsarProduct = True) or (IsPulsarProduct = False)) then
            RoleName = rs("Role")
            if (rs("Role") = "Commercial Marketing" and IsPulsarProduct = True) then
                RoleName = "Marketing/Product Mgmt"
            end if
	        strRows = strRows & "<TR><TD>" & RoleName & "&nbsp;&nbsp;&nbsp;</TD><TD><a href=""mailto:" & rs("Email") & """>" & rs("Name") & "</a>&nbsp;&nbsp;&nbsp;</TD><TD>" & rs("Phone") & "</TD></TR>"	
        end if
		rs.MoveNext
	loop
	rs.Close

	if trim(strRows) = "" then
		Response.Write "<font face=verdana size=2><BR>No System Team defined for this product.</font>"
	else
		Response.Write "<table cellspacing=0 cellpadding=1><tr><th>Function</th><th>Name</th><th>Phone</th></tr>" & strRows & "</table>"
		Response.Write "<font size=1 face=verdana><BR><BR>Please contact one of the System Team managers to identify any team members not specified here.</font>"
	end if
	
	rs.Open "spListSystemTeam_Original " & clng(request("ID")),cn,adOpenForwardOnly
    If Not (rs.EOF and rs.BOF) Then
        strRows=""
       	Response.Write "<br><br><br><font size=2 face=verdana><b>" & strProductName & " Original System Team Roster</b></font><BR><BR>"
	    strRow=""
	    Team = 1
	    do while not rs.EOF
		    if Team = 1 and rs("PrimaryTeam") = 0 then
			    team = 0
			    strRows = strRows & "<TR><TH colspan=3>Extended System Team</TH></TR>"
		    end if
			    strRows = strRows & "<TR><TD>" & rs("Role") & "&nbsp;&nbsp;&nbsp;</TD><TD><a href=""mailto:" & rs("Email") & """>" & rs("Name") & "</a>&nbsp;&nbsp;&nbsp;</TD><TD>" & rs("Phone") & "</TD></TR>"
		    rs.MoveNext
	    loop
        Response.Write "<table cellspacing=0 cellpadding=1><tr><th>Function</th><th>Name</th><th>Phone</th></tr>" & strRows & "</table>"
    End If

	rs.Close
	set rs = nothing
	
	cn.Close
	set cn=nothing

	%>


</BODY>
</HTML>

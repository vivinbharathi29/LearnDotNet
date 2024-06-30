<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function ShowOTSDetails(strID) {

	var NewLeft = (screen.width - 655)/2;
	var NewTop = (screen.height - 650)/2;

	window.open("../search/ots/Report.asp?txtReportSections=1&txtObservationID=" + strID,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No,scrollbars=Yes") 
}


function CompareDeliverables(strID) {

	var NewLeft = (screen.width - 655)/2;
	var NewTop = (screen.height - 650)/2;

	window.open("../MobileSE/Today/CompareVersions.asp?ProdID=" + txtID.value + "&ID=" + strID,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No,scrollbars=Yes") 
}

//-->
</SCRIPT>
</HEAD>
<Style>
TD{
    FONT-Size: xx-small;
}
TH{
    FONT-Size: xx-small;
}
</Style>
<BODY>

<%

dim strProduct
dim cm
dim p
dim CurrentUser
dim CurrentUserPartner


strproduct=""

if request("ID") = "" then
	response.write "Not Enough Information Supplied for this Report"	
else
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
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

	'rs.Open "spGetProductVersionName " & request("ID"),cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		strProduct = rs("name") & ""

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

if strProduct <> "" then
%>

<h3><%=strProduct%> Deliverable Status</h3>

<b>Open Observations on Targeted Deliverables</b><BR>
<%
	dim strID
	dim strVersion
	dim strLastDeliverable


	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetOTSTargetedObservations"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Report", 3, &H0001)
	p.Value = 1
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	'rs.Open "spGetOTSTargetedObservations " & request("ID"),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "none<BR><BR>"
		rs.Close
	else
		Response.Write "<TABLE width=100% class=HighlightRow border=1 cellspacing=1 cellpadding=2><THEAD><TR>"
		Response.Write "<TH>ID</TH>"
		Response.Write "<TH>Owner</TH>"
		Response.Write "<TH>Pr</TH>"
		Response.Write "<TH>State</TH>"
		Response.Write "<TH nowrap>Targeted</TH>"
		Response.Write "<TH nowrap>Found On</TH>"
		Response.Write "<TH nowrap>Developer</TH>"
		Response.Write "<TH>Summary</TH>"
		Response.Write "</TR></THEAD>"
		do while not rs.EOF
			if strLastDeliverable <> rs("Deliverable") then
				Response.Write "<TR><TD class=BrownRow colspan=8>" & rs("Deliverable") & "</TD>"
			end if
			strLastDeliverable = rs("Deliverable") & ""
			
			strVersion = rs("Version") & ""
			if trim(rs("Revision")) & "" <> "" then
				strVersion = strVersion & "," &  rs("Revision")
			end if
			if trim(rs("Pass")) & "" <> "" then
				strVersion = strVersion & "," &  rs("Pass")
			end if
			
			if not isnull(rs("ObservationID")) then
				Response.Write "<TR><TD><a href=""javascript: ShowOTSDetails('" & rs("ObservationID") & "')"">" & rs("ObservationID") & "</a></TD>"
				Response.Write "<TD>" & rs("owner") & "</TD>"
				Response.Write "<TD>" & rs("priority") & "</TD>"
				Response.Write "<TD>" & rs("State") & "</TD>"
				Response.Write "<TD><a href=""javascript: CompareDeliverables(" & rs("RootID") & ")"">" & strVersion & "</a></TD>"
				Response.Write "<TD>" & rs("OTSComponentVersion") & "</TD>"
				Response.Write "<TD>" & rs("Developer") & "</TD>"
				Response.Write "<TD>" & rs("Summary") & "</TD>"
				Response.Write "</TR>"
			end if
		
		
			rs.MoveNext
		loop
		
		Response.Write "</TABLE><BR><BR>"
		rs.Close
	end if


%>

<b>No Deliverable Version Targeted</b><BR>

<%

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListRootsWithNoTargetedVersions"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	'rs.Open "spListRootsWithNoTargetedVersions " & request("ID"),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "none<BR><BR>"
		rs.Close
	else
		Response.Write "<TABLE width=100% class=HighlightRow border=1 cellspacing=1 cellpadding=2><THEAD><TR>"
		Response.Write "<TH>Deliverable</TH>"
		Response.Write "<TH nowrap>Development Manager</TH>"
		Response.Write "</TR></THEAD>"
		do while not rs.EOF
			Response.Write "<TR><TD>" & rs("Name") & "</TD>"
			Response.Write "<TD>" & rs("DevManager") & "</TD>"
			Response.Write "</TR>"
		
			rs.MoveNext
		loop
		
		Response.Write "</TABLE>"
		rs.Close
	end if



%>
<BR><BR>
<b>Targeted Deliverables with no Open Observations</b><BR>
<%

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetOTSTargetedObservations"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Report", 3, &H0001)
	p.Value = 2
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	'rs.Open "spGetOTSTargetedObservations " & request("ID") &  ",2",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "none<BR><BR>"
		rs.Close
	else
		Response.Write "<TABLE width=100% class=HighlightRow border=1 cellspacing=1 cellpadding=2><THEAD><TR>"
		Response.Write "<TH>ID</TH>"
		Response.Write "<TH>Name</TH>"
		Response.Write "<TH>Version</TH>"
		Response.Write "<TH>Developer</TH>"
		Response.Write "<TH>Notes</TH>"
		Response.Write "</TR></THEAD>"
		do while not rs.EOF
			if isnull(rs("ObservationID")) then
			
				strVersion = rs("Version") & ""
				if trim(rs("Revision")) & "" <> "" then
					strVersion = strVersion & "," &  rs("Revision")
				end if
				if trim(rs("Pass")) & "" <> "" then
					strVersion = strVersion & "," &  rs("Pass")
				end if
				
				Response.Write "<TR><TD><a href=""javascript: CompareDeliverables(" & rs("RootID") & ")"">" & rs("ID") & "</a></TD>"
				Response.Write "<TD>" & rs("Deliverable") & "</TD>"
				Response.Write "<TD>" & strVersion & "</TD>"
				Response.Write "<TD nowrap>" & rs("Developer") & "</TD>"
				Response.Write "<TD>" & rs("TargetNotes") & "&nbsp;</TD>"
				Response.Write "</TR>"
			end if		
			rs.MoveNext
		loop
		
		Response.Write "</TABLE><BR><BR>"
		rs.Close
	end if

	set rs = nothing
	cn.Close
	set cn=nothing
end if

%>

<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
</BODY>
</HTML>

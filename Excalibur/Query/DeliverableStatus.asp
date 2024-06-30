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
	
%>

<%if request("txtTitle") = "" then%>
	<h3>Deliverable Status</h3>
<%else%>
	<h3><%=request("txtTitle")%></h3>
<%end if%>



<br><br>
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

end if
%>

<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
</BODY>
</HTML>

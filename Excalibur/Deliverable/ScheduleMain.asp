<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../includes/Date.asp" -->



function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }

function dateAdd( start, interval, number ) {
	
    // Create 3 error messages, 1 for each argument. 
    var startMsg = "Sorry the start parameter of the dateAdd function\n"
        startMsg += "must be a valid date format.\n\n"
        startMsg += "Please try again." ;
		
    var intervalMsg = "Sorry the dateAdd function only accepts\n"
        intervalMsg += "d, h, m OR s intervals.\n\n"
        intervalMsg += "Please try again." ;

    var numberMsg = "Sorry the number parameter of the dateAdd function\n"
        numberMsg += "must be numeric.\n\n"
        numberMsg += "Please try again." ;
		
    // get the milliseconds for this Date object. 
    var buffer = Date.parse( start ) ;
	
    // check that the start parameter is a valid Date. 
    if ( isNaN (buffer) ) {
        alert( startMsg ) ;
        return null ;
    }
	
    // check that an interval parameter was not numeric. 
    if ( interval.charAt == 'undefined' ) {
        // the user specified an incorrect interval, handle the error. 
        alert( intervalMsg ) ;
        return null ;
    }

    // check that the number parameter is numeric. 
    if ( isNaN ( number ) )	{
        alert( numberMsg ) ;
        return null ;
    }

    // so far, so good...
    //
    // what kind of add to do? 
    switch (interval.charAt(0))
    {
        case 'd': case 'D': 
            number *= 24 ; // days to hours
            // fall through! 
        case 'h': case 'H':
            number *= 60 ; // hours to minutes
            // fall through! 
        case 'm': case 'M':
            number *= 60 ; // minutes to seconds
            // fall through! 
        case 's': case 'S':
            number *= 1000 ; // seconds to milliseconds
            break ;
        default:
        // If we get to here then the interval parameter
        // didn't meet the d,h,m,s criteria.  Handle
        // the error. 		
        alert(intervalMsg) ;
        return null ;
    }
    return new Date( buffer + number ) ;
}


function cmdDate_onclick(strID){
	var strRC;
	var strRelease;
	
	strRC = window.showModalDialog("../mobilese/today/caldraw1.asp",document.getElementById("txtDate" + strID).value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strRC) != "undefined" )
		{
			var d = new Date(strRC);
			d=dateAdd(d,"d",1);
			var curr_date = d.getDate();
			var curr_month = d.getMonth();
			curr_month++;
			var curr_year = d.getFullYear();

			strRelease = curr_month + "/" + curr_date + "/" + curr_year;
	
			
			document.getElementById("txtDate" + strID).value=strRC;
	
			if (typeof(lblRelease)!= "undefined")
				{
				lblRelease.innerText = strRelease;
				frmUpdate.txtRelease.value = strRelease;
				}
			
		}
	
	

}

function y2k(number) {
	if (number < 50)
		return number +2000;
	else if (number <100)
		return number+1900;
	else		
		return number;
 
}

function UpdateReleaseDate(){
	if (typeof(lblRelease)!= "undefined" && isDate(document.getElementById("txtDate" + txtLastMilestone.value).value))
		{
		var d = new Date(document.getElementById("txtDate" + txtLastMilestone.value).value);
		d=dateAdd(d,"d",1);
		var curr_date = d.getDate();
		var curr_month = d.getMonth();
		curr_month++;
		var curr_year = y2k(d.getYear());

		strRelease = curr_month + "/" + curr_date + "/" + curr_year;
		lblRelease.innerText = strRelease;
		frmUpdate.txtRelease.value = strRelease;
		}

}


function window_onload() {
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>

<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">


<%

	dim cn
	dim rs
	dim cm 
	dim p
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserGroup
	dim blnFound
	dim strLastMilestone
	dim strMilestones
	
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
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserGroup = rs("workgroupID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	blnFound = false
	if request("ID") <> "" then
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetDeliverableversionproperties"
		
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p


		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		'rs.Open "spGetDeliverableversionproperties " & request("ID"),cn,adOpenForwardOnly
		
		if not (rs.EOF and rs.BOF) then
			strDeliverable = rs("Name") & " [" &  rs("Version")
			if rs("Revision") <> "" then
				strDeliverable = strDeliverable & "," & rs("Revision")
			end if
			if rs("Pass") <> "" then
				strDeliverable = strDeliverable & "," & rs("Pass")
			end if
			strDeliverable = strDeliverable & "]"
			Response.Write "<h4>" & strDeliverable & "</h4>"
			blnFound = true
		end if
		rs.Close
	end if

	if not blnFound then
		Response.Write "<BR><font size=2 face=verdana>Unable to find selected deliverable.</font>"
	else
%>


<form id="frmUpdate" method="post" action="ScheduleSave.asp">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
<%

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetDelMilestoneList"
	

	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

'	rs.open "spGetDelMilestoneList 0," & request("ID")  ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD>No Workflow Steps Found</TD></TR>"
	else
		Response.Write "<TR><TD><b>Workflow</b></TD><TD><b>Planned</b></TD><TD><b>Actual</b></TD></TR>"
		do while not rs.EOF
			response.write "<TR>"
			Response.Write "<td valign=top nowrap>" & rs("Milestone") & "</td>"
			if lcase(trim(rs("Milestone"))) = "release team" then
				if strMilestones <> "" then
					Response.Write "<td valign=top nowrap><label ID=lblRelease>" & rs("Planned") & "</label>&nbsp;<font size=1 face=verdana color=blue>(Updated&nbsp;Automatically)</font>"
				else
					Response.Write "<td valign=top nowrap><label ID=lblRelease>" & rs("Planned") & "</label>"
				end if
				Response.Write "<INPUT type=""hidden"" id=tagRelease name=tagRelease value=""" & rs("Planned") &  """>"
				Response.Write "<INPUT type=""hidden"" id=txtRelease name=txtRelease value=""" & rs("Planned") &  """>"
				Response.Write "</td>"
				strMilestones = strMilestones & "," & rs("ID")
			elseif not isnull(rs("Actual")) then
				Response.Write "<td valign=top nowrap>" & rs("Planned") & "</td>"
			else
				Response.Write "<td nowrap valign=top><INPUT class=""text"" type=""text"" id=txtDate" & trim(rs("ID")) & " name=txtDate value=""" & formatdatetime(rs("Planned"),2) &  """ LANGUAGE=javascript onfocusout=""return UpdateReleaseDate()""><INPUT type=""hidden"" id=tagDate" & trim(rs("ID")) & " name=tagDate value=""" & rs("Planned") &  """>&nbsp;"
				Response.Write "<a href=""javascript: cmdDate_onclick(" & rs("ID") & ")""><img ID=""picTarget"" SRC=""../mobilese/today/images/calendar.gif"" alt=""Choose Date"" border=0 WIDTH=26 HEIGHT=21></a></td>"
				strLastMilestone = rs("ID")
				strMilestones = strMilestones & "," & rs("ID")
			end if
			if isnull(rs("Actual")) then
				Response.Write "<td width=120 valign=top nowrap>&nbsp;</td>"
			else
				Response.Write "<td width=120 valign=top nowrap>" & formatdatetime(rs("Actual"),vbshortdate) & "</td>"
			end if
			response.write "</TR>"
			
			rs.MoveNext
		loop
	end if
	rs.Close
	if len(strMilestones) > 0 then
		strMilestones = mid(strMilestones,2)
	end if
%>

</table>


<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtMilestones name=txtMilestones value="<%=strMilestones%>">

</form>
<%end if

	cn.Close
	set cn = nothing
	set rs = nothing


%>
<INPUT type="hidden" id=txtLastMilestone name=txtLastMilestone value="<%=strLastMilestone%>">
</BODY>
</HTML>



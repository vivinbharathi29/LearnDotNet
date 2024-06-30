<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var oPopup = window.createPopup();


function item_onmouseover() {
	window.event.srcElement.parentElement.style.backgroundColor="Lavender";
	window.event.srcElement.parentElement.style.cursor="hand";
}

function item_onmouseout() {
	window.event.srcElement.parentElement.style.backgroundColor="ivory";
}

function item_onclick(strID) {
	if (document.all("Details" + strID).style.display=="none")
		document.all("Details" + strID).style.display="";
	else
		document.all("Details" + strID).style.display="none";	
}

function actionitem_onclick(strID,blnEdit) {
	if (blnEdit!=1)
		window.open ("../Query/ActionReport.asp?ID=" + strID);
	else
		{
		var strResult;
		strResult = window.showModalDialog("action.asp?ID=" + strID ,"","dialogWidth:655px;dialogHeight:470px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No"); 
		if (typeof(strResult) != "undefined")
			window.location.reload(true);
		}
}

function EditMilestone(strID) {
		var strResult;
		strResult = window.showModalDialog("schedule.asp?ID=" + strID ,"","dialogWidth:655px;dialogHeight:470px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No"); 
		if (typeof(strResult) != "undefined")
			window.location.reload(true);
}


function ShowMenu(ProdID, RoadmapID) {
    var lefter = event.clientX //- event.offsetX;
    var topper = event.clientY //(event.clientY - event.offsetY)+ event.srcElement.offsetHeight;
    var popupBody;
    
	popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddTask(" + ProdID + "," + RoadmapID + ")'\" >&nbsp;&nbsp;&nbsp;Add&nbsp;Task&nbsp;</SPAN></FONT></DIV>";

//	popupBody = popupBody + "<DIV>";
//  popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
//
//	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
//	popupBody = popupBody + "<FONT face=Arial size=2>";
//	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ShowClosedTasks()'\" >&nbsp;&nbsp;&nbsp;Show&nbsp;Closed&nbsp;Tasks&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

	
	popupBody = popupBody + "</DIV>";
	oPopup.document.body.innerHTML = popupBody; 

	oPopup.show(lefter, topper, 130, 85, document.body);

	//Adjust window size
	if (oPopup.document.body.scrollHeight> 1 || oPopup.document.body.scrollWidth> 1)
		{
		NewHeight = oPopup.document.body.scrollHeight;
		NewWidth = oPopup.document.body.scrollWidth;
		oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
		}
		
}

function AddTask(ProdID, RoadmapID){
	var strID;
	strID = window.showModalDialog("action.asp?ID=0&Working=0&RoadmapID=" + RoadmapID + "&ProdID=" + ProdID + "&Type=2","","dialogWidth:655px;dialogHeight:485px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No"); 
	if (typeof(strID) != "undefined")
		{
			window.location.reload(true);
		}
}

function ShowClosedTasks(){
	alert("Not implemented yet.");
}

function TaskLink_onmouseover(){
	window.event.srcElement.style.color="red";
	window.event.srcElement.style.cursor="hand";
}

function TaskLink_onmouseout(){
	window.event.srcElement.style.color="black";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
TD
{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
	BORDER-TOP: lightgrey 1px solid;	
}
TH
{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
	TEXT-ALIGN: left
}

A:visited
{
	COLOR: Black;
}
A:hover
{
	COLOR: Red;
}
A
{
	COLOR: Black;
}
</STYLE>
<BODY>

<%
	dim cn
	dim rs
	dim cm
	dim rs2
	dim p
	dim TaskCount
	dim EditOK

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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("Email") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close

    EditOk = 0

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	'if currentuserid = 31 or currentuserid = 8 then
		'EditOK = 1 'Administrators
	'elseif (currentuserid=694 or currentuserid = 1396) and (request("ID") = 235) then
		'EditOK = 1 'Excalibur
	'elseif (currentuserid = 648 or currentuserid = 695) and (request("ID") = 309) then
		'EditOK = 1 'Release Tools
	'elseif (currentuserid = 695 or currentuserid=607 or currentuserid=647) and (request("ID") =300 or request("ID") = 319 ) then
		'EditOK = 1 'Conveyor, Preinstall Tools
	'elseif (currentuserid = 695) and (request("ID") = 310 ) then
		'EditOK = 1 'OTS
	'elseif (currentuserid = 695) and (request("ID") = 352 ) then
		'EditOK = 1 'Test management
	'else	
		'EditOK = 0
	'end if

	if request("ID") = "" then
		Response.Write "<b><font size=2 face=verdana>Not enough information suplied to display this page.</font></b>"
	elseif currentuserid=0 then
		Response.Write "<b><font size=2 face=verdana>You do not have access to display this page.</font></b>"
	else
		rs.Open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<b><font size=2 face=verdana>Product not found.</font></b>"
		else
			Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " Roadmap</b><BR><BR></font>"
		end if
		rs.Close
	
		rs.Open "spListActionRoadmapSummary " & clng(request("ID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<b><font size=2 face=verdana>No Roadmap found for this product.</font></b>"
		else
			Response.Write "<table width=""100%"" cellpadding=2 cellspacing=0 bgcolor=ivory>"
			Response.Write "<TR bgcolor=beige><TD><b>Task</b></TD><TD><b>Open</b></TD><TD><b>Closed</b></TD><TD><b>Estimated<BR>Completion</b></TD><TD><b>Primary<BR>Owner</b></TD><TD><b>Notes</b></TD></TR>"
			do while not rs.EOF
				set rs2 = server.CreateObject("ADODB.recordset")
				strTimeframe = rs("Timeframe") & ""
				if trim(strTimeframe) = "" then
					strTimeframe = "TBD"
				elseif isdate(strTimeframe) then
					if year(cdate(strTimeframe)) > 1600 then
						strTimeframe = formatdatetime(cdate(strTimeframe),vbshortdate)
					end if
				end if
			'	strStatus = rs("Status") & ""
			'	if strStatus <> "Blocked" then
					strOpenTasks = 0
					strClosedTasks = 0
					rs2.open "spGetActionRoadmapTaskCounts " & rs("ID"),cn,adOpenForwardOnly
					if not (rs.EOF and rs.BOF) then
						if trim(rs2("CompleteCount")) = "0" then
							strStatus = "Open"
						elseif trim(rs2("TaskCount")) <> trim(rs2("CompleteCount"))  then
							strStatus = "In Progress"
						else
							strStatus = "Complete"
						end if						
						strOpenTasks = rs2("TaskCount") - rs2("CompleteCount") 
						strClosedTasks = rs2("CompleteCount") 
					end if
					rs2.Close
			'	end if
				Response.Write "<TR LANGUAGE=javascript onclick=""return item_onclick(" & rs("ID") & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD>" & rs("Summary") & "&nbsp;&nbsp;</TD><TD>" & strOpenTasks & "&nbsp;&nbsp;</TD><TD nowrap>" & strClosedTasks & "&nbsp;&nbsp;</TD><TD>" & strTimeframe & "&nbsp;&nbsp;</TD><TD>" & shortname(rs("Owner") & "") & "&nbsp;&nbsp;</TD><TD>" & rs("Notes") & "&nbsp;&nbsp;</TD></TR>"
				rs2.open "spGetActionRoadmapItemProperties " & rs("ID"),cn,adOpenForwardOnly
				strDetails = replace(rs2("Details") & "",vbcrlf,"<BR>")

				if strDetails = "" then
					strDetails = "No additional information found for this item."
				end if
				Response.Write "<TR bgcolor=white style=""Display:none"" ID=Details" & trim(rs("ID")) & "><TD colspan=6><Blockquote><b>"
				if EditOK<>1 then
					Response.Write "<u>Details:</u></b><BR>" & strDetails & "&nbsp;"
				else
					Response.Write "<a href=""javascript:EditMilestone(" & rs("ID") & ");"">Details:</a></b><BR>" & strDetails & "&nbsp;"
				end if
				
				rs2.close
				
				rs2.open "spListRoadmapTasks " & rs("ID"),cn,adOpenForwardOnly
				if rs2.EOF and rs2.BOF then
					Response.Write "<BR><BR><b><u ID=TaskLink LANGUAGE=javascript onclick=""return ShowMenu(" & clng(request("ID")) & "," & rs("ID") & ");"" onmouseover=""return TaskLink_onmouseover();"" onmouseout=""return TaskLink_onmouseout();"">Tasks</u></b><BR>None."
				else
					Response.Write "<BR><BR><b><u ID=TaskLink LANGUAGE=javascript onclick=""return ShowMenu(" & clng(request("ID")) & "," & rs("ID") & ");"" onmouseover=""return TaskLink_onmouseover();"" onmouseout=""return TaskLink_onmouseout();"">Remaining Tasks</u></b><BR>"
					Response.Write "<TABLE cellpadding=2 cellspacing=0 bgcolor=ivory width=""90%""><TR bgcolor=beige><TH>ID</TH><TH>Summary</TH><TH width=150>&nbsp;&nbsp;Status</TH><TH width=150>Owner</TH></TR>"
					do while not rs2.EOF
						if rs2("InProgress") and trim(rs2("Status")) = "Open" then
							strStatus = "In&nbsp;Progress"
						elseif trim(rs2("Status")) = "Open" then
							strStatus = "Investigating"
						else
							strStatus = rs2("Status") & ""
						end if
						
						Response.Write "<TR LANGUAGE=javascript onclick=""return actionitem_onclick(" & rs2("ID") & "," & EditOK & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD valign=top>" & rs2("ID") & "&nbsp;&nbsp;</TD><TD>" & rs2("Summary") & "&nbsp;&nbsp;</TD><TD valign=top nowrap>&nbsp;&nbsp;" & strStatus & "&nbsp;&nbsp;</TD><TD valign=top nowrap>" & rs2("Owner") & "</TD></TR>"
						rs2.MoveNext
					loop
					Response.Write "</Table>"
				end if
				rs2.close

				'Closed Tasks
				if trim(strTimeframe) = "N/A" then
					rs2.open "spListRoadmapTasksClosed " & rs("ID") & ",'" &  Now()-30 & "'",cn,adOpenForwardOnly
				else
					rs2.open "spListRoadmapTasksClosed " & rs("ID"),cn,adOpenForwardOnly
				end if
				Response.Write "<BR><BR><b><u ID=TaskLink LANGUAGE=javascript onclick=""return ShowMenu(" & clng(request("ID")) & "," & rs("ID") & ");"" onmouseover=""return TaskLink_onmouseover();"" onmouseout=""return TaskLink_onmouseout();"">"
				if trim(strTimeframe) = "N/A" then
					response.write "Completed Tasks</u></b><font size=1 color=green> (Last 30 Days)</font><BR>"
				else
					response.write "Completed Tasks</u></b><BR>"
				end if
				if rs2.EOF and rs2.BOF then
					Response.Write "None."
				else
					Response.Write "<TABLE cellpadding=2 cellspacing=0 bgcolor=ivory width=""90%""><TR bgcolor=beige><TH>ID</TH><TH>Summary</TH><TH width=150>&nbsp;&nbsp;Date Completed</TH><TH width=150>Owner</TH></TR>"
					do while not rs2.EOF
						if isnull(rs2("ActualDate")) then
							if isnull(rs2("Created")) then
								strStatus = "-"
							else
								strStatus = formatdatetime(rs2("Created"),vbshortdate)
							end if
						else
							strStatus = formatdatetime(rs2("ActualDate"),vbshortdate)
						end if
						Response.Write "<TR LANGUAGE=javascript onclick=""return actionitem_onclick(" & rs2("ID") & "," & EditOK & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD valign=top>" & rs2("ID") & "&nbsp;&nbsp;</TD><TD>" & rs2("Summary") & "&nbsp;&nbsp;</TD><TD valign=top nowrap>&nbsp;&nbsp;" & strStatus & "&nbsp;&nbsp;</TD><TD valign=top nowrap>" & rs2("Owner") & "</TD></TR>"
						rs2.MoveNext
					loop
					Response.Write "</Table>"
				end if
				rs2.close


				rs2.open "spListRoadmapIssues " & rs("ID"),cn,adOpenForwardOnly
				if not (rs2.EOF and rs2.BOF) then
					Response.Write "<b><u><BR><BR>Open Issues</u></b><BR><TABLE cellpadding=2 cellspacing=0 bgcolor=ivory width=""90%""><TR bgcolor=beige><TH>ID</TH><TH>Summary</TH><TH>Owner</TH></TR>"
					do while not rs2.EOF
						Response.Write "<TR LANGUAGE=javascript onclick=""return actionitem_onclick(" & rs2("ID") & "," & EditOK & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD valign=top>" & rs2("ID") & "&nbsp;&nbsp;</TD><TD>" & rs2("Summary") & "&nbsp;&nbsp;</TD><TD valign=top nowrap>" & rs2("Owner") & "</TD></TR>"
						rs2.MoveNext
					loop
					Response.Write "</Table>"
				end if
				rs2.close
				set rs2=nothing

				Response.Write "</Blockquote></TD></TR>"
				rs.MoveNext
			loop

			set rs2 = server.CreateObject("ADODB.recordset")
			TaskCount = 0
			rs2.open "spCountRoadmapTasks4Product " & clng(request("ID")),cn,adOpenForwardOnly
			if not (rs2.eof and rs2.bof) then
				TaskCount = rs2("TaskCount")
			end if
			rs2.close
			
		'	if TaskCount > 0 then
			
				strOpenTasks = 0
				strClosedTasks = 0
				rs2.open "spGetActionRoadmapTaskCounts 0," & clng(request("ID")),cn,adOpenForwardOnly
				if not (rs.EOF and rs.BOF) then
					strOpenTasks = rs2("TaskCount") - rs2("CompleteCount") 
					strClosedTasks = rs2("CompleteCount") 
				end if
				rs2.Close
			
				Response.Write "<TR LANGUAGE=javascript onclick=""return item_onclick(0)"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD>Miscellaneous Tasks&nbsp;&nbsp;</TD><TD>" & strOpenTasks & "&nbsp;&nbsp;</TD><TD nowrap>" & strClosedTasks & "&nbsp;&nbsp;</TD><TD>N/A&nbsp;&nbsp;</TD><TD>N/A&nbsp;&nbsp;</TD><TD>Tasks not scheduled on the roadmap yet.&nbsp;&nbsp;</TD></TR>"
				Response.Write "<TR bgcolor=white style=""Display:none"" ID=Details0><TD colspan=6><Blockquote><b><u>Details:</u></b><BR>This is a ""bucket"" used to hold all tasks that are not scheduled on the roadmap yet.&nbsp;"
				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spListRoadmapTasks4product " & clng(request("ID")),cn,adOpenForwardOnly
				if rs2.EOF and rs2.BOF then
					Response.Write "<BR><BR><b><u ID=TaskLink LANGUAGE=javascript onclick=""return ShowMenu(" & clng(request("ID")) & ",0);"" onmouseover=""return TaskLink_onmouseover();"" onmouseout=""return TaskLink_onmouseout();"">Tasks</u></b><BR>None Defined."
				else
					Response.Write "<BR><BR><u><b><u ID=TaskLink LANGUAGE=javascript onclick=""return ShowMenu(" & clng(request("ID")) & ",0);"" onmouseover=""return TaskLink_onmouseover();"" onmouseout=""return TaskLink_onmouseout();"">Remaining Tasks</u></b><BR><TABLE cellpadding=2 cellspacing=0 bgcolor=ivory width=""90%""><TR bgcolor=beige><TH>ID</TH><TH>Summary</TH><TH>&nbsp;&nbsp;Status</TH><TH>Owner</TH></TR>"
					do while not rs2.EOF
						if rs2("InProgress") and trim(rs2("Status")) = "Open" then
							strStatus = "In&nbsp;Progress"
						else
							strStatus = rs2("Status") & ""
						end if	
						Response.Write "<TR LANGUAGE=javascript onclick=""return actionitem_onclick(" & rs2("ID") & "," & EditOK & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD valign=top>" & rs2("ID") & "&nbsp;&nbsp;</TD><TD>" & rs2("Summary") & "&nbsp;&nbsp;</TD><TD valign=top nowrap>&nbsp;&nbsp;" & strStatus & "&nbsp;&nbsp;</TD><TD valign=top nowrap>" & rs2("Owner") & "</TD></TR>"
						rs2.MoveNext
					loop
					Response.Write "</Table>"
				end if	
				rs2.close
				
				
				'Closed Misc Tasks
				rs2.open "spListRoadmapTasksClosed 0,'" &  Now()-30 & "'," & clng(request("ID")),cn,adOpenForwardOnly
				response.write "<BR><BR><b><u>Completed Tasks</u></u></b><font size=1 color=green> (Last 30 Days)</font><BR>"
				if rs2.EOF and rs2.BOF then
					Response.Write "None."
				else
					Response.Write "<TABLE cellpadding=2 cellspacing=0 bgcolor=ivory width=""90%""><TR bgcolor=beige><TH>ID</TH><TH>Summary</TH><TH width=150>&nbsp;&nbsp;Date Completed</TH><TH width=150>Owner</TH></TR>"
					do while not rs2.EOF
						if isnull(rs2("ActualDate")) then
							if isnull(rs2("Created")) then
								strStatus = "-"
							else
								strStatus = formatdatetime(rs2("Created"),vbshortdate)
							end if
						else
							strStatus = formatdatetime(rs2("ActualDate"),vbshortdate)
						end if
						Response.Write "<TR LANGUAGE=javascript onclick=""return actionitem_onclick(" & rs2("ID") & "," & EditOK & ")"" onmouseover=""return item_onmouseover()"" onmouseout=""return item_onmouseout()""><TD valign=top>" & rs2("ID") & "&nbsp;&nbsp;</TD><TD>" & rs2("Summary") & "&nbsp;&nbsp;</TD><TD valign=top nowrap>&nbsp;&nbsp;" & strStatus & "&nbsp;&nbsp;</TD><TD valign=top nowrap>" & rs2("Owner") & "</TD></TR>"
						rs2.MoveNext
					loop
					Response.Write "</Table>"
				end if
				rs2.close
				
				
				
				
			'end if			
			
			set rs2=nothing
			Response.Write "</table>"

			
		end if

		rs.Close
		
	
	
	end if


	set rs = nothing
	cn.Close
	set cn=nothing

	Response.Write "<font size=1 face=verdana><BR><BR><BR><BR><BR><BR>Report Generated: " & Now() & "</b></font>"

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function

%>


</BODY>
</HTML>

<%@ Language=VBScript %>
<%
    Dim AppRoot : AppRoot = Session("ApplicationRoot")
If Request("PddExport") Then
	Response.ContentType = "application/vnd.ms-excel"
Else
	  Response.Buffer = True
	  Response.ExpiresAbsolute = Now() - 1
	  Response.Expires = 0
	  Response.CacheControl = "no-cache"
End If
	'Response.ContentType = request("Format")
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="JavaScript" />
<title>Product Status - Confidential</title>
<style>
<!--
<!-- #include file="Style/programoffice.css" -->
//-->
</style>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

function PddExport()
{
    window.location.href = window.location + '&PddExport=True';
}

function window_onload() {
   dt = new Date();   //Gets today's date right now (to the millisecond).
        month = dt.getMonth() + 1;
   day = dt.getDate();
   year = dt.getFullYear();

	lblPleaseWait.style.display = "none";
	lblDateRange.style.display = "";
	
}

function ROW_onmouseover() {
        event.srcElement.style.cursor = "hand";
/*
	if (event.srcElement.className == "Row")
		event.srcElement.style.backgroundColor = "Thistle";
	else if (event.srcElement.parentElement.className == "Row")
		event.srcElement.parentElement.style.backgroundColor = "Thistle";
	else if (event.srcElement.parentElement.parentElement.className == "Row")
		event.srcElement.parentElement.parentElement.style.backgroundColor = "Thistle";
	else if (event.srcElement.parentElement.parentElement.parentElement.className == "Row")
		event.srcElement.parentElement.parentElement.parentElement.style.backgroundColor = "Thistle";
*/

	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

        if (srcElem.tagName != "TR") return;

        if (srcElem.className == "Row")
		srcElem.style.backgroundColor = "Thistle";


}

function ROW_onmouseout() {
	/*
	if (event.srcElement.className == "Row")
		event.srcElement.style.backgroundColor = "white";
	else if (event.srcElement.parentElement.className == "Row")
		event.srcElement.parentElement.style.backgroundColor = "white";
	else if (event.srcElement.parentElement.parentElement.className == "Row")
		event.srcElement.parentElement.parentElement.style.backgroundColor = "white";
	else if (event.srcElement.parentElement.parentElement.parentElement.className == "Row")
		event.srcElement.parentElement.parentElement.parentElement.style.backgroundColor = "white";
	*/
	
	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

        if (srcElem.tagName != "TR") return;

        if (srcElem.className == "Row")
		srcElem.style.backgroundColor = "White";
	
}


    function DCRROW_onclick(ID) {
	var strResult;
	strResult = window.open("Query/ActionReport.asp?ID=" + ID,"_blank","width=700, height=400,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes") 
	
}

    function OTSROW_onclick(ID) {
	var strResult;
	strResult = window.open("search/ots/Report.asp?txtReportSections=1&txtObservationID=" + ID,"_blank","width=700, height=400,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes") 

}

    function DelROW_onclick(ID, RootID) {
	var strResult;
	strResult = window.showModalDialog("WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + ID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 

}

//-->
</SCRIPT>
<script type="text/javascript">
var ns = (navigator.appName.indexOf("Netscape") != -1);
var d = document;
var px = document.layers ? "" : "px";
function JSFX_FloatDiv(id, sx, sy)
{
        var el = d.getElementById ? d.getElementById(id) : d.all ? d.all[id] : d.layers[id];
	window[id + "_obj"] = el;
        if (d.layers) el.style = el;
        el.cx = el.sx = sx; el.cy = el.sy = sy;
        el.sP = function (x, y) { this.style.left = x + px; this.style.top = y + px; };
        el.init = false;
	el.flt=function()
	{
		var pX, pY;
		pX = (this.sx >= 0) ? 0 : ns ? innerWidth : 
		document.documentElement && document.documentElement.clientWidth ? 
		document.documentElement.clientWidth : document.body.clientWidth;
		pY = ns ? pageYOffset : document.documentElement && document.documentElement.scrollTop ? 
		document.documentElement.scrollTop : document.body.scrollTop;
            if (this.sy < 0)
		pY += ns ? innerHeight : document.documentElement && document.documentElement.clientHeight ? 
		document.documentElement.clientHeight : document.body.clientHeight;
            this.cx += (pX + this.sx - this.cx) / 8; this.cy += (pY + this.sy - this.cy) / 8;
		if(!this.init)
		{
                this.init = true;
                this.cx = pX + this.sx;
                this.cy = pY + this.sy;
		}
		this.sP(this.cx, this.cy);
		setTimeout(this.id + "_obj.flt()", 0);
	}
	return el;
}
//JSFX_FloatDiv("divTopLeft",       10,   10).flt();
//JSFX_FloatDiv("divTopRight", 	  -100,   10).flt();
//JSFX_FloatDiv("divBottomLeft",    10, -100).flt();
//JSFX_FloatDiv("divBottomRight", -100, -100).flt();
</script>
</head>
<body LANGUAGE=javascript onload="return window_onload()">

<% If Not Request("PddExport") Then Response.Write "<div style=""position:absolute; right:20px;""><label style=""color:blue; text-decoration:underline;"" onClick=""PddExport();"" onMouseOver=""this.style.cursor='hand';"" onMouseOut=""this.style.cursor='';"">Excel Export</label></div>"%>
<p align=center>
<font face=verdana size=3>
<%

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
	
	Function GetWeekIndex(StartWeek, StartYear, GetWeek,GetYear)
		if StartYear = GetYear then
			GetWeekIndex = (GetWeek - StartWeek)
		else
			GetWeekIndex = GetWeek + (53-StartWeek) + (53* (GetYear - StartYear-1))
		end if
	end function

	dim cm
	dim cn
	dim p
	dim rs

  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout =120
	cn.ConnectionTimeout =120
	cn.Open

  'Create a recordset
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn
	dim CurrentUser
	dim CurrentUSerID
	dim CurrentUserPartner
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
	if rs.EOF and rs.BOF then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=0"	
	else
		CurrentUserID = rs("ID")
		CurrentUserPartner = rs("PartnerID")
	end if		
	rs.Close
	
  Dim PreRow
  Dim MidRow
  Dim PostRow
  Dim PreRowA
  Dim MidRowA
  Dim PostRowA
  dim PrimaryColor 
  dim strComments
  dim strMilestone
  dim strPOR
  dim strTarget
  dim strActual
  dim strproduct
  dim strLastMilestone
  dim rsCount   
  dim rowcount
  dim strExpired
  dim Milestonecount
  dim strDate
  dim strTargetDate
  dim strDays
  dim strCertificationStatus
  dim strSWQAStatus
  dim strPlatformStatus
  dim P1Width
  dim strPORDate
  dim strProductName
  dim strProductID
  dim SEPMID 
  dim ReportDays
  if request("ReportDays") = "" then
	  ReportDays = 14
  else
	  ReportDays = clng(request("ReportDays"))
  end if
  
  dim dtReportStart
  If Request("StartDt") = "" Then
	dtReportStart = "NULL"
  Else
	dtReportStart = "'" & Request("StartDt") & "'"
  End If
  
  dim dtReportEnd
  If Request("EndDt") = "" Then
	dtReportEnd = "NULL"
  Else
    dtReportEnd = "'" & Request("EndDt") & "'"
  End If

  
	if request("ID") <> "" then
		rs.Open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strproductName = ""
			strProductID = ""
			SEPMID =  ""
		else
			strproductName = rs("Name") & ""
			strProductID = request("ID")
			SEPMID = rs("SEPMID") & ""
			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					rs.Close
					set rs = nothing
					set cn=nothing
					
					Response.Redirect "../NoAccess.asp?Level=0"
				end if
			end if		
		end if
		rs.close
	elseif request("Product") <> "" then

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersionByName"
	
		Set p = cm.CreateParameter("@ID", 200, &H0001,255)
		p.Value = left(request("Product"),255)
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		'rs.Open "spGetProductVersionByName '" &  & "'",cn,adOpenForwardOnly
		strproductName = request("Product")
		strProductID = rs("ID") & ""
		SEPMID = rs("SEPMID") & ""
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				rs.Close
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../NoAccess.asp?Level=0"
			end if
		end if		
		rs.close
	else
		strproductName = ""
	end if
	
	if strProductName = "" or strproductID = "" then
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to find the selected product.</font>"
		set rs = nothing
		set cn = nothing
	else

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=white valign=top><TD><font size=1>"
  PreRowDone = "<TR bgcolor=ivory valign=top><TD ><FONT size=1>"
  MidRow = "</FONT></TD><TD valign=top ><FONT size=1>"
  MidRowCenter = "</FONT></TD><TD align=center valign=top ><FONT size=1>"
  MidRowCenterColSpan2 = "</FONT></TD><TD align=center valign=top colspan=2><FONT size=1>"
  
  PostRow = "</FONT></TD></TR>"
	if request("ReportTitle") = "" then
		Response.Write "<b>" & strproductName & " Status</b></font><BR><BR>"
	else
		Response.Write "<b>" & strproductName & " " & request("ReportTitle") & "</b></font><BR><BR>"
	end if

	If Request("PddExport")	Then
		Response.Write "<font face=verdana size=2><label id=lblDateRange>"
	Else
		Response.Write "<font face=verdana size=2><label id=lblPleaseWait>Preparing Report.  Please wait...</label><label id=lblDateRange style='display:none'>"
	End If
	
	If Request("StartDt") <> "" And Request("EndDt") <> "" Then
		Response.Write Request("StartDt") & " - " & Request("EndDt")
	Else
		Response.Write FormatDateTime(Now()-ReportDays,2) & " - " & FormatDateTime(Now(),2)
	End If
	Response.Write "</label></font>" 
	If Not Request("PddExport") Then
		'Response.Flush
	End If
%>
<!--</TD></TR></TABLE>--></P>

<%
	Dim SectionArray
	dim SectionCounter
	if request("Sections") = "" then
		SectionArray = split("1,2,3,22,4,5,6,7,8,9,10,11",",")
	else
		SectionArray = split(request("Sections"),",")
	end if
	
	for SectionCounter = lbound(SectionArray) to ubound(SectionArray)
		if isnumeric(SectionArray(SectionCounter)) then
			Select case SectionArray(SectionCounter)
				case 1
%>

<!-- Start Scope Section -->
	<BR><table ID=ScopeTable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<tr  bgcolor=Gainsboro> <TD  align=center><font bgcolor=black size=2><b>Scope</b></font></TD></TR>
<%
	rs.Open "SELECT BaseUnit, CurrentROM, OSSupport,IMagePO, ImageChanges, SystemBoardID, MachinePNPID, CommonImages, CertificationStatus, SWQAStatus, PlatformStatus, PDDReleased FROM ProductVersion v with (NOLOCK) WHERE v.id = " & clng(strproductID) & ";",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		dim strPHWebNames
		dim strStreetNames
		dim SeriesArray
		dim strSeriesText
		
		strPHWebNames = ""
		strStreetNames = ""
		
		
		set rs2 = server.CreateObject("ADODB.recordset")
		rs2.open "spListBrands4Product " & clng(strproductID),cn,adOpenForwardOnly
		do while not rs2.EOF
			if trim(rs2("SeriesSummary") & "") <> ""  then
				SeriesArray = split(rs2("SeriesSummary"),",")
				if trim(rs2("ProductVersion")) <> "" then
				if isnumeric(right(rs2("ProductVersion"),1) ) then
					strShortVersionName = left(rs2("ProductVersion"),len(rs2("ProductVersion"))-1)
				else
					strShortVersionName = left(rs2("ProductVersion"),len(rs2("ProductVersion"))-2)
				end if
				end if  
				for each strSeriesText in SeriesArray
					strStreetNames = strStreetNames & "<BR>" & rs2("StreetName2")  & " " & strSeriesText 
					strPHWebNames = strPHWebNames & "<BR>" & rs2("ProductFamily") & " " & rs2("Abbreviation") & " " & strShortVersionName & "X - " & rs2("StreetName")  & " " & strSeriesText
				next
			end if	
			rs2.MoveNext
		loop		
		rs2.Close
		set rs2 = nothing
		if trim(strStreetNames) = "" then
			strStreetNames = "TBD"
		end if
		if trim(strPHWebNames) = "" then
			strPHWebNames = "TBD"
		end if
		
		
%>
		<% response.write prerow & "<font size=1 face=verdana><b>Base Unit Definition:</b><br>" & replace(rs("BaseUnit")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Current ROM Versions:</b><br>" & replace(rs("CurrentROM")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Operating Systems:</b><br>" & replace(rs("OSSupport")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>PHWeb Names:</b> " & strPHWebNames & "&nbsp;</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Street Names:</b> " & strStreetNames & "&nbsp;</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>System Board ID:</b> " & replace(rs("SystemBoardID")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Machine PnP ID:</b> " & replace(rs("MachinePnPID")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Image PO:</b><br>" & replace(rs("ImagePO")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Image Changes:</b><br>" & replace(rs("ImageChanges")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
		<% response.write prerow & "<font size=1 face=verdana><b>Common Image Support:</b><br>" & replace(rs("CommonImages")& "",vbcrlf,"<BR>") & "</font>" & postrow %>
<%
	strCertificationStatus = rs("CertificationStatus")
	strSWQAStatus =  rs("SWQAStatus")
	strPlatformStatus =  rs("PlatformStatus")
	strPORDate = rs("PDDReleased") & ""
	end if
%>	</TABLE>

<%
	rs.Close
		case 2
%>	
	
<!-- Start Notes Section -->
	<BR>
	<table ID=NotesTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgcolor=Gainsboro> <TD align=center><font color=black size=2><b>Status / Accomplishments</b></font></TD></TR>
<%
	rs.Open "spListActionItems4Status " & clng(strproductID) & ",4" & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white  valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		Response.Write "<TR><TD>"
		do while not rs.EOF	

			'if isnull(rs("Updated")) then
			'	Response.Write "<TR><TD>"
			'elseif DateDiff("d",rs("Updated"),Now()) < 0 then
			'	Response.Write "<TR bgcolor=ivory><TD>"
			'else
			'	Response.Write "<TR><TD>"
			'end if

			if trim(rs("Status") & "") <> "" then
				select case  rs("Status")
				case 1
					strStatus = "Open"
				case 2
					strStatus = "Closed"
				case 3
					strStatus = "Need More Information"
				case 4
					strStatus = "Approved"
				case 5
					strStatus = "Disapproved"
				case 6
					strStatus = "Investigating"
				end select 
			else
				strDesc = ""
			end if


			if trim(SEPMID) = trim(CurrentUserID) then
				Response.Write "<TR><TD><Table width=100% cellspacing=0 cellpadding=1 border=1 bordercolor=gainsboro>"
				Response.Write "<TR><TD><font size=1 face=verdana><b>" & rs("Summary") & "</B></font></TD></TR>"
				Response.Write "<TR><TD><font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"") & "</font></td></tr></table>"
				Response.Write "</TD></TR>"
			else
				Response.Write "<p><font size=2 face=verdana><b>" & rs("Summary") & "</B></font></P>"
				Response.Write "<BLOCKQUOTE dir=ltr style=""MARGIN-RIGHT: 0px"">"
				Response.Write "<font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"") & "</font></BLOCKQUOTE><BR>"
			end if
			rs.MoveNext
		loop
		if trim(SEPMID) = trim(CurrentUserID) then
			Response.Write "</TD></TR>"
		end if
	 %>

<%
	end if
%>	</TABLE><%
	rs.Close
	
	
	

	
	
		case 3
	%>
<!-- Start Scope Change Section -->
	<BR>
	<table ID=ChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgColor=Gainsboro> <TD align=center><font color=black size=2><b>Change Requests (DCR)</b></font></TD></TR>
<%
	rs.Open "spListActionItems4Status " & clng(strproductID) & ",3" & ", " & dtReportStart & ", " & dtReportEnd & ", 0",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		do while not rs.EOF	

'			Response.Write "<FONT face=verdana size=2><b>ID#: <a href=""javascript:DisplayAction(" & rs("ID") & "," & rs("Type") & ");"">" & rs("ID") & "</a></b></font><BR>"
'			Response.Write "<FONT face=verdana size=1><b>Summary: " & rs("Summary") & "</b></font><BR>"
'			Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"">"

			strResolution = trim(rs("Resolution") & "")
			if isnull(rs("ActualDate")) then
				Response.Write "<TR><TD>"
			elseif rs("Status") & "" = "1" or rs("Status") & "" = "3" or rs("Status") & "" = "6" then
				Response.Write "<TR><TD>"
			else
				Response.Write "<TR bgcolor=ivory><TD>"
			end if


			if trim(rs("Status") & "") <> "" then
				select case  rs("Status")
				case 1
					strStatus = "Proposed"
				case 2
					strStatus = "Closed"
				case 3
					strStatus = "Need More Information"
				case 4
					strStatus = "Approved"
				case 5
					strStatus = "Disapproved"
				case 6
					strStatus = "Investigating"
				end select 
			else
				strDesc = ""
			end if


			Response.Write "<Table width=100% cellspacing=0 cellpadding=1 border=1 bordercolor=gainsboro>"
			Response.Write "<TR><TD colspan=3><font size=1 face=verdana><b>" & rs("Summary") & "</B></font></TD></TR>"

			Response.Write "<TR><TD><TABLE  width=""100%""><TR><TD nowrap><font face=verdana size=1>ID:</font></TD><TD><font face=verdana size=1>" & rs("ID") & "</font></TD></TR><TR><TD><font face=verdana size=1>Product:</font></TD><TD><font face=verdana size=1>" & rs("Product") & "</font></TD></TR><TR><TD><font face=verdana size=1>Status:</font><TD><font face=verdana size=1>" & strStatus & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Date Created:</font></TD><TD><font face=verdana size=1>" & rs("Created") & "</font></TD></TR><TR><TD><font face=verdana size=1>Days Open:</font></TD><TD><font face=verdana size=1>" & DateDiff("d",rs("Created"),Date()) & "</font></TD></TR><TR><TD><font face=verdana size=1>Target Date:</font><TD><font face=verdana size=1>" & rs("TargetDate") & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Submitter:</font></TD><TD><font face=verdana size=1>" & rs("Submitter") & "</font></TD></TR><TR><TD><font face=verdana size=1>Owner:</font></TD><TD><font face=verdana size=1>" & rs("Owner") & "</font></TD></TR><TR><TD><font face=verdana size=1>Core Team Rep:</font><TD><font face=verdana size=1>" & rs("CoreTeamRep") & "</font></TD></TR></table></TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>Author</u><BR>" & rs("AuthorFullname") & "<BR>" & rs("AuthorGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Module PM</u><BR>" & rs("PMFullName") & "<BR>" & rs("PMGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Developer</u><BR>" & rs("DeveloperFullName") & "<BR>" & rs("DeveloperGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Owner</u><BR>" & rs("FullName") & "<BR>" & rs("OwnerGroup") & "</font></TD></tr></table>   </TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>System Build</u><BR>" & rs("Systemboardrev") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>System ROM</u><BR>" & rs("SystemROM") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>OS</u><BR>" & rs("OSRelease") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Release</u><BR>" & strImage & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Language</u><BR>" & strImageLanguage & "</font></TD></tr></table></TD></TR>"
			Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"<BR>") & "</font></td></tr></table></TD></TR>"
			if trim(rs("Approvals") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Approvals: </font></td><td><font size=1 face=verdana>" & replace(rs("Approvals"),vbcrlf,"<BR>")  &"</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Justification") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Justification: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Actions") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Actions: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Actions") & "",vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(strResolution) <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Resolution: </font></td><td><font size=1 face=verdana>" & replace(replace(strResolution,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") &"</font></td></tr></table></TD></TR>"
			end if
			Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "</TD></TR>"


			
			rs.MoveNext
		loop

	end if
%>	</TABLE><%
	rs.Close
		case 4
	%>
<!-- Start Issue Section -->
	<BR>
	<table ID=ChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgColor=Gainsboro> <TD align=center><font color=black size=2><b>Issues/Risks</b></font></TD></TR>
<%
	rs.Open "spListActionItems4Status " & clng(strproductID) & ",1" & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white valign=top><TD  height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		do while not rs.EOF	

'			Response.Write "<FONT face=verdana size=2><b>ID#: <a href=""javascript:DisplayAction(" & rs("ID") & "," & rs("Type") & ");"">" & rs("ID") & "</a></b></font><BR>"
'			Response.Write "<FONT face=verdana size=1><b>Summary: " & rs("Summary") & "</b></font><BR>"
'			Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"">"
			if isnull(rs("ActualDate")) then
				Response.Write "<TR><TD>"
			elseif rs("Status") & "" = "1" or rs("Status") & "" = "3" or rs("Status") & "" = "6" then
				Response.Write "<TR><TD>"
			else
				Response.Write "<TR bgcolor=ivory><TD>"
			end if


			if trim(rs("Status") & "") <> "" then
				select case  rs("Status")
				case 1
					strStatus = "Open"
				case 2
					strStatus = "Closed"
				case 3
					strStatus = "Need More Information"
				case 4
					strStatus = "Approved"
				case 5
					strStatus = "Disapproved"
				case 6
					strStatus = "Investigating"
				end select 
			else
				strDesc = ""
			end if

			strResolution = trim(rs("Resolution") & "")

			Response.Write "<Table width=100% cellspacing=0 cellpadding=1 border=1 bordercolor=gainsboro>"
			Response.Write "<TR><TD colspan=3><font size=1 face=verdana><b>" & rs("Summary") & "</B></font></TD></TR>"

			Response.Write "<TR><TD><TABLE  width=""100%""><TR><TD nowrap><font face=verdana size=1>ID:</font></TD><TD><font face=verdana size=1>" & rs("ID") & "</font></TD></TR><TR><TD><font face=verdana size=1>Product:</font></TD><TD><font face=verdana size=1>" & rs("Product") & "</font></TD></TR><TR><TD><font face=verdana size=1>Status:</font><TD><font face=verdana size=1>" & strStatus & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Date Created:</font></TD><TD><font face=verdana size=1>" & rs("Created") & "</font></TD></TR><TR><TD><font face=verdana size=1>Days Open:</font></TD><TD><font face=verdana size=1>" & DateDiff("d",rs("Created"),Date()) & "</font></TD></TR><TR><TD><font face=verdana size=1>Target Date:</font><TD><font face=verdana size=1>" & rs("TargetDate") & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Submitter:</font></TD><TD><font face=verdana size=1>" & rs("Submitter") & "</font></TD></TR><TR><TD><font face=verdana size=1>Owner:</font></TD><TD><font face=verdana size=1>" & rs("Owner") & "</font></TD></TR><TR><TD><font face=verdana size=1>Core Team Rep:</font><TD><font face=verdana size=1>" & rs("CoreTeamRep") & "</font></TD></TR></table></TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>Author</u><BR>" & rs("AuthorFullname") & "<BR>" & rs("AuthorGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Module PM</u><BR>" & rs("PMFullName") & "<BR>" & rs("PMGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Developer</u><BR>" & rs("DeveloperFullName") & "<BR>" & rs("DeveloperGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Owner</u><BR>" & rs("FullName") & "<BR>" & rs("OwnerGroup") & "</font></TD></tr></table>   </TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>System Build</u><BR>" & rs("Systemboardrev") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>System ROM</u><BR>" & rs("SystemROM") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>OS</u><BR>" & rs("OSRelease") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Release</u><BR>" & strImage & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Language</u><BR>" & strImageLanguage & "</font></TD></tr></table></TD></TR>"
			Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"<BR>") &"</font></td></tr></table></TD></TR>"
			if trim(rs("Approvals") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Approvals: </font></td><td><font size=1 face=verdana>" & replace(rs("Approvals"),vbcrlf,"<BR>")  &"</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Justification") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Justification: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Actions") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Actions: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Actions") & "",vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(strResolution) <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Resolution: </font></td><td><font size=1 face=verdana>" & replace(replace(strResolution,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") &"</font></td></tr></table></TD></TR>"
			end if
			Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "</TD></TR>"


			
			rs.MoveNext
		loop
	 %>

<%
	end if
%>	</TABLE><%
	rs.Close
	
		Case 5
	%>
<!-- Start Action Item Section -->
	<BR>
	<table ID=ChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgColor=Gainsboro> <TD align=center><font color=black size=2><b>Action Items</b></font></TD></TR>
<%
	rs.Open "spListActionItems4Status " & clng(strproductID) & ",2" & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		do while not rs.EOF	

'			Response.Write "<FONT face=verdana size=2><b>ID#: <a href=""javascript:DisplayAction(" & rs("ID") & "," & rs("Type") & ");"">" & rs("ID") & "</a></b></font><BR>"
'			Response.Write "<FONT face=verdana size=1><b>Summary: " & rs("Summary") & "</b></font><BR>"
'			Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"">"
			if isnull(rs("ActualDate")) then
				Response.Write "<TR><TD>"
			elseif rs("Status") & "" = "1" or rs("Status") & "" = "3" or rs("Status") & "" = "6" then
				Response.Write "<TR><TD>"
			else
				Response.Write "<TR bgcolor=ivory><TD>"
			end if


			if trim(rs("Status") & "") <> "" then
				select case  rs("Status")
				case 1
					strStatus = "Open"
				case 2
					strStatus = "Closed"
				case 3
					strStatus = "Need More Information"
				case 4
					strStatus = "Approved"
				case 5
					strStatus = "Disapproved"
				case 6
					strStatus = "Investigating"
				end select 
			else
				strDesc = ""
			end if

			strResolution = trim(rs("Resolution") & "")

			Response.Write "<Table width=100% cellspacing=0 cellpadding=1 border=1 bordercolor=gainsboro>"
			Response.Write "<TR><TD colspan=3><font size=1 face=verdana><b>" & rs("Summary") & "</B></font></TD></TR>"

			Response.Write "<TR><TD><TABLE  width=""100%""><TR><TD nowrap><font face=verdana size=1>ID:</font></TD><TD><font face=verdana size=1>" & rs("ID") & "</font></TD></TR><TR><TD><font face=verdana size=1>Product:</font></TD><TD><font face=verdana size=1>" & rs("Product") & "</font></TD></TR><TR><TD><font face=verdana size=1>Status:</font><TD><font face=verdana size=1>" & strStatus & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Date Created:</font></TD><TD><font face=verdana size=1>" & rs("Created") & "</font></TD></TR><TR><TD><font face=verdana size=1>Days Open:</font></TD><TD><font face=verdana size=1>" & DateDiff("d",rs("Created"),Date()) & "</font></TD></TR><TR><TD><font face=verdana size=1>Target Date:</font><TD><font face=verdana size=1>" & rs("TargetDate") & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Submitter:</font></TD><TD><font face=verdana size=1>" & rs("Submitter") & "</font></TD></TR><TR><TD><font face=verdana size=1>Owner:</font></TD><TD><font face=verdana size=1>" & rs("Owner") & "</font></TD></TR><TR><TD><font face=verdana size=1>Core Team Rep:</font><TD><font face=verdana size=1>" & rs("CoreTeamRep") & "</font></TD></TR></table></TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>Author</u><BR>" & rs("AuthorFullname") & "<BR>" & rs("AuthorGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Module PM</u><BR>" & rs("PMFullName") & "<BR>" & rs("PMGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Developer</u><BR>" & rs("DeveloperFullName") & "<BR>" & rs("DeveloperGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Owner</u><BR>" & rs("FullName") & "<BR>" & rs("OwnerGroup") & "</font></TD></tr></table>   </TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>System Build</u><BR>" & rs("Systemboardrev") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>System ROM</u><BR>" & rs("SystemROM") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>OS</u><BR>" & rs("OSRelease") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Release</u><BR>" & strImage & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Language</u><BR>" & strImageLanguage & "</font></TD></tr></table></TD></TR>"
			Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"<BR>") &"</font></td></tr></table></TD></TR>"
			if trim(rs("Approvals") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Approvals: </font></td><td><font size=1 face=verdana>" & replace(rs("Approvals"),vbcrlf,"<BR>")  &"</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Justification") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Justification: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Actions") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Actions: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Actions") & "",vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(strResolution) <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Resolution: </font></td><td><font size=1 face=verdana>" & replace(replace(strResolution,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") &"</font></td></tr></table></TD></TR>"
			end if
			Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "</TD></TR>"


			
			rs.MoveNext
		loop
	 %>

<%
	end if
%>	</TABLE><%
	rs.Close
		case 6
	%>



<!--Deliverable Matrix Section-->

<%
  rs.Open "spListDeliverableMatrixUpdates " & clng(strproductID) & ",2," & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=DeliverableTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>Deliverable Matrix Changes</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=DeliverableTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 3 align=center><font color=black size=2><b>Deliverable Matrix Changes</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Change</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Deliverable</font></strong></td>
	<td nowrap><strong><font color=black size=1>Date</font></strong></td></tr>
<%  
  'Display requirements
  do while not rs.EOF
	strDate = rs("Updated")& ""
	if strDate = "" then
		strDate = "&nbsp;"
	end if

	strDeliverable = rs("Deliverable") & " " & rs("Version")
	if rs("Revision") & "" <> "" then
		strDeliverable = strDeliverable & "," & rs("Revision")
	end if
	if rs("Pass") & "" <> "" then
		strDeliverable = strDeliverable & "," & rs("Pass")
	end if
	
	if rs("Action") & "" = "Targeted" then
		strAction = "Added"		
	elseif rs("Action") & "" = "Target Removed" then
		strAction = "Removed"		
	else
		strAction = rs("Action") & "&nbsp;"
	end if
	
	if datediff("d",rs("Updated"),Now()-7)<0 then
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return DelROW_onclick('" & rs("Versionid") & "','" & rs("RootID") & "')""><TD class=static><font size=1>"
	else
		Response.Write "<TR bgcolor=ivory valign=top class=""Row"" LANGUAGE=javascript onclick=""return DelROW_onclick('" & rs("Versionid") & "','" & rs("RootID") & "')""><TD class=static><font size=1>"
	end if
%>
		<font size=1 face=verdana><%=strAction%> </FONT><%= midrow %><FONT size=1><%= strDeliverable%></FONT><%= midrow %><FONT size=1><%= rs("Updated")%> </font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if
 
	case 7
%>




<!-- Start System Ceritication Section -->
	<BR>
	<table ID=ChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgColor=Gainsboro> <TD colspan=1 align=center><font color=black size=2><b>System Certification Status</b></font></TD></TR>
		
<%
	response.write prerow & replace(strCertificationStatus & "&nbsp;",vbcrlf,"<BR>")  & postrow 
%>	</TABLE>

<%
	case 8
%>

<!-- Start SWQA Section -->
	<BR>
	<table ID=SWQATable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<tr  bgcolor=Gainsboro> <TD  align=center><font bgcolor=black size=2><b>Software QA Test Status</b></font></TD></TR>
<%
	
	response.write prerow & replace(strSWQAStatus & "&nbsp;",vbcrlf,"<BR>")  & postrow 
%>	</TABLE>

<% Case 9%>
<!-- Start Platform Status Section -->
	<BR>
	<table ID=PlatformTable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<tr  bgcolor=Gainsboro> <TD  align=center><font bgcolor=black size=2><b>Platform Validation Test Status</b></font></TD></TR>
<%
	response.write prerow & replace(strPlatformStatus & "&nbsp;",vbcrlf,"<BR>") & "</font>" & postrow 
%>	</TABLE>

<!--Schedule Section-->

<%	

	case 10

		rs.Open "usp_SelectProductReportScheduleData " & clng(strproductID),cn,adOpenStatic

  'Initializae the cell backgroud color selection
  primarycolor = false
  
  'Start Table
  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=ScheduleTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD colspan = 6 align=center><font color=Black size=2><b>Schedule</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD colspan = 6 height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE><BR><BR>"	
		rs.Close
  else

  'Display requirements
  LastMilestone = ""
  MilestoneCount = 0
  LastSchedule = ""
  NotFirstSchedule = False
    
  do while not rs.EOF
	If LastSchedule <> rs("schedule_id") Then
		LastSchedule = rs("schedule_id")
		MilestoneCount = 0
		
		If NotFirstSchedule Then
			Response.Write "</TBODY></table>"
		End If
		
		NotFirstSchedule = True
%>
  <BR>
	<table ID=ScheduleTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD colspan = 10 align=center><font color=black size=2><b><%= rs("schedule_name")%>&nbsp;Schedule</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=50 rowspan=2><strong><font color=black size=1>ID#</font></strong></td>	  
    <td nowrap width="45%" rowspan=2><strong><font color=black size=1>Item</font></strong></td>
	<td align=center nowrap width=140 colspan=2><font color=black size=1><strong>POR</strong></td>
	<td align=center nowrap width=140 colspan=2><font color=black size=1><strong>Projected</strong></td>
	<td align=center nowrap width=140 colspan=2><font color=black size=1><strong>Actual</strong></td>
	<td width="100%" rowspan=2><font color=black size=1><strong>Comments</strong></td></tr>
	<tr bgColor=Gainsboro>
	<td align=center nowrap width=70><font color=black size=1><strong>Start</strong></td>
	<td align=center nowrap width=70><font color=black size=1><strong>Finish</strong></td>
	<td align=center nowrap width=70><font color=black size=1><strong>Start</strong></td>
	<td align=center nowrap width=70><font color=black size=1><strong>Finish</strong></td>
	<td align=center nowrap width=70><font color=black size=1><strong>Start</strong></td>
	<td align=center nowrap width=70><font color=black size=1><strong>Finish</strong></td></tr>
<%  
	End If
	Milestonecount = Milestonecount + 1
	primarycolor = not primarycolor
	strComments = replace(replace(replace(rs("item_notes") & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>") & "&nbsp;"
	strMilestone = replace(replace(replace(rs("item_description") & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>")& "&nbsp;"
	strPORStart = rs("por_start_dt") & "&nbsp;"
	strPorEnd = rs("por_end_dt") & "&nbsp;"
	strActualStart = rs("actual_start_dt") & "&nbsp;"
	strActualEnd = rs("actual_end_dt") & "&nbsp;"
	
	if isnull(rs("actual_start_dt")) then
		strTargetStart = rs("projected_start_dt") & ""
	else
		strTargetStart = "&nbsp;"
	end if
	
	if isnull(rs("actual_end_dt")) then
		strTargetEnd = rs("projected_end_dt") & ""
	else
		strTargetEnd = "&nbsp;"
	end if
	
	
	strexpired = ""
	if stractualend = "&nbsp;" then
		if isdate(strtargetend)then
			if datediff("d",strtargetend,now)>= 0 then
				strExpired = " <IMG SRC=""images/alert.gif"">"
			end if
		end if
	end if
	
	strTargetEnd = strTargetEnd & "&nbsp;"

	if lastPhase <> rs("item_phase") then
		Response.Write "<TR bgcolor=Gainsboro valign=top><TD colspan=10><font size=1 face=verdana><b>" & rs("phase_name") & "</font></b></TD></TR>"
		lastphase = rs("item_phase")
	end if

	strproduct = ""
	if trim(strActualEnd) = "&nbsp;" then
		Response.Write prerow
	else
		Response.Write prerowdone
	end if

	Response.Write MilestoneCount & " " & strexpired
	Response.Write midrow
	Response.Write strMilestone

	if UCase(rs("milestone_yn")) = "Y" then
		Response.Write midRowCenterColSpan2
		Response.Write strPorStart & "&nbsp;"
		Response.Write midrowcentercolspan2
		Response.Write strTargetStart & "&nbsp;"
		Response.Write midrowcentercolspan2
		Response.Write strActualStart & "&nbsp;"
	else
		Response.Write midRowCenter
		Response.Write strPorStart & "&nbsp;"
		Response.Write midrowcenter
		Response.Write strPorEnd & "&nbsp;"
		Response.Write midrowcenter
		Response.Write strTargetedStart & "&nbsp;"
		Response.Write midrowcenter
		Response.Write strTargetedEnd & "&nbsp;"
		Response.Write midrowcenter
		Response.Write strActualStart & "&nbsp;"
		Response.Write midrowcenter
		Response.Write strActualEnd & "&nbsp;"
	end if
	Response.Write midrow
	Response.Write strComments & "&nbsp;"
	Response.Write postrow
		
	rs.MoveNext
  loop
  
  if rs.EOF and rs.BOF then
	Response.Write "</TBODY></table><FONT face=verdana size = 2><strong>No Schedule is defined for this product.</strong></font>"	
  else
   'Finish off table 
    Response.Write "</TBODY></table>"
  end if
  
  'Cleanup
  rs.Close
  Response.Write "</TBODY></table>"
  
  end if
  
  case 11
%>
	
	
<!-- OTS Summary Section -->

<%

dim strOTSName
dim strOTSCycle
dim strOTS
dim blnOTSLinked
rs.Open "SELECT DOTSName, CYCLE FROM ProductVersion with (NOLOCK) WHERE ID = " & clng(strproductID), cn, adOpenForwardOnly
If rs.EOF And rs.EOF Then
	strOTS = "<font size=1 color=red face=verdana>OTS Unavailable.</font>"
	blnOTSLinked = false
ElseIf rs("DOTSName") & "" = "" Or rs("cycle") & "" = "" Then
	strOTS = "<font size=1 color=red face=verdana>OTS Link Not Defined.</font>"
	blnOTSLinked = false
else
	strOTS = ""
	strOTSName = rs("DOTSName") & ""
	strOTSCycle = rs("Cycle") & ""
	blnOTSLinked = true
end if
rs.Close

%>

	<BR>
	<table ID=PlatformTable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<tr  bgcolor=Gainsboro> <TD colspan=4 align=center><font bgcolor=black size=2><b>OTS Summary</b></font></TD></TR>
<%
	if strOTS <> "" then	
		response.write "<TR bgcolor=white valign=top><TD colspan=4>" & replace(strOTS & "&nbsp;",vbcrlf,"<BR>") & "</font>" & postrow 
	else
		dim strAllOpen
		dim strAllClosed
		dim strAllOpenP1
		dim strAllClosedP1
		dim strP1UI
		dim strP1FixVerified
		dim strP1FixinProgress
		dim strUnassigned
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.commandtimeout=120
		cm.CommandText = "spOTSStatus"

		Set p = cm.CreateParameter("@ID", 200, &H0001,20)
		p.Value = left(strOTSName,20)
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		'rs.Open "spOTSStatus '" & strOTSName & "'",cn,adOpenForwardOnly
		strAllOpen = rs("AllOpen") & ""
		strAllClosed = rs("AllClosed") & ""
		strAllOpenP1 = rs("AllOpenP1") & ""
		strAllClosedP1 = rs("AllClosedP1") & ""
		'strP1FixVerified  = rs("P1FixVerified") & ""
		'strP1FixinProgress  = rs("P1FixInProgress") & ""
		'strP1UI  = rs("P1UI") & ""
		strUnassigned = rs("Unassigned") & ""
		rs.Close

		Response.Write "<TR bgcolor=Gainsboro valign=top><TD colspan=4><font size=1 face=verdana><b>Program</font></b></TD></TR>"
		Response.Write "<TR><TD><font size=1>&nbsp;</font></TD><TD width=90><b><font size=1>All</font></b></b></TD><TD width=90><b><font size=1>Open</font></b></TD><TD width=90><b><font size=1>Closed</font></b></TD></TR>"
		Response.Write "<TR><TD><b><font size=1>All Priorities</font></b></TD><TD><font size=1>" & clng(strAllOpen) + clng(strAllClosed) & "</font></TD><TD><font size=1>" & strAllOpen & "</font></TD><TD><font size=1>" & strAllClosed & "</font></TD></TR>"
		Response.Write "<TR><TD><b><font size=1>Priority 1</font></b></TD><TD><font size=1>" & clng(strAllOpenP1) + clng(strAllClosedP1) & "</font></TD><TD><font size=1>" & strAllOpenP1 & "</font></TD><TD><font size=1>" & strAllClosedP1 & "</font></TD></TR>"

		Response.Write "<TR bgcolor=Gainsboro valign=top><TD colspan=4><font size=1 face=verdana><b>Status (P1 only)</font></b></TD></TR>"
		Response.Write "<TR><TD><font size=1>&nbsp;</font></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD><TD><b><font size=1>Open</font></b></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD></TR>"
		Response.Write "<TR><TD><b><font size=1>Unassigned</font></b></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD><TD><font size=1>" & strUnassigned & "</font></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD></TR>"
		
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetOTSStates"

		Set p = cm.CreateParameter("@Product", 200, &H0001,20)
		p.Value = left(strOTSName,20)
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		do while not rs.eof
			Response.Write "<TR><TD><b><font size=1>" & rs("State") & "" & "</font></b></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD><TD><font size=1>" & rs("StateCount") & "" & "</font></TD><TD bgcolor=Gainsboro><font size=1>&nbsp;</font></TD></TR>"
			rs.movenext
		loop
		rs.close
		
		Response.Write "<TR bgcolor=Gainsboro valign=top><TD colspan=4><font size=1 face=verdana><b>Category (P1 only)</font></b></TD></TR>"
		Response.Write "<TR><TD><font size=1>&nbsp;</font></TD><TD><b><font size=1>All</font></b></TD><TD><font size=1><b>Open</b></font></TD><TD><font size=1><b>Investigating</b></font></TD></TR>"
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetOTSStatus4Category"

		Set p = cm.CreateParameter("@ID", 200, &H0001,20)
		p.Value = left(strOTSName,20)
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
		
		
		dim strCatOpen
		dim strCatClosed
		dim strCatInvest 

		dim strCat
		dim CatRow
		strCat = ""
		CatRow = ""
		strCatOpen = 0
		strCatClosed = 0
		strCatInvest = 0
		do while not rs.EOF
			if strcat <> rs("Category") & ""  and strCat <> "" then
				Response.Write "<TR><TD><b><font size=1>" & strCat & "</font></b></TD><TD><font size=1>" & clng(strCatOpen) + clng(strCatClosed) & "</font></TD><TD><font size=1>" & strcatOpen & "</font></TD><TD><font size=1>" & strcatInvest & "</font></TD></TR>"	
				strCatOpen = 0
				strCatClosed = 0
				strCatInvest = 0
			end if
			strCat = rs("Category") & ""
			if rs("Status") = "1" then
				strcatOpen = rs("OTSCount")
			elseif rs("Status") = "2" then
				strcatClosed = rs("OTSCount")
			else
				strcatInvest = rs("OTSCount")
			end if
			rs.MoveNext
		loop
		rs.Close
		Response.Write "<TR><TD><b><font size=1>" & strCat & "</font></b></TD><TD><font size=1>" & clng(strCatOpen) + clng(strCatClosed) & "</font></TD><TD><font size=1>" & strcatOpen & "</font></TD><TD><font size=1>" & strcatInvest & "</font></TD></TR>"	
	end if
	if (clng(strAllOpen)+clng(strAllClosed)) = 0 then
		p1Width = 100
	else
		if strAllOpen="" and strAllClosed=""  then
			P1Width = 100
		elseif ((clng(strAllOpenP1)+clng(strAllClosedP1))/(clng(strAllOpen)+clng(strAllClosed))*100) > 100 then
			P1Width=100
		else
			P1Width = ((clng(strAllOpenP1)+clng(strAllClosedP1))/(clng(strAllOpen)+clng(strAllClosed))*100)
		end if
	end if	
%>	</TABLE>

<BR>

<!--
<%if not (strAllOpen = 0 and strAllClosed = 0) then%>
<TABLE border=1 width="100%" cellpadding=0 cellspacing=0 bordercolor=white>
	<TR>
		<TD width=400><font size=1>All&nbsp;Observations:</TD>
		<%
		if strAllOpen > 0 then
			strOpenWidth = (clng(strAllOpen)/(clng(strAllOpen) + clng(strAllClosed))) * 100 
			Response.Write "<TD bgcolor=pink width=""" & strOpenWidth  &  "%""><font size=1 color=black><b>" & strAllOpen & "&nbsp;Open</b></font></TD>" 
		end if
		if strAllClosed > 0 then
			strClosedWidth = (clng(strAllClosed)/(clng(strAllOpen) + clng(strAllClosed))) * 100 
			Response.Write "<TD bgcolor=green width=""" & strClosedWidth  &  "%""><font color=white size=1><b>" & strAllClosed & "&nbsp;Closed</b></font></TD>" 
		end if
		%>
	</TR>
</Table>
<%if not (strAllOpenP1 = 0 and strAllClosedP1 = 0) then%>
<TABLE border=1 width="<%=P1Width%>%" cellpadding=0 cellspacing=0 bordercolor=white>
	<TR>
		<TD width=400><font size=1>P1:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
		<%
		if strAllOpenP1 > 0 then
			strOpenWidth = (clng(strAllOpenP1)/(clng(strAllOpenP1) + clng(strAllClosedP1))) * 100 
			Response.Write "<TD bgcolor=pink width=""" & strOpenWidth  &  "%""><font size=1 color=black><b>" & strAllOpenP1 & "&nbsp;Open</b></font></TD>" 
		end if
		if strAllClosedP1 > 0 then
			strClosedWidth = (clng(strAllClosedP1)/(clng(strAllOpenP1) + clng(strAllClosedP1))) * 100 
			Response.Write "<TD bgcolor=green width=""" & strClosedWidth  &  "%""><font color=white size=1><b>" & strAllClosedP1 & "&nbsp;Closed</b></font></TD>" 
		end if
		%>
	</TR>
</TABLE>
<%end if%>

<%end if%>
-->


<!--
<TABLE border=1 width=<%=((clng(strAllOpen) + clng(strAllClosed) * 7)+200)%> cellpadding=0 cellspacing=0 bordercolor=white>
	<TR width = <%=((clng(strAllOpen) + clng(strAllClosed)) * 7)+200%>>
		<TD width=200><font size=1>All Observations:</TD>
		<%
		strOpenWidth = clng(strAllOpen) * 7 '/ (clng(strAllOpen) + clng(strAllClosed)) * 100)
		Response.Write "<TD bgcolor=red width=" & strOpenWidth  &  "><font size=1 color=white><b>" & clng(strOpenWidth)/7 & "&nbsp;Open</b></font></TD>" 
		strClosedWidth = clng(strAllClosed) * 7'/ (clng(strAllOpen) + clng(strAllClosed))*100)
		Response.Write "<TD bgcolor=green width=" & strClosedWidth  &  "><font color=white size=1><b>" & clng(strClosedWidth)/7 & "&nbsp;Closed</b></font></TD>" 
		%>
	</TR>
</Table>
<TABLE border=1 width=<%=((clng(strAllOpenp1) + clng(strAllClosedp1) * 7)+200)%> cellpadding=0 cellspacing=0 bordercolor=white>
	<TR width = <%=((clng(strAllOpenp1) + clng(strAllClosedp1)) * 7)+200%>>
		<TD width=200><font size=1>P1 Observations:</TD>
		<%
		strOpenWidth = clng(strAllOpenp1) * 7 '/ (clng(strAllOpen) + clng(strAllClosed)) * 100)
		Response.Write "<TD bgcolor=red width=" & strOpenWidth  &  "><font size=1 color=white><b>" & clng(strOpenWidth)/7 & "&nbsp;Open</b></font></TD>" 
		strClosedWidth = clng(strAllClosedp1) * 7'/ (clng(strAllOpen) + clng(strAllClosed))*100)
		Response.Write "<TD bgcolor=green width=" & strClosedWidth  &  "><font color=white size=1><b>" & clng(strClosedWidth)/7 & "&nbsp;Closed</b></font></TD>" 
		%>
	</TR>
</TABLE>-->





<%if blnOTSLinked and  (not (strAllOpenP1 = 0 and strAllClosedP1 = 0)) then%>
<BR>
<%
'	set cn = server.CreateObject("ADODB.Connection")
'	cn.ConnectionString = Session("PDPIMS_ConnectionString")
'	cn.Open

'	set rs = server.CreateObject("ADODB.recordset")

	dim strWeekData
	dim WeekCount
	dim AllOpen
	dim Weeks()
	dim StartWeek
	dim StartYear
	dim WeekIndex
	dim PORWeek
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListOTSSubmittedByWeek"

	Set p = cm.CreateParameter("@ID", 200, &H0001,50)
	p.Value = left(strOTSName,50)
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	'Determine Week COunt
	if (rs.EOF and rs.BOF) then
		WeekCount = 0
	else
		StartWeek = rs("SubmitWeek")
		StartYear = rs("SubmitYear")

		if isdate(strPORDate) then
			PORWeek = GetWeekIndex(StartWeek, StartYear, datepart("ww",strPORDate), datepart("yyyy",strPORDate))
		else
			PORWeek = 0
		end if
		
		'Response.Write "PORWeek:" & PORWeek

		if StartYear = Datepart("yyyy",Now) then
			WeekCount = (Datepart("ww",Now) - StartWeek)+ 1
		else
			WeekCount = (Datepart("ww",Now) + (53-StartWeek) +(53* (Datepart("yyyy",Now) - StartYear-1))) + 1
		end if
		Redim Weeks(WeekCount,3)
	end if

	if WeekCount <> 0 then
		do while not rs.EOF
			WeekIndex = GetWeekIndex(StartWeek, StartYear, rs("SubmitWeek"), rs("SubmitYear"))
			Weeks(WeekIndex,0) = rs("PerWeek")
			rs.MoveNext
		loop
		rs.Close
	end if
	

	strWeekData = ""
	for i = 0 to weekcount -1
		if weeks(i,0) = "" then
			strWeekData = strWeekData & ",0"
		else
			strWeekData = strWeekData & "," & weeks(i,0)
		end if
	next
	if len(strWeekData) > 0 then
		strWeekData = mid(strWeekData,2)
	end if

	strOpenData = strWeekData

%>
<INPUT type="hidden"  id=txtOpenData name=txtOpenData value=" <%=strWeekData%>" style="WIDTH: 563px; HEIGHT: 22px" size=72>
<INPUT type="hidden" id=txtRowCount name=txtRowCount value=" <%=WeekCount%>">
<%
	if WeekCount <> 0 then
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListOTSClosedByWeek"

		Set p = cm.CreateParameter("@ID", 200, &H0001,50)
		p.Value = left(strOTSName,50)
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
		
		'rs.Open "spListOTSClosedByWeek '" & strOTSName & "'",cn,adOpenForwardOnly
		do while not rs.EOF
			if not (isnull(rs("SubmitWeek")) or isnull(rs("SubmitYear"))) then
				WeekIndex = GetWeekIndex(StartWeek, StartYear, rs("SubmitWeek"), rs("SubmitYear"))
				Weeks(WeekIndex,1) = rs("PerWeek") & ""
			end if
			rs.MoveNext
		loop
		rs.Close
	end if

	strWeekData = ""
	for i = 0 to weekcount -1
		if weeks(i,1) = "" then
			strWeekData = strWeekData & ",0"
		else
			strWeekData = strWeekData & "," & weeks(i,1)
		end if
	next

	if len(strWeekData) > 0 then
		strWeekData = mid(strWeekData,2)
	end if
	
	strClosedData = strWeekData

%>
<INPUT type="hidden"  id=txtClosedData name=txtClosedData value=" <%=strWeekData%>" style="WIDTH: 563px; HEIGHT: 22px" size=72>
<%	
	AllOpen = 0
	strWeekData = ""
	for i = 0 to weekcount -1
		AllOpen = AllOpen + (weeks(i,0) - weeks(i,1))
		strWeekData = strWeekData & "," & AllOpen
	next 

	if len(strWeekData) > 0 then
		strWeekData = mid(strWeekData,2)
	end if

	strBacklogData = strWeekData

%>

<INPUT type="hidden"  id=txtBackLog name=txtBackLog value=" <%=strWeekData%>" style="WIDTH: 563px; HEIGHT: 22px" size=72>
<INPUT type="hidden"  id=txtProduct name=txtProduct value=" <%=strproductName%>" style="WIDTH: 563px; HEIGHT: 22px" size=72>
<INPUT type="hidden"  id=txtPORWeek name=txtPORWeek value=" <%=PORWeek%>" style="WIDTH: 563px; HEIGHT: 22px" size=72>

<%end if
	case 12

  rs.Open "spListDCRThisWeek 2," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd & ", 0",cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=DCROpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>(DCR) Change Requests Opened</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=DCROpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 4 align=center><font color=black size=2><b>(DCR) Change Requests Opened</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>Submitter</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td></tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return DCRROW_onclick('" & rs("id") & "')""><TD class=static><font size=1>"
		select case rs("Status")
		case 1
			strStatus = "Open"			
		case 2
			strStatus = "Need More Input"			
		case 3
			strStatus = "Closed"			
		case 4
			strStatus = "Approved"			
		case 5
			strStatus = "Disapproved"			
		case 6
			strStatus = "Investigating"			
		case else
			strStatus = "N/A"
		end select		
%>
		<font size=1 face=verdana><%=rs("ID")%> </FONT><%= midrow %><FONT size=1><%=strStatus %></FONT><%= midrow %><FONT size=1><%= shortname(rs("Submitter"))%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if

	case 13

%>
<!--DCR Closed Section-->

<%

  rs.Open "spListDCRThisWeek 3," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd & ", 0",cn,adOpenForwardOnly
	'Response.Write "spListDCRThisWeek 3," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd
  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=DCRClosedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>(DCR) Change Requests Closed</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=DCRClosedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 4 align=center><font color=black size=2><b>(DCR) Change Requests Closed</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>Submitter</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td></tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return DCRROW_onclick('" & rs("id") & "')""><TD class=static><font size=1>"
		select case rs("Status")
		case 1
			strStatus = "Open"			
		case 2
			strStatus = "Need More Input"			
		case 3
			strStatus = "Closed"			
		case 4
			strStatus = "Approved"			
		case 5
			strStatus = "Disapproved"			
		case 6
			strStatus = "Investigating"			
		case else
			strStatus = "N/A"
		end select		
%>
		<font size=1 face=verdana><%=rs("ID")%> </FONT><%= midrow %><FONT size=1><%=strStatus %></FONT><%= midrow %><FONT size=1><%= shortname(rs("Submitter"))%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if


	case 14

%>
<!--Observations Added-->

<%

  rs.Open "spListObservationsThisWeek " & clng(strproductID) & ",3," & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=OTSOpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>Observations Opened</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=OTSOpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 6 align=center><font color=black size=2><b>Observations Opened</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Owner</font></strong></td>
	<td nowrap><strong><font color=black size=1>PR</font></strong></td>
	<td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>State</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td>
	</tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return OTSROW_onclick('" & rs("Observationid") & "')""><TD class=static><font size=1>"
		strStatus = rs("Status") & ""
%>
		<font size=1 face=verdana><%=rs("ObservationID")%> </FONT><%= midrow %><FONT size=1><%=rs("Owner") %></FONT><%= midrow %><FONT size=1><%= rs("Priority") %>&nbsp;</font><%= midrow %><FONT size=1><%= strStatus%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("State")%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if


	case 15

%>
<!--Observations Closed-->

<%

  rs.Open "spListObservationsThisWeek " & clng(strproductID) & ",2," & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=OTSOpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>Observations Closed</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=OTSOpenedTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 6 align=center><font color=black size=2><b>Observations Closed</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Owner</font></strong></td>
	<td nowrap><strong><font color=black size=1>PR</font></strong></td>
	<td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>State</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td>
	</tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return OTSROW_onclick('" & clng(rs("Observationid")) & "')""><TD class=static><font size=1>"
		
        strStatus = rs("Status") & ""
%>
		<font size=1 face=verdana><%=rs("ObservationID")%> </FONT><%= midrow %><FONT size=1><%=rs("Owner") %></FONT><%= midrow %><FONT size=1><%= rs("Priority") %>&nbsp;</font><%= midrow %><FONT size=1><%= strStatus%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("State")%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if

	Case 16
	'
	' Agency Status History
	'
	Response.Write "<BR>"

	rs.Open "usp_SelectAgencyStatusHistory " & strProductID & ", NULL, " & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

	Dim sUser

	if rs.EOF and rs.BOF then
		Response.Write _
			"<table ID=AgencyStatusChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
			"<tr bgcolor=Gainsboro><TD align=center><font color=Black size=2><b>Agency Status Changes</b></font></TD></TR>" & _
			"<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
	
		rs.Close
	else
	Response.Write _
		"<table ID=AgencyStatusChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
		"<tr bgcolor=Gainsboro><TD colspan = 4 align=center><font color=black size=2><b>Agency Status Changes</b></font></TD></TR>" & _
		"<tr bgColor=Gainsboro>" & _
		"<td nowrap width=10><strong><font color=black size=1>Type</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Changed By</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Date</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Summary</font></strong></td></tr>"

	'Display requirements
		Do Until rs.EOF
			sUser = rs("Changed_By")
			If InStr(sUser, ",") Then
				sUser = Left(sUser, InStr(sUser, ",") + 2) & "."
			End If
		
			Response.Write prerow
			Response.Write Trim(rs("Change_Type") & "")
			Response.Write midrow
			Response.Write sUser
			Response.Write midrow
			Response.Write rs("Date_Of_Change")
			Response.Write midrow
			Response.Write rs("Change_Summary")
			Response.Write postrow
			rs.MoveNext
		Loop
  
		'Finish off table 
		Response.Write "</TBODY></table>"
  
		'Cleanup
		rs.Close
  
	End If

			Case 17
	'
	' Schedule Change History
	'
	Response.Write "<BR>"

	rs.Open "usp_SelectScheduleDataHistory " & strProductID & ", " & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly


	if rs.EOF and rs.BOF then
		Response.Write _
			"<table ID=ScheduleChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
			"<tr bgcolor=Gainsboro><TD align=center><font color=Black size=2><b>Schedule Changes</b></font></TD></TR>" & _
			"<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
	
		rs.Close
	else
	Response.Write _
		"<table ID=ScheduleChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
		"<tr bgcolor=Gainsboro><TD colspan = 8 align=center><font color=black size=2><b>Schedule Changes</b></font></TD></TR>" & _
		"<tr bgColor=Gainsboro>" & _
		"<td nowrap width=10 rowspan=2><strong><font color=black size=1>Schedule</font></strong></td>" & _
		"<td nowrap rowspan=2><strong ><font color=black size=1>Changed By</font></strong></td>" & _
		"<td nowrap rowspan=2><strong><font color=black size=1>Date Changed</font></strong></td>" & _
		"<td nowrap rowspan=2><strong><font color=black size=1>Item</font></strong></td>" & _
		"<td nowrap colspan=2 align=center><strong><font color=black size=1>Old Date</font></strong></td>" & _
		"<td nowrap colspan=2 align=center><strong><font color=black size=1>New Date</font></strong></td></tr>" & _
		"<tr bgColor=Gainsboro>" & _
		"<td nowrap align=center><strong><font color=black size=1>Start</font></strong></td>" & _
		"<td nowrap align=center><strong><font color=black size=1>End</font></strong></td>" & _
		"<td nowrap align=center><strong><font color=black size=1>Start</font></strong></td>" & _
		"<td nowrap align=center><strong><font color=black size=1>End</font></strong></td></tr>"

	'Display requirements
		Do Until rs.EOF
			sUser = rs("User_Name")
			If InStr(sUser, ",") Then
				sUser = Left(sUser, InStr(sUser, ",") + 2) & "."
			End If


			If rs("old_actual_end_dt") & "" = "" And _
				rs("new_actual_end_dt") & "" <> "" And _
				rs("schedule_definition_data_id") = 7 Then
				
				Response.Write prerow
				Response.Write Trim(rs("schedule_name") & "")
				Response.Write midrow
				Response.Write sUser
				Response.Write midrow
				Response.Write FormatDateTime(rs("last_upd_date")&"", vbShortDate)
				Response.Write "</FONT></TD><TD valign=top colspan=5><FONT size=1>"
				Response.Write Trim(rs("item_description") & "") & " achieved on " & FormatDateTime(rs("new_actual_end_dt") & "", vbShortDate)
				Response.Write postrow
			
			ElseIf rs("old_actual_end_dt") & "" = "" And _
				rs("new_actual_end_dt") & "" = "" And _
				UCase(rs("milestone_yn") & "") = "N" Then
				
				Response.Write prerow
				Response.Write Trim(rs("schedule_name") & "")
				Response.Write midrow
				Response.Write sUser
				Response.Write midrow
				Response.Write FormatDateTime(rs("last_upd_date")&"", vbShortDate)
				Response.Write midrow
				Response.Write rs("item_description")
				Response.Write MidRowCenter
				
				If rs("old_projected_start_dt")&"" = "" Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("old_projected_start_dt") & ""
				End If
				
				Response.Write MidRowCenter
				
				If rs("old_projected_end_dt")&"" = "" Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("old_projected_end_dt") & ""
				End If
				
				Response.Write MidRowCenter
				If rs("new_projected_start_dt") & "" = ""Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("new_projected_start_dt") & ""
				End If
				Response.Write MidRowCenter
				
				If rs("new_projected_end_dt") & "" = ""Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("new_projected_end_dt") & ""
				End If
				Response.Write postrow

			ElseIf rs("old_actual_end_dt") & "" = "" And _
				rs("new_actual_end_dt") & "" = "" Then
				
				Response.Write prerow
				Response.Write Trim(rs("schedule_name") & "")
				Response.Write midrow
				Response.Write sUser
				Response.Write midrow
				Response.Write FormatDateTime(rs("last_upd_date")&"", vbShortDate)
				Response.Write midrow
				Response.Write rs("item_description")
				Response.Write MidRowCenterColSpan2
				
				If rs("old_projected_end_dt")&"" = "" Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("old_projected_end_dt") & ""
				End If
				
				Response.Write MidRowCenterColSpan2
				If rs("new_projected_end_dt")&"" = "" Then
					Response.Write "&nbsp;"
				Else
					Response.Write rs("new_projected_end_dt") & ""
				End If
				Response.Write postrow
				
			End If
			rs.MoveNext
		Loop
  
		'Finish off table 
		Response.Write "</TBODY></table>"
  
		'Cleanup
		rs.Close
  
	End If
		
			Case 18
	'
	' Country Change History
	'
	Response.Write "<BR>"

	rs.Open "usp_SelectProdBrandCountryHistory " & strProductID & ", " & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

	
	
	if rs.EOF and rs.BOF then
		Response.Write _
			"<table ID=CountryChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
			"<tr bgcolor=Gainsboro><TD align=center><font color=Black size=2><b>Supported Country Changes</b></font></TD></TR>" & _
			"<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
	
		rs.Close
	else
	Response.Write _
		"<table ID=CountryChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
		"<tr bgcolor=Gainsboro><TD colspan = 5 align=center><font color=black size=2><b>Supported Country Changes</b></font></TD></TR>" & _
		"<tr bgColor=Gainsboro>" & _
		"<td nowrap width=10><strong><font color=black size=1>Type</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Changed By</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Date</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Summary</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>DCR</font></strong></td></tr>"

	'Display requirements
		Do Until rs.EOF
			sUser = rs("Last_Upd_User")
			If InStr(sUser, ",") Then
				sUser = Left(sUser, InStr(sUser, ",") + 2) & "."
			End If
			
			If LCase(rs("ChangeType")) = "added" Then
				sChangeSummary = rs("Country") & " was " & LCase(rs("ChangeType")) & " to the " & rs("Brand") & " brand."
			Else
				sChangeSummary = rs("Country") & " was " & LCase(rs("ChangeType")) & " from the " & rs("Brand") & " brand."
			End If
			
			Response.Write prerow
			Response.Write Trim(rs("ChangeType") & "")
			Response.Write midrow
			Response.Write sUser
			Response.Write midrow
			Response.Write rs("Last_Upd_Date")
			Response.Write midrow
			Response.Write sChangeSummary
			Response.Write midrow
			Response.Write rs("DcrID")
			Response.Write postrow
			rs.MoveNext
		Loop
  
		'Finish off table 
		Response.Write "</TBODY></table>"
  
		'Cleanup
		rs.Close
  
	End If

			Case 19
	'
	' Localization Change History
	'
	Response.Write "<BR>"

	rs.Open "usp_SelectProdBrandCountryLocalizationHistory " & strProductID & ", " & ReportDays & ", " & dtReportStart & ", " & dtReportEnd,cn,adOpenForwardOnly

	if rs.EOF and rs.BOF then
		Response.Write _
			"<table ID=AgencyStatusChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
			"<tr bgcolor=Gainsboro><TD align=center><font color=Black size=2><b>Localization Changes</b></font></TD></TR>" & _
			"<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
	
		rs.Close
	else
	Response.Write _
		"<table ID=AgencyStatusChangeTable border=1 borderColor=black cellPadding=2 cellSpacing=0 width='100%' >" & _
		"<tr bgcolor=Gainsboro><TD colspan = 5 align=center><font color=black size=2><b>Localization Changes</b></font></TD></TR>" & _
		"<tr bgColor=Gainsboro>" & _
		"<td nowrap width=10><strong><font color=black size=1>Type</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Changed By</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Date</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>Summary</font></strong></td>" & _
		"<td nowrap><strong><font color=black size=1>DCR</font></strong></td></tr>"

	'Display requirements
		Do Until rs.EOF
			sUser = rs("Last_Upd_User")
			If InStr(sUser, ",") Then
				sUser = Left(sUser, InStr(sUser, ",") + 2) & "."
			End If
		
			sChangeSummary = rs("OptionConfig") & "/" & rs("Dash") & " was " & LCase(rs("ChangeType")) & " for " & rs("Country") & " on " & rs("Brand") & "."

			Response.Write prerow
			Response.Write Trim(rs("ChangeType") & "")
			Response.Write midrow
			Response.Write sUser
			Response.Write midrow
			Response.Write rs("Last_Upd_Date")
			Response.Write midrow
			Response.Write sChangeSummary
			Response.Write midrow
			Response.Write rs("DcrID")
			Response.Write postrow
			rs.MoveNext
		Loop
  
		'Finish off table 
		Response.Write "</TBODY></table>"
  
		'Cleanup
		rs.Close
  
	End If

	case 20

%>
<!--DCR Opened Section-->

<%

  rs.Open "spListDCRThisWeek 2," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd & ", 1",cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=Table1 border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>(BCR) Change Requests Opened</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=Table2 border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 4 align=center><font color=black size=2><b>(BCR) Change Requests Opened</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>Submitter</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td></tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return DCRROW_onclick('" & rs("id") & "')""><TD class=static><font size=1>"
		select case rs("Status")
		case 1
			strStatus = "Open"			
		case 2
			strStatus = "Need More Input"			
		case 3
			strStatus = "Closed"			
		case 4
			strStatus = "Approved"			
		case 5
			strStatus = "Disapproved"			
		case 6
			strStatus = "Investigating"			
		case else
			strStatus = "N/A"
		end select		
%>
		<font size=1 face=verdana><%=rs("ID")%> </FONT><%= midrow %><FONT size=1><%=strStatus %></FONT><%= midrow %><FONT size=1><%= shortname(rs("Submitter"))%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if

	case 21

%>
<!--DCR Closed Section-->

<%

  rs.Open "spListDCRThisWeek 3," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd & ", 1",cn,adOpenForwardOnly
	'Response.Write "spListDCRThisWeek 3," & ReportDays & "," & clng(strproductID) & ", " & dtReportStart & ", " & dtReportEnd
  if rs.EOF and rs.BOF then
	%>
	<BR>
	<table ID=Table3 border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
	  <tr bgcolor=Gainsboro> <TD align=center><font color=Black size=2><b>(BCR) Change Requests Closed</b></font></TD></TR>
		<% response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow & "</TABLE>"	
		rs.Close
  else
  %>
  <BR>
	<table ID=Table4 border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%" LANGUAGE=javascript onmouseover="return ROW_onmouseover()" onmouseout="return ROW_onmouseout()">
	  <tr bgcolor=Gainsboro> <TD colspan = 4 align=center><font color=black size=2><b>(BCR) Change Requests Closed</b></font></TD></TR>

	  <tr bgColor=Gainsboro>
    <td nowrap width=10><strong><font color=black size=1>Number</font></strong></td>	  
    <td nowrap><strong><font color=black size=1>Status</font></strong></td>
	<td nowrap><strong><font color=black size=1>Submitter</font></strong></td>
	<td nowrap><strong><font color=black size=1>Summary</font></strong></td></tr>
<%  
  'Display requirements
  do while not rs.EOF
		Response.Write "<TR bgcolor=white valign=top class=""Row"" LANGUAGE=javascript onclick=""return DCRROW_onclick('" & rs("id") & "')""><TD class=static><font size=1>"
		select case rs("Status")
		case 1
			strStatus = "Open"			
		case 2
			strStatus = "Need More Input"			
		case 3
			strStatus = "Closed"			
		case 4
			strStatus = "Approved"			
		case 5
			strStatus = "Disapproved"			
		case 6
			strStatus = "Investigating"			
		case else
			strStatus = "N/A"
		end select		
%>
		<font size=1 face=verdana><%=rs("ID")%> </FONT><%= midrow %><FONT size=1><%=strStatus %></FONT><%= midrow %><FONT size=1><%= shortname(rs("Submitter"))%>&nbsp;</font><%= midrow %><FONT size=1><%= rs("Summary")%>&nbsp;</font>
<%			Response.Write postrow
	rs.MoveNext
  loop
  
   'Finish off table 
    Response.Write "</TBODY></table>"
  
  'Cleanup
  rs.Close
  
  end if

		case 22
	%>
<!-- Start Scope Change Section -->
	<BR>
	<table ID=Table5 border=1 borderColor=black cellPadding=2 cellSpacing=0 width="100%">
		<tr bgColor=Gainsboro> <TD align=center><font color=black size=2><b>Change Requests (BCR)</b></font></TD></TR>
<%
	rs.Open "spListActionItems4Status " & clng(strproductID) & ",3" & ", " & dtReportStart & ", " & dtReportEnd & ", 1",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=white valign=top><TD height=""30px"">&nbsp;" & "</TD>" &  postrow
	else
		do while not rs.EOF	

'			Response.Write "<FONT face=verdana size=2><b>ID#: <a href=""javascript:DisplayAction(" & rs("ID") & "," & rs("Type") & ");"">" & rs("ID") & "</a></b></font><BR>"
'			Response.Write "<FONT face=verdana size=1><b>Summary: " & rs("Summary") & "</b></font><BR>"
'			Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"">"

			strResolution = trim(rs("Resolution") & "")
			if isnull(rs("ActualDate")) then
				Response.Write "<TR><TD>"
			elseif rs("Status") & "" = "1" or rs("Status") & "" = "3" or rs("Status") & "" = "6" then
				Response.Write "<TR><TD>"
			else
				Response.Write "<TR bgcolor=ivory><TD>"
			end if


			if trim(rs("Status") & "") <> "" then
				select case  rs("Status")
				case 1
					strStatus = "Proposed"
				case 2
					strStatus = "Closed"
				case 3
					strStatus = "Need More Information"
				case 4
					strStatus = "Approved"
				case 5
					strStatus = "Disapproved"
				case 6
					strStatus = "Investigating"
				end select 
			else
				strDesc = ""
			end if


			Response.Write "<Table width=100% cellspacing=0 cellpadding=1 border=1 bordercolor=gainsboro>"
			Response.Write "<TR><TD colspan=3><font size=1 face=verdana><b>" & rs("Summary") & "</B></font></TD></TR>"

			Response.Write "<TR><TD><TABLE  width=""100%""><TR><TD nowrap><font face=verdana size=1>ID:</font></TD><TD><font face=verdana size=1>" & rs("ID") & "</font></TD></TR><TR><TD><font face=verdana size=1>Product:</font></TD><TD><font face=verdana size=1>" & rs("Product") & "</font></TD></TR><TR><TD><font face=verdana size=1>Status:</font><TD><font face=verdana size=1>" & strStatus & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Date Created:</font></TD><TD><font face=verdana size=1>" & rs("Created") & "</font></TD></TR><TR><TD><font face=verdana size=1>Days Open:</font></TD><TD><font face=verdana size=1>" & DateDiff("d",rs("Created"),Date()) & "</font></TD></TR><TR><TD><font face=verdana size=1>Target Date:</font><TD><font face=verdana size=1>" & rs("TargetDate") & "</font></TD></TR></table></TD>"
			Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Submitter:</font></TD><TD><font face=verdana size=1>" & rs("Submitter") & "</font></TD></TR><TR><TD><font face=verdana size=1>Owner:</font></TD><TD><font face=verdana size=1>" & rs("Owner") & "</font></TD></TR><TR><TD><font face=verdana size=1>Core Team Rep:</font><TD><font face=verdana size=1>" & rs("CoreTeamRep") & "</font></TD></TR></table></TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>Author</u><BR>" & rs("AuthorFullname") & "<BR>" & rs("AuthorGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Module PM</u><BR>" & rs("PMFullName") & "<BR>" & rs("PMGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Developer</u><BR>" & rs("DeveloperFullName") & "<BR>" & rs("DeveloperGroup") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Owner</u><BR>" & rs("FullName") & "<BR>" & rs("OwnerGroup") & "</font></TD></tr></table>   </TD></TR>"
'			'Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><font face=verdana size=1><u>System Build</u><BR>" & rs("Systemboardrev") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>System ROM</u><BR>" & rs("SystemROM") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>OS</u><BR>" & rs("OSRelease") & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Release</u><BR>" & strImage & "</font></TD><TD align=center nowrap><font face=verdana size=1><u>Image Language</u><BR>" & strImageLanguage & "</font></TD></tr></table></TD></TR>"
			Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & replace(rs("Description"),vbcrlf,"<BR>") & "</font></td></tr></table></TD></TR>"
			if trim(rs("Approvals") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Approvals: </font></td><td><font size=1 face=verdana>" & replace(rs("Approvals"),vbcrlf,"<BR>")  &"</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Justification") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Justification: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(rs("Actions") & "") <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Actions: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Actions") & "",vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") & "</font></td></tr></table></TD></TR>"
			end if
			if trim(strResolution) <> "" then
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=85 nowrap valign=top><font size=1 face=verdana>Resolution: </font></td><td><font size=1 face=verdana>" & replace(replace(strResolution,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR><BR>") &"</font></td></tr></table></TD></TR>"
			end if
			Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "</TD></TR>"


			
			rs.MoveNext
		loop

	end if
%>	</TABLE><%
	rs.Close


		
				end select
		end if
	next	
	set rs = nothing
	set cn = nothing





%>

<br>
<br>
<font size="1">Report Generated <%=formatdatetime(date(),vblongdate) %></font>
<br>
<br>
<font Size="2" Color="red"><p><strong>Confidential</strong></p></font>


<%end if%>

</body>
</html>

<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<%
	dim cn, rs, strSQL, strProgram, strFamily, strVersion, WorkflowStep
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")

	on error resume next
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	If cn.State = 0 and request("redir") <> "true" then
		err.Clear
		session.Abandon
		Response.Redirect "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?redir=true"
	ElseIf cn.State = 0 Then
		Server.Transfer "/_Error/500.asp"
	End If

	on error goto 0
		
	set rs = server.CreateObject("ADODB.recordset")


	'Get User
	dim CurrentuserID
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPartner
    dim CurrentUserSysAdmin
    dim bSystemAdmin
	dim isSEPM
	
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
	DisableEmail="disabled"
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserPartner = rs("PartnerID") & ""
        
        '***PBI 8650 / Task 16199 - Harris, Valerie - Replace hardcoded SystemAdmin Users with System Admin boolean check: ---
        CurrentUserSysAdmin = rs("SystemAdmin")
        bSystemAdmin = CBool(CurrentUserSysAdmin)
	else
		rs.Close
		set rs=nothing
		cn.Close
		set cn = nothing
		Response.Redirect "../Excalibur.asp"
	end if
	rs.Close
	
	isSEPM = false
	if CurrentUserID  <> 0 then
		rs.Open "spListPMsActive 3",cn,adOpenForwardOnly
		do while not rs.EOF
			if CurrentUserID = rs("ID") then
				isSEPM = true
				exit do
			end if
			rs.MoveNext
		loop
		rs.Close
	end if


%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="JavaScript">
<title>SE Program Information - HP Restricted</title>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<!-- #include file="../includes/bundleConfig.inc" -->
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oPopup = window.createPopup();

function UpdateUserAccess() {
    window.location.href = "../UpdateUserAccess.asp";
}

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function programMenu(strProgram,strID)
{
    // The variables "lefter" and "topper" store the X and Y coordinates
    // to use as parameter values for the following show method. In this
    // way, the popup displays near the location the user clicks. 

    var lefter = event.clientX;
    var topper = event.clientY;

    var popupBody;
		
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=parent.location.href=\"javascript:ProgramOptions(1,1,'" + strProgram + "')\" >&nbsp;&nbsp;&nbsp;Documents...</SPAN></FONT></DIV>";


	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";


	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=parent.location.href=\"javascript:ProgramOptions(2," + strID + ",'" + strProgram + "')\"> &nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";



	popupBody = popupBody + "</DIV>";

txtHidden.value = popupBody;
    
    oPopup.document.body.innerHTML = popupBody; 

	oPopup.show(lefter, topper, 130, 85, document.body);

	//Adjust window size
	var NewHeight;
	var NewWidth;

	NewHeight = oPopup.document.body.scrollHeight;
	NewWidth = oPopup.document.body.scrollWidth;
	oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);


}

function ProgramOptions(FunctionID,ID,strProgram){
	if (FunctionID == 1){
		    window.open ("\\\\houhpqexcal03.auth.hpicorp.net\\se_web\\SoftwarePOR\\" + strProgram,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes"); 
	}else{
	    modalDialog.open({ dialogTitle: 'Update Group', dialogURL: '../program/program.asp?ID=' + ID + '', dialogHeight: 500, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });
	}
}


function AddProgram(){
    modalDialog.open({ dialogTitle: 'Add Group', dialogURL: '../program/program.asp', dialogHeight: 500, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });
}

function PRow_onmouseover() {
   	window.event.srcElement.parentElement.style.color = "red";
	window.event.srcElement.parentElement.style.cursor = "hand";
}

function PRow_onmouseout() {
   	window.event.srcElement.parentElement.style.color = "blue";
}

function OpenStatusOptions(ID){
    modalDialog.open({ dialogTitle: 'Status Options', dialogURL: '../ProductStatusOptions.asp?ID=' + ID + '', dialogHeight: 400, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
}

//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  11/08/2016 - Harris, Valerie - PBI 28261/ Task 29166   
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();
}
//-->
</script>
</head>
<style>
A:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
body { 
  background-repeat: repeat-y;
}

</style>
<body onload="window_onload();" background="images/shadow.gif" ><!-- This is the Header Table --><!-- This is the Header Table -->
<div style="display:none"><a href="../UpdateUserAccess.asp"></a></div>
<table style="display:none" border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td nowrap width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
	<td><font size="5" face="Tahoma" color="#006697"><b>SE Program Office Information</b></font></td>
<!--    <td><IMG height=50 src="images/information.gif" width=137></td>    <td align="right"><IMG height=50 src="images/programoffice.gif" width=283></td> --> </tr>
</table>
    
	<%if CurrentUserPartner = 1 then%>
    <!-- #include file = "menubar.asp" -->
<%else%>
    <!-- #include file = "menubarODM.asp" -->
<%end if%>

	<%
		dim blnEditPrograms
		
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
        if isSEPM = true Or bSystemAdmin = true then
			blnEditPrograms = true
		else
			blnEditPrograms = false
		end if
		
	if CurrentUserPartner = 1 then

        dim strInactiveSwitch
        strInactiveSwitch = "<a href=""default.asp?showall=1"">Show Inactive Programs</a>"
        if request("showall") = "1" then
            strInactiveSwitch = "<a href=""default.asp?showall=0"">Hide Inactive Programs</a>"
        end if
	%>

    <!--<table><tr><td nowrap><font size="4" face="tahoma" color="#006697"><b>General Documents</b></font></td></tr></table>-->

	<table border="0" cellpadding="3" width="100%">
		<tr><td width="10">&nbsp;</td>
		<td colspan="2" nowrap valign="top"><b><font size="4" color="#006697">Product Groups</font></b><br><hr color="steelblue">
		<!--<b>Owner: </b>Lorri Jefferson<br>-->
		<!--&nbsp;&nbsp;&nbsp;<img src="images/FOLDER.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/softwarePor/General">General&nbsp;Documents</a>-->

		</td></tr>
		<% if blnEditPrograms then%>
		<tr><td width="10">&nbsp;</td><td colspan="2">&nbsp;&nbsp;&nbsp;<a href="javascript: AddProgram();">Add Group</a>&nbsp;<%=strInactiveSwitch %></td></tr>
        <% else %>
        <tr><td width="10">&nbsp;</td><td colspan="2">&nbsp;&nbsp;&nbsp;<%=strInactiveSwitch %></td></tr>
		<% end  if %>

		<tr><td width="10">&nbsp;</td><td colspan="2"><table width="100%" border="0" cellpadding="2" cellspacing="1">
		<%
		dim strIcon
		dim lastProgramID
        dim LastProgramGroup
		dim lastProgramActive
		dim showInactive
            		
        'lastProgramActive = true
        LastProgramGroup = ""
		LastProgram = ""
        strProducts = ""
		lastProgramID = 0
		strProduct=""
        showInactive = false
        if request("showall") = "1" then
            showInactive = true
        end if



		rs.Open "spGetProgramTree",cn,adOpenForwardOnly
		do while not rs.EOF
			if LastProgram <> rs("Program") and LastProgram <> "" then
				if strProducts <> "" then
					strProducts = mid(strProducts,3)	
				end if
			'	if LastProgramGroup = "Commercial" then
			'		if blnEditPrograms then
			'			Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return ProgramOptions(2," & lastProgramID & ",'')"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & replace(lastProgram," Business NB Cycle","") & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
			'		else
			'			Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;" & replace(lastProgram," Business NB Cycle","") & "</td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
			'		end if
			'	else
					if blnEditPrograms then
						'Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return programMenu('" & replace(lastProgram," ","") &  "'," & lastProgramID & ")"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & lastProgram & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
                        if cInt(lastProgramActive) = 1 then
						    Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><font color=blue><b><u ID=PRow LANGUAGE=javascript onclick=""return ProgramOptions(2," & lastProgramID & ",'" & replace(lastProgram," ","") &  "')"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & lastProgram & "</u></b></font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
					    elseif showInactive then
                            Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return ProgramOptions(2," & lastProgramID & ",'" & replace(lastProgram," ","") &  "')"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & lastProgram & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
                        end if
                    else
                        if cInt(lastProgramActive) = 1 then
                            Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;" & lastProgram & "</td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"                        
					    elseif showInactive then
						    Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle><font color=SlateGray>&nbsp;" & lastProgram & "</font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
                        end if

'						Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<a target=""_blank"" href=""file://\\houhpqexcal03.auth.hpicorp.net\se_web\SoftwarePOR\" & replace(lastProgram," ","") & """>" & lastProgram & "</a></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
					end if
			'	end if
				strProducts=""
				strCurrent = ""
			end if
			if LastProgramGroup <> rs("ProgramGroup") then
				if LastProgramGroup <> "" then
					Response.Write "<TR height=5><td width=10></TD><TD bgcolor=white colspan=2><font size=1 face=verdana></font></TD></TR>"
				end if
				Response.Write "<TR><td width=10></TD><TD bgcolor=#006697 colspan=2><b><font size=2 face=verdana color=white>" & rs("ProgramGroup") & "</font></b></TD></TR>"
				Response.Write "<tr style=display:none><td width=10>&nbsp;</td><td bgcolor=#9999cc><font color=black size=2 face=verdana><b>Cycle</b></font></td><td bgcolor=#9999cc><font color=black size=2 face=verdana><b>Products</font></b></td></tr>"

			end if
			LastProgramGroup = rs("ProgramGroup")
			strProducts = strProducts & ", " & "<a target=""_blank"" href=""../pmView.asp?ID=" & rs("ID") & "&List=General"">" & rs("product") & "</a>"
			LastProgram = rs("Program") & ""
			lastProgramID = rs("ProgramID")
            lastProgramActive = rs("active")
            
			'if LastProgram = "2C04" then
			'if rs("CurrentProgram") then
			'	strCurrent = "<IMG Title=""Current Cycle"" SRC=""../images/red.gif"">"
			'end if
			rs.MoveNext
		loop
		rs.Close



		if strProducts <> "" then
			strProducts = mid(strProducts,3)	
		end if
		if LastProgram <> "" then
			'if LastProgramGroup = "Commercial" then
			'	if blnEditPrograms then
			'		Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return ProgramOptions(2," & lastProgramID & ",'')"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & replace(lastProgram," Business NB Cycle","") & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
			'	else
			'		Response.Write "<TR><td width=10>" & strCurrent & "<td nowrap bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;" & replace(lastProgram," Business NB Cycle","") & "</td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
			'	end if
			'else
				if blnEditPrograms then
                    if cInt(lastProgramActive) = 1 then
					    Response.Write "<TR><td width=10>" & strCurrent & "<td bgcolor=""Lavender""><font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return ProgramOptions(2," & lastProgramID & ",'" & replace(lastProgram," ","") &  "')"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & replace(lastProgram," Business NB Cycle","") & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</td></tr>"
                    else
                        Response.Write "<TR><td width=10>" & strCurrent & "<td bgcolor=""Lavender""><font color=blue>" & replace(lastProgram," Business NB Cycle","") & "</font></td><td bgcolor=""Lavender"">" & strProducts & "</td></tr>"
                    end if
'					Response.Write "<TR><td width=10>" & strCurrent & "<td bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<font color=blue><u ID=PRow LANGUAGE=javascript onclick=""return programMenu('" & replace(lastProgram," ","") &  "'," & lastProgramID & ")"" onmouseover=""return PRow_onmouseover()"" onmouseout=""return PRow_onmouseout()"">" & replace(lastProgram," Business NB Cycle","") & "</u></font></td><td bgcolor=""Lavender"">" & strProducts & "</td></tr>"
				else
					'Response.Write "<TR><td width=10>" & strCurrent & "<td bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;<a target=""_blank"" href=""file://\\houhpqexcal03.auth.hpicorp.net\se_web\SoftwarePOR\" & replace(lastProgram," Business NB Cycle","Business") & """>" & replace(lastProgram," Business NB Cycle","") & "</a></td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
					Response.Write "<TR><td width=10>" & strCurrent & "<td bgcolor=""Lavender""><img src=""images/FOLDER.GIF"" HEIGHT=16 width=16 align=middle>&nbsp;" & replace(lastProgram," Business NB Cycle","") & "</td><td bgcolor=""Lavender"">" & strProducts & "</a></td></tr>"
				end if
			'end if
		end if



		%>
        </table>

        </td></tr>
	</table>
<br> 
<% end if%>
    <%if false then%>
	<table border="0" cellpadding="3" width="100%">
		<tr><td width="10">&nbsp;</td>
		<td colspan="2" nowrap valign="top"><b><font size="3" color="#006697">Miscellaneous Documents</font></b><br><hr color="steelblue">
		<!--<b>Owner: </b>Lorri Jefferson-->
		</td></tr>
		<%if CurrentUserPartner = 1 then%>
		<tr><td width="10">&nbsp;</td><td width="335">
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/Microsoft/MS%20Hotfix%20Tracking.xls">MS Hotfix Tracking</a><br>
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank"  href="http://teams1.sharepoint.hp.com/teams/BNBSE/Shared%20Documents/SE%20Documents%20and%20Templates/PwrMgmtSetting%202009.xlsx">Power Management Settings 2009</a><br>
        <img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="../MobileSE/ConfigCodes.asp">HP code-CPQ DASH Conversion Table</a><br>
        <img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="http://houcmitrel02.auth.hpicorp.net:81/models/product_id.aspx">Product/Series/Model Cross-Reference</a><br>
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/Sustaining/11504%20to%2021504/SE%20Softpaq%20Status%20115%20to%20215.xls">SPRP Sustaining</a><br>
		<img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="#Sustaining">Sustaining&nbsp;Products</a><br>        
		<img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/3PSW/3rd%20Party%20Dashboard.xls">3rd&nbsp;Party&nbsp;Dashboard</a>        
        </td><td valign="top">
		<img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="http://psgmarketing.corp.hp.com/earlyeval">Early&nbsp;Eval&nbsp;-&nbsp;Post&nbsp;Mortem&nbsp;Reports</a><br>        
        <img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="Countries.asp">Master Country List</a><br>
        <img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="../MobileSE/ConfigCodes.asp">Master Localization List</a><br>
        <img src="images/ICON-DOC-PPT.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="http://teams1.sharepoint.hp.com/teams/BNBSE/Shared%20Documents/Image%20Map%20Recovery/Image%20Recovery%20Map%202009.xlsx">Image Recovery Map 2009</a><br>        
        <img src="images/ICON-DOC-PPT.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="http://teams1.sharepoint.hp.com/teams/BNBSE/Shared%20Documents/SE%20Documents%20and%20Templates/BNB%20Windows%20Recovery%20OS%20Matrix%20(updated).pptx">BNB Windows Recovery OS Matrix</a><br>        
		<img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="ProductScheduleSummary.asp">Schedule&nbsp;Summary</a><br>        
		<img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="ProductSummary.asp">Product&nbsp;Summary</a><br>        
        </td></tr>
        <%else%>
		<tr><td width="10">&nbsp;</td><td width="335">
       
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/SWStrategy/RegionalDefaults.xls">Regional Default Settings</a><br>
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/SWStrategy/LangSub-Table.xls">General Language Substitution Info</a><br>
        <img src="images/ICON-DOC-EXCEL.GIF" HEIGHT="16" width="16" align="middle"> <a target="_blank" href="file://<%= strFileServer%>/SWStrategy/AppLangSub-Table.xls">Application Specific Substitution Info</a><br>
        </td><td valign="top">
		<img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="Countries.asp">Master Country List</a><br>
        <img src="images/ICON-DOC-HTML.GIF" HEIGHT="16" width="16" align="middle"> <a href="../MobileSE/ConfigCodes.asp">Master Localization List</a>
        </td></tr>
        <%end if%>
	</table>
<br><br>
  
    <table><tr><td nowrap><font size="4" face="tahoma" color="#006697"><b>Product Information</b></font></td></tr></table>
	<%	if CurrentUserPartner = 1 then%>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="ProductSummary.asp?ShowAll=5&DevCenter=0&ODM=0" target="_blank">Summary of Active Products</a>
	<%end if%>
	<table border="0" cellpadding="20">
<%
	strSQL = "spGetProducts"
	rs.Open strSQL,cn,adOpenForwardOnly

	rs.MoveFirst

	dim columncounter
	columncounter = 1
	Do while not rs.EOF
		if rs("Division") = 1 and rs("Name") & "" <> "Test Product"  and rs("active") and (trim(CurrentUserPartner) = "1" or trim(CurrentUserPartner) = trim(rs("PartnerID")) ) then
			if columncounter =1 then
				Response.Write "<tr>"
			end if
			if trim(rs("DevCenter")) = "2" or trim(rs("DevCenter")) = "3" or trim(rs("DevCenter")) = "4" then
				lblPM = "PM:"
			else
				lblPM = "CM:"
			end if
			
			Response.Write "<td width=25% valign=top ><b><font size=3 color=#006697 FONT-FAMILY: Tahoma>" + rs("Name") + " " + rs("Version") + "</font></b><br><hr color=steelblue>"
			if rs("SEPM") <> "" then
				Response.write "<b>" & lblPM & " </b>" + rs("PM") + "<br>"
				Response.write "<b>SE PM: </b>" + rs("SEPM") + "<br><br>"
			else		
				Response.write "<b>SE PM: </b> Not Assigned" + "<br><br>"
			end if

			set rs2 = server.CreateObject("ADODB.recordset")
			rs2.Open  "spListbrands4Product " & rs("ID") & ",1",cn,adOpenForwardOnly
			strSeries = ""
			do while not rs2.EOF
				if trim(rs2("SeriesSummary") & "") = "" then
					strSeries = ""
				else
					SeriesArray = split(rs2("SeriesSummary"),",")
					for i = 0 to ubound(SeriesArray)
						if rs2("StreetName") <> "" then
							strSeries =  strSeries & rs2("StreetName") & " " 
						end if
						strSeries = strSeries & seriesArray(i) 
						if trim(rs2("Suffix") & "") <> "" then
							strSeries = strSeries & " " & trim(rs2("Suffix") & "")
						end if
						strSeries = strSeries & "<BR>"
					next
				end if
				rs2.MoveNext
			loop
			rs2.Close
			set rs2 = nothing
			
			if trim(strSeries) <> "" then
				Response.write "<font color=green size=1>" & strSeries & "</font><br>"
			end if
	
			if rs("Sustaining") then
			'	Response.write "<b><font size=1>Sustaining</font></b><br><br>"
				workflowstep = 3
			elseif rs("PDDReleased") & "" <> "" then
			'	Response.write "<b><font size=1>Product Definition Complete</font></b><br><br>"
				workflowstep = 3
			elseif rs("PRDReleased") & ""  <> "" then
				Response.write "<b><font size=1 color=red>In Product Definition Phase</font></b><br><br>"
				workflowstep = 2
			else
				Response.write "<b><font size=1 color=red>In Product Definition Phase</font></b><br><br>"
				workflowstep = 1
			end if
			
			dim strProductFileName
			strProductFileName = replace(rs("Name") + " " + rs("Version")," ", " ")
			strProductFileName = replace(strProductFileName,"/", "")
			
			%>
			<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="_blank" href="../SystemTeam.asp?ID=<%=rs("ID")%>">System Team Roster</a><br>
	

			<%if rs("OnlineReports") & "" = "1" then%>
				<%if strproductFilename = "Raptor 1.1" then%>
					<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\Raptor%201.1\status.doc">Current Status</a><br>
				<%else%>
					<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="_blank" href="../ProductStatus.asp?ID=<%=rs("ID")%>">Current Status</a><br>
				<%end if%>
			<%else%>
				<%if strProductFilename <> "Topaz 1.0" then%>
					<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_WEB\<%= strProductFilename %>\Status.doc">Current Status</a><br>
				<%end if%>
			<%end if%>

			<%if ucase(left(strProductFilename,11)) = "TORNADO 2.1" or ucase(left(strProductFilename,11)) = "TORNADO 2.2"  or ucase(left(strProductFilename,11)) = "TORNADO 4.7"  or ucase(left(strProductFilename,11)) = "TORNADO 5.6"  or ucase(left(strProductFilename,11)) = "TORNADO 5.5"  or ucase(left(strProductFilename,11)) = "TORNADO 4.6" then%>
				<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_WEB\General\1.5C%20Tornado%20Image%20Rollout%20rev%2019.xls">Tornado 1.5C03 Image Rollout Matrix</a><br>
			<%end if %>
		

			<%if ucase(left(strProductFilename,10)) = "FENWAY 2.0" or ucase(left(strProductFilename,10)) = "FENWAY 2.X" or ucase(left(strProductFilename,10)) = "AWARDS 2.0" or ucase(left(strProductFilename,11)) = "CROCKET 1.0" or ucase(left(strProductFilename,12)) = "MAGELLAN 1.0"  or ucase(left(strProductFilename,9)) = "BOONE 1.0"  or ucase(left(strProductFilename,9)) = "BOONE 1.1"  or ucase(left(strProductFilename,9)) = "BOXER ALL"  or ucase(left(strProductFilename,13)) = "DIRECTORS 1.X"  or ucase(left(strProductFilename,9)) = "DYLAN 1.X"  or ucase(left(strProductFilename,10)) = "EAGLES 1.X"  or ucase(left(strProductFilename,10)) = "FUSION 1.2"  or ucase(left(strProductFilename,11)) = "LEADERS 1.7"   or ucase(left(strProductFilename,12)) = "MAGELLAN 1.1"   or ucase(left(strProductFilename,10)) = "NASCAR ALL"   or ucase(left(strProductFilename,13)) = "STARBUCKS ALL"  or ucase(left(strProductFilename,11)) = "TEMPEST ALL"   or ucase(left(strProductFilename,8)) = "U2II 1.X"   then%>
				<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_WEB\<%= strProductFilename %>\DeliverableMatrix.xls">Deliverable Matrix</a><br>
			<%elseif rs("OnlineReports") & "" = "1" then'if ucase(left(strProductFilename,10)) = "LOPEZ 1.0" or ucase(left(strProductFilename,10)) = "RAPTOR 2.0" or ucase(left(strProductFilename,8)) = "RYAN 1.0" or ucase(left(strProductFilename,8)) = "FORD 1.0"  or ucase(left(strProductFilename,8)) = "RUBY 1.0" or  ucase(left(strProductFilename,11)) = "DIAMOND 1.0" or ucase(left(strProductFilename,11)) = "DIAMOND 2.0"  then%>
				<a target="_blank" href="../image/DeliverableMatrix.asp?ProdID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="../image/DeliverableMatrix.asp?ProdID=<%=rs("ID")%>">Deliverable Matrix</a><br>
				<a target="_blank" href="../image/localization.asp?ProdID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="../image/localization.asp?ProdID=<%=rs("ID")%>">Localization Matrix</a><br>
				<%if rs("OnCommodityMatrix") then%>
					<!--<a target="_blank" href="../Deliverable/commodity/QualMatrix.asp?Products=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="../Deliverable/commodity/QualMatrix.asp?Products=<%=rs("ID")%>">Commodity Qual Matrix</a><br>-->
					<a target="_blank" href="../Deliverable/HardwareMatrix.asp?lstProducts=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="../Deliverable/HardwareMatrix.asp?lstProducts=<%=rs("ID")%>">Hardware Qual Matrix</a><br>
				<%end if%>
				<%if 0=1 and ucase(left(strProductFilename,10)) = "LOPEZ 1.0"  then%>
					<a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Lopez%201.0\Lopez%20rollout%20plan%20SE.xls"><img SRC="images/ICON-DOC-EXCEL.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Lopez%201.0\Lopez%20rollout%20plan%20SE.xls">SE Rollout Plan</a><br>
					<a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Lopez%201.0\LopezSDD.doc"><img SRC="images/ICON-DOC-WORD.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Lopez%201.0\LopezSDD.doc">System Definition</a><br>

				<%end if%>
			<%end if%>
			
			<%if ucase(left(strProductFilename,8)) = "RUBY 1.0"  then%>
				<a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Ruby%201.0\Ruby%201.0%20System%20Definition%20Document.doc"><img SRC="images/ICON-DOC-WORD.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Ruby%201.0\Ruby%201.0%20System%20Definition%20Document.doc">System Definition</a><br>
			<%elseif ucase(left(strProductFilename,11)) = "DIAMOND 1.0"  then%>
				<a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\Diamond%201.0\Diamond%201.0%20SDD.doc"><img SRC="images/ICON-DOC-WORD.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\Diamond%201.0\Diamond%201.0%20SDD.doc">System Definition</a><br>
			<%elseif ucase(left(strProductFilename,10)) = "RAPTOR 2.0"  then%>
				<a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\Raptor%202.0\Raptor2.0_SDD.doc"><img SRC="images/ICON-DOC-WORD.GIF" border="0" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\Raptor%202.0\Raptor2.0_SDD.doc">System Definition</a><br>
			<%end if%>
			
			
			<%if ucase(left(strProductFilename,10)) = "FENWAY 3.0" or ucase(left(strProductFilename,12)) = "SAPPHIRE 1.0"then%>
				<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Sapphire%201.0\Legacy%20OS%20Engineering%20Support%20-%20Fenway%203.0%20%20Sapphire.doc">Legacy NT/98 Support Document</a><br>
			<%end if%>
			<%if ucase(left(strProductFilename,12)) = "SAPPHIRE 1.0"then%>
				<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Sapphire%201.0\SDD.doc">System Definition Document</a><br>
			<%end if%>
			<%if ucase(left(strProductFilename,11)) = "CRYSTAL 1.0"then%>
				<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_Web\Crystal%201.0\Crystal%201.0%20SDD.doc">System Definition Document</a><br>
			<%end if	
	
			if 0=1 and ucase(left(strProductFilename,5)) = "DYLAN" or ucase(left(strProductFilename,6)) = "EAGLES" then%>
				<img SRC="images/ICON-DOC-WORD.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_WEB\<%= strProductFilename %>\ProductDef.doc">Product Definition</a><br>
			<%elseif 0=1 and rs("ProductFilePath") & "" <> "" then%>
				<img SRC="images/folder.GIF" WIDTH="16" HEIGHT="16"> <a href="file://<%= rs("ProductFilePath") & rs("PDDPath") %>">Product Definition</a><br>
			<% end if%>
		
		
			<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="new" href="../search/ots/default.asp?lstProduct=<%=rs("Name") + " " + rs("Version")%>">Custom OTS Reports</a><br>
			<%if rs("OnlineReports") & "" = "1" then%>
				<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="new" href="../pmView.asp?ID=<%=rs("ID")%>&amp;List=General">Excalibur Product Info</a><br>
				<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="new" href="../Image/Buildplan.asp?ID=<%=rs("ID")%>">Rollout Plan</a><br>
				<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="new" href="../Image/Buildplan.asp?ID=<%=rs("ID")%>&amp;Report=1">Ramp Plan</a><br>
			<%end if%>	

			<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a href="javascript:OpenStatusOptions(<%=rs("ID")%>);">Changes This Week</a><br>
			<%if rs("PRDReleased") & ""  <> "" then%>
				<img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16"> <a target="_blank" href="../PNPDevices.asp?Report=2&ProductID=<%=rs("ID")%>">Device ID List</a><br>
			<%end if%>
			
			
			<%if CurrentUserPartner = 1 then%>
				<%if rs("PDDPath") & "" <> "" then%>
					<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="<%=replace(rs("PDDPath"),"%20"," ")%>">PDD</a><br>
				<%end if%>
				<%if rs("SCMPath") & "" <> "" then%>
					<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="<%=replace(rs("SCMPath"),"%20"," ")%>">SCM</a><br>
				<%end if%>
				<%if rs("STLStatusPath") & "" <> "" then%>
					<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="<%=replace(rs("STLStatusPath"),"%20"," ")%>">STL Status</a><br>
				<%end if%>
				<%if rs("ProgramMatrixPath") & "" <> "" then%>
					<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="<%=replace(rs("ProgramMatrixPath"),"%20"," ")%>">Product Data Matrices</a><br>
				<%end if%>
			<%end if%>			


			<%if trim(rs("Distribution")) <> "" then%>
				<img SRC="images/ICON-DOC-OUTLOOK.GIF" WIDTH="16" HEIGHT="16"> <a HREF="mailto:<%= rs("Distribution") %>">Send Email to System Team</a><br>
			<%else%>
				<img SRC="images/ICON-DOC-OUTLOOK.GIF" WIDTH="16" HEIGHT="16"> No Email List Defined<br>		
			<%end if %>
			<%if 0=1 and  trim(rs("ProductFilePath")) <> "" then%>
				<img SRC="images/folder.gif" WIDTH="16" HEIGHT="16"> <a target="new" href="file://<%=rs("ProductFilePath") %>">Other Documents</a><br>
			<%end if%>
			
			
			</td>
			<%
			if columncounter = 4 then
				Response.Write "</tr>"
				columncounter = 1
			else
				columncounter = columncounter + 1
			end if
	end if
			rs.movenext
		loop
	%>	
	</table>
    <%end if%>

<!--Sustaining Products-->
<%if false then 'CurrentUserPartner = 1 then%>

<br><br><br>
<a name="Sustaining">
    <table><tr><td nowrap><font size="4" face="tahoma" color="#006697"><b>Products Release to the Factory</b></font></td></tr></table>
	<table border="0" cellpadding="20">

<%	rs.MoveFirst

	columncounter = 1
	Do while not rs.EOF
		if rs("Division") = 1 and rs("Name") & "" <> "Test Product"  and rs("Sustaining") and (trim(CurrentUserPartner) = "1" or trim(CurrentUserPartner) = trim(rs("PartnerID"))) then
			if columncounter =1 then
				Response.Write "<tr>"
			end if
			Response.Write "<td nowrap valign=top ><b><font size=3 color=#006697 FONT-FAMILY: Tahoma>" + rs("Name") + " " + rs("Version") + "</font></b><br><hr color=steelblue>"
'			if rs("SEPM") <> "" then
'				Response.write "<b>PM: </b>" + rs("PM") + "<br><br>"
'			else		
'				Response.write "<b>PM: </b> Not Assigned" + "<br><br>"
'			end if
	
			workflowstep = 3
			
			strProductFileName = replace(rs("Name") + " " + rs("Version")," ", " ")
			strProductFileName = replace(strProductFileName,"/", "")


			
			%>
			<%if ucase(left(strProductFilename,10)) = "FENWAY 2.0" or ucase(left(strProductFilename,10)) = "FENWAY 2.X" or ucase(left(strProductFilename,10)) = "AWARDS 2.0" or ucase(left(strProductFilename,11)) = "CROCKET 1.0" or ucase(left(strProductFilename,12)) = "MAGELLAN 1.0" or ucase(left(strProductFilename,9)) = "BOONE 1.0"  or ucase(left(strProductFilename,9)) = "BOONE 1.1"  or ucase(left(strProductFilename,12)) = "CROCKETT 1.0" or ucase(left(strProductFilename,10)) = "FUSION 1.X" or ucase(left(strProductFilename,8)) = "KONG 1.0"  or ucase(left(strProductFilename,13)) = "DIRECTORS 1.X"  or ucase(left(strProductFilename,9)) = "DYLAN 1.X"  or ucase(left(strProductFilename,10)) = "EAGLES 1.X"  or ucase(left(strProductFilename,10)) = "FUSION 1.2"  or ucase(left(strProductFilename,11)) = "LEADERS 1.7"   or ucase(left(strProductFilename,11)) = "TORNADO 5.1"    or ucase(left(strProductFilename,8)) = "U2II 1.X"   then%>
				<img SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\SE_WEB\<%= strProductFilename %>\DeliverableMatrix.xls">Deliverable Matrix</a><br>
				<img SRC="images/folder.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\<%= strproductFilename %>">Other Documents</a><br>
			<%else%>
				<img SRC="images/folder.GIF" WIDTH="16" HEIGHT="16"> <a href="file://\\houhpqexcal03.auth.hpicorp.net\se_web\<%= strproductFilename %>">Documentation</a><br>
			<%end if%>
			

			<%'if rs("OnlineReports") & "" = "1" then%>
			<%'end if%>
			
			<%
			if columncounter = 4 then
				Response.Write "</tr>"
				columncounter = 1
			else
				columncounter = columncounter + 1
			end if
	end if
			rs.movenext
		loop
	%>	
	</table>

</td></tr></table>
<%end if%>

 <div id="cc"><h1>HP Restricted</h1><p>Last Update <%=Date%></p></div>
<textarea style="Display:none" rows="4" cols="80" id="txtHidden" name="txtHidden">

</textarea>
</body>
</html>
<%
	'rs.Close
%>
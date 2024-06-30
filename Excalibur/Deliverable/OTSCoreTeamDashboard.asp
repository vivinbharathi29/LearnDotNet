<%@ Language=VBScript %>

	<%
	
	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
	else
        Response.Buffer = True
        Response.ExpiresAbsolute = Now() - 1
        Response.Expires = 0
        Response.CacheControl = "no-cache"
	end if
	
	if lcase(Session("LoggedInUser")) = "auth\dwhorton" or lcase(Session("LoggedInUser")) = "auth\lyoung" then 
		blnAdmin = true
	else
		blnAdmin = false
	end if
	
	Dim AppRoot : AppRoot = Session("ApplicationRoot")

    ' Turn On Error Handling
    On Error Resume Next
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>


<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Deliverable Status - Confidential</title>
	<STYLE>
		TD{
	    FONT-SIZE: xx-small;
	    FONT-FAMILY: Verdana;
	    }
		Body{
	    FONT-FAMILY: Verdana;
	    FONT-SIZE: x-small;
	    }
A:link
{
    COLOR: Blue;
}
A:visited
{
    COLOR: Blue;
}
A:hover
{
    COLOR: red;
} 
	</STYLE>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "..\_ScriptLibrary/sort.js" -->


function window_onload() {
	lblOTS.style.display="none";
}

function ActionMouseOver(ID){
    document.all("ActionRow" + ID).style.cursor="hand";
}

function ActionClick(ID){
	var strResult;
    var strStatus;
    var strAction;

	strResult = window.showModalDialog("Scorecard/RootScorecard.asp?ID=" + ID + "&Action=1","","dialogWidth:700px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
    if (typeof(strResult) != "undefined")
		{	
            if (strResult.length > 1)
                {
                strStatus= strResult.substring(0,1);
                if (strStatus=="0")
                    {
                    document.all("ActionStatus" + ID).innerHTML = "No&nbsp;Status";
                    document.all("ActionStatus" + ID).style.backgroundColor = "white";
                    }
                else if (strStatus=="1")
                    {
                    document.all("ActionStatus" + ID).innerText = "Open";
                    document.all("ActionStatus" + ID).style.backgroundColor = "#9999FF";
                    }
                else if (strStatus=="2")
                    {
                    document.all("ActionStatus" + ID).innerText = "Closed";
                    document.all("ActionStatus" + ID).style.backgroundColor = "#99CC66";
                    }
                document.all("ActionRow" + ID).innerHTML = strResult.substring(1);
                }
        }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%

	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

    if request("Report") = "1" then
        response.write "<table width=""100%""><tr><td><font size=3 face=verdana><b>Executive Summary Report</b></td><td align=right><font size=1>Yellow and Red statuses only</font></font></td></tr></table><BR>"
    elseif request("Report") = "2" then
        response.write "<table width=""100%""><tr><td><font size=3 face=verdana><b>Action Item Report</b></td><td align=right><font size=1>Populated items only</font></font></td></tr></table><BR>"
    end if

    strProducts = request("lstProducts")

	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		
	'	strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
	
  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout=120
	cn.IsolationLevel=256
	cn.Open

  'Create a recordset
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn
	
%>


	<label ID=lblOTS>
	<%if request("cboFormat") <> "1" and request("cboFormat") <> "2" then%>
	Accessing OTS.  Please Wait...
	<%end if%>
	</label>
<%
 
    dim rs2

    if request("RootID") = "" and request("CoreTeamID") = "" then
        response.Write "Not enough information supplied to run this report."
    else
        dim imageHeight
        dim imageWidth
        if request("RootID") = "" then
            imageHeight = 260
            imageWidth = 380
        else
            imageHeight = 350
            imageWidth = 500
        end if
        
    dim strExecutiveSummary
    dim ExecutiveSummaryStatus
    dim strHPPeopleProcess
    dim HPPeopleProcessStatus
    dim strHPEquipment
    dim EquipmentStatus
    dim strSupplierPeopleProcess
    dim SupplierPeopleProcessStatus
    dim strSupplierDeliverables
    dim SupplierDeliverablesStatus
    dim strAction
    dim ActionStatus
    dim StatusColorArray
    dim StatusNameArray
    dim ActionStatusColorArray
    dim ActionStatusNameArray
    dim NextCodeFreeze
    dim P1backlog
    dim DateUpdated
    dim ActionUpdated
    dim ExecutiveSummaryUpdated

    dim SI_error
    SI_error = 0

'    StatusColorArray = split("white,#FFCCCC,#FFFFCC,#E6FFCC",",")
    StatusColorArray = split("white,#FF9999,#FFFF66,#99CC66",",")
    StatusNameArray = split("No&nbsp;Status,Red,Yellow,Green",",")

    ActionStatusColorArray = split("white,#9999FF,#99CC66",",")
    ActionNameArray = split("No&nbsp;Status,Open,Closed",",")

    dim MilestoneCount
    dim strMilestoneCells()
    
    redim strMilestoneCells(3,2)

    
    
        strSQL = "Select r.id , ct.name as Coreteam, r.Name, e.name as DevManager, vd.name as Vendor " & _
                 "from deliverableroot r with (NOLOCK), deliverablecoreteam ct with (NOLOCK), vendor vd with (NOLOCK), employee e with (NOLOCK) " & _
                 "where ct.id = r.coreteamid " & _
                 "and r.vendorid = vd.id " & _
                 "and e.id = r.devmanagerid "
        if request("RootID") <> "" then
            strSQl = strSQl & " and r.id in ( "  & scrubsql(request("RootID")) & ") "
        else
            strSQl = strSQl & " and r.showonstatus=1 and r.coreteamid in( "  & scrubsql(request("CoreTeamID")) & ") "
        end if
        strSQL = strSQL & " order by r.name"
        rs.open strSQl, cn
      if Err.Number = 0 then
        response.Write "<table style=""width:100%"" border=1 cellpadding=3 cellspacing=0>"
        do while not rs.eof
            set rs2 = server.CreateObject("ADODB.recordset")
            rs2.open "spListDelRootUpcomingProductMilestones " & rs("ID"),cn

            strMilestoneCells(0,0) = "CF Product: N/A"
            strMilestoneCells(0,1) = "CF Date: N/A"
            strMilestoneCells(1,0) = "CF Product: N/A"
            strMilestoneCells(1,1) = "CF Date: N/A"
            strMilestoneCells(2,0) = "CF Product: N/A"
            strMilestoneCells(2,1) = "CF Date: N/A"
            NextCodeFreeze= "N/A"
            MilestoneCount = 0
            do while not rs2.eof
                strMilestoneCells(MilestoneCount,0) = "Code freeze for " & rs2("Product") 
                strMilestoneCells(MilestoneCount,1) = "CF Date: " & rs2("MilestoneDate") & "&nbsp;&nbsp;" & rs2("LeadDays") & " Days"
                NextCodeFreeze = rs2("LeadDays") & " Days"
                if MilestoneCount = 2 then
                    exit do
                end if
                MilestoneCount = MilestoneCount + 1
                rs2.movenext
            loop
            rs2.close
            set rs2 = nothing

            set rs2 = server.CreateObject("ADODB.recordset")
            rs2.open "spGetRootScoreCard " & rs("ID"),cn
            if rs2.eof and rs2.bof then
                strExecutiveSummary = ""
                ExecutiveSummaryStatus = 0
                strHPPeopleProcess = ""
                HPPeopleProcessStatus = 0
                strHPEquipment = ""
                EquipmentStatus = 0
                strSupplierPeopleProcess = ""
                SupplierPeopleProcessStatus = 0
                strSupplierDeliverables = ""
                SupplierDeliverablesStatus = 0
                strAction = ""
                ActionStatus = 0
                DateUpdated=""
                ActionUpdated=""
                ExecutiveSummaryUpdated=""
            else
                strExecutiveSummary = replace(rs2("ExecutiveSummary") & "",vbcrlf,"<br>")
                ExecutiveSummaryStatus = rs2("ExecutiveSummaryStatus") & ""
                strHPPeopleProcess = replace(rs2("HPPeopleProcess") & "",vbcrlf,"<br>")
                HPPeopleProcessStatus = rs2("HPPeopleProcessStatus") & ""
                strHPEquipment = replace(rs2("HPEquipment") & "",vbcrlf,"<br>")
                EquipmentStatus = rs2("HPEquipmentStatus") & ""
                strSupplierPeopleProcess = replace(rs2("SupplierPeopleProcess") & "",vbcrlf,"<br>")
                SupplierPeopleProcessStatus = rs2("SupplierPeopleProcessStatus") & ""
                strSupplierDeliverables = replace(rs2("SupplierDeliverables") & "",vbcrlf,"<br>")
                SupplierDeliverablesStatus = rs2("SupplierDeliverablesStatus") & ""
                strAction = replace(rs2("Action") & "",vbcrlf,"<br>")
                ActionStatus = rs2("ActionStatus") & ""
                DateUpdated=trim(rs2("DateUpdated") & "")
                ActionUpdated=trim(rs2("ActionUpdated") & "")
                ExecutiveSummaryUpdated=trim(rs2("ExecutiveSummaryUpdated") & "")
            end if
            rs2.close
            set rs2=nothing

            if isdate(DateUpdated) then
                DateUpdated = "Updated:&nbsp;" & formatdatetime(DateUpdated,vbshortdate) 
            end if
            if isdate(ActionUpdated) then
                ActionUpdated = "Updated:&nbsp;" & formatdatetime(ActionUpdated,vbshortdate) 
            end if
            if isdate(ExecutiveSummaryUpdated) then
                ExecutiveSummaryUpdated = "Updated:&nbsp;" & formatdatetime(ExecutiveSummaryUpdated,vbshortdate) 
            end if

            if request("Report") = "1" then
                set rs2 = server.CreateObject("ADODB.recordset")
                rs2.open "spgetOTSStatusTotals4Root " & rs("ID"),cn
                if rs2.eof and rs2.bof then
                    P1backlog = "Unknown"
                else
                    P1backlog = rs2("TotalP1") & ""
                end if
                rs2.close
                if ExecutiveSummaryStatus = 1 or ExecutiveSummaryStatus=2 then
                    response.Write "<tr style=""height:10px""><td bgcolor=blue colspan=10></td></tr>"
                    response.Write "<tr>"
                    response.Write "<td bgcolor=gainsboro colspan=2 nowrap><table width=""100%"" cellpadding=0 cellspacing=0><tr><td align=left><b>" & rs("name") & "</td><td align=right>" & ExecutiveSummaryUpdated & "</td></tr></table></td>"
                    response.Write "<td bgcolor=gainsboro colspan=2 align=center><b>" & rs("Devmanager") & "</td>"
                    response.Write "</tr>"
                    response.Write "<tr>"
                    response.Write "<td valign=top rowspan=2 bgcolor=gainsboro>Executive<br>Summary<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(ExecutiveSummaryStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(ExecutiveSummaryStatus) & "</span></td>"
                    response.Write "<td valign=top rowspan=2 style=""width:100%"">" & strExecutiveSummary & "&nbsp;</td>"
                    response.Write "<td nowrap bgcolor=gainsboro>Next Code Freeze</td>"
                    response.Write "<td nowrap bgcolor=gainsboro>P1 OTS Backlog</td></tr><tr>"
                    response.Write "<td nowrap bgcolor=white align=center>" & NextCodeFreeze & "</td>"
                    response.Write "<td nowrap bgcolor=white align=center>" & P1backlog & "</td>"
                    response.Write "</tr>"
                end if
            elseif request("Report") = "2" then
                if strAction <> "" then
                    set rs2 = server.CreateObject("ADODB.recordset")
                    rs2.open "spgetOTSStatusTotals4Root " & rs("ID"),cn
                    if rs2.eof and rs2.bof then
                        P1backlog = "Unknown"
                    else
                        P1backlog = rs2("TotalP1") & ""
                    end if
                    rs2.close
                    response.Write "<tr style=""height:10px""><td bgcolor=blue colspan=10></td></tr>"
                    response.Write "<tr>"
                    response.Write "<td bgcolor=gainsboro colspan=2 nowrap><table width=""100%"" cellpadding=0 cellspacing=0><tr><td align=left><b>" & rs("name") & "</td><td align=right>" & ActionUpdated & "</td></tr></table></td>"
                    response.Write "<td bgcolor=gainsboro colspan=2 align=center><b>" & rs("Devmanager") & "</td>"
                    response.Write "</tr>"
                    response.Write "<tr>"
                    response.Write "<td valign=top rowspan=2 bgcolor=gainsboro>Actions<BR><BR><span id=ActionStatus" & trim(rs("ID")) & " style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & ActionStatusColorArray(ActionStatus) & "; padding-right: 3px; padding-left: 3px;"">" & ActionNameArray(ActionStatus) & "</span></td>"
                    response.Write "<td valign=top rowspan=2 style=""width:100%"" onmouseover=""javascript:ActionMouseOver(" & trim(rs("ID")) & ");"" onclick=""javascript:ActionClick(" & trim(rs("ID")) & ");"" id=""ActionRow" & rs("ID") & """>" & strAction & "&nbsp;</td>"
                    response.Write "<td nowrap bgcolor=gainsboro>Next Code Freeze</td>"
                    response.Write "<td nowrap bgcolor=gainsboro>P1 OTS Backlog</td></tr><tr>"
                    response.Write "<td nowrap bgcolor=white align=center>" & NextCodeFreeze & "</td>"
                    response.Write "<td nowrap bgcolor=white align=center>" & P1backlog & "</td>"
                    response.Write "</tr>"
                end if
            else
                response.Write "<tr style=""height:10px""><td bgcolor=blue colspan=10></td></tr>"
                response.Write "<tr>"
                response.Write "<td bgcolor=gainsboro colspan=2 nowrap><table width=""100%"" cellpadding=0 cellspacing=0><tr><td align=left><b>" & rs("Coreteam") & " - " & rs("name") & "</td><td align=right>" & DateUpdated & "</td></tr></table></td>"
                'response.Write "<td bgcolor=gainsboro align=center><b>" & rs("name") & "</td>"
                response.Write "<td bgcolor=gainsboro colspan=3 align=center><b>" & rs("Devmanager") & "</td>"
                response.Write "<td bgcolor=gainsboro colspan=4 align=center><b>" & rs("Vendor") & "</td>"
                response.Write "</tr>"
                response.Write "<tr>"
                response.Write "<td valign=top rowspan=3 bgcolor=gainsboro>Executive<br>Summary<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(ExecutiveSummaryStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(ExecutiveSummaryStatus) & "</span></td>"
                response.Write "<td valign=top rowspan=3 style=""width:100%"">" & strExecutiveSummary & "&nbsp;</td>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=3>" & strMilestoneCells(0,0) & "</td>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=4>" & strMilestoneCells(0,1) & "</td>"
                response.Write "</tr>"
                response.Write "<tr>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=3>" & strMilestoneCells(1,0) & "</td>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=4>" & strMilestoneCells(1,1) & "</td>"
                response.Write "</tr>"
                response.Write "<tr>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=3>" & strMilestoneCells(2,0) & "</td>"
                response.Write "<td nowrap bgcolor=gainsboro colspan=4>" & strMilestoneCells(2,1) & "</td>"
                response.Write "</tr>"

                response.Write "<tr>"
                response.Write "<td valign=top rowspan=2 bgcolor=gainsboro><u>HP</u><br>People&nbsp;&<br>Process<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(HPPeopleProcessStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(HPPeopleProcessStatus) & "</span></td>"
                response.Write "<td valign=top rowspan=2>" & strHPPeopleProcess & "&nbsp;</td>"
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Backlog<BR>P1</td>" '<img src=""Images/backlog1.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Investigating<BR>P1</td>" '<img src=""Images/Investigating.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Identified<BR>P1</td>" '<img src=""Images/Identified.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>FIP<BR>P1</td>" '<img src=""Images/fip.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Retest<BR>P1</td>" '<img src=""Images/retest.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Web<BR>P1</td>" '<img src=""Images/retest.gif""/>
                response.Write "<td nowrap bgcolor=gainsboro align=center valign=bottom>Backlog<BR>P2</td>" '<img src=""Images/backlog2.gif""/>
                response.Write "</tr>"
            
                set rs2 = server.CreateObject("ADODB.recordset")
                rs2.open "spgetOTSStatusTotals4Root " & rs("ID"),cn
                if rs2.eof and rs2.bof then
                    response.Write "<tr>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "<td nowrap bgcolor=gainsboro align=center>Unknown</td>"
                    response.Write "</tr>"
                else
                    response.Write "<tr>"
                    if rs2("TotalP1") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("TotalP1") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=1"">" & rs2("TotalP1") & "</a></td>"
                    end if
                    if rs2("UI") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("UI") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=2"">" & rs2("UI") & "</a></td>"
                    end if
                    if rs2("Identified") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("Identified") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=3"">" & rs2("Identified") & "</a></td>"
                    end if
                    if rs2("FIP") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("FIP") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=4"">" & rs2("FIP") & "</a></td>"
                    end if
                    if rs2("Retest") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("Retest") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=5"">" & rs2("Retest") & "</a></td>"
                    end if
                    if rs2("Web") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("Web") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=6"">" & rs2("Web")  & "</a></td>"
                    end if
                    if rs2("TotalP2") = 0 then
                        response.Write "<td nowrap bgcolor=gainsboro align=center>" & rs2("TotalP2") & "</td>"
                    else
                        response.Write "<td nowrap bgcolor=gainsboro align=center><a target=_blank href=""OTSCoreTeamDashboardOTSList.asp?RootID=" & rs("ID") & "&StatusID=7"">" & rs2("TotalP2") & "</a></td>"
                    end if
                    response.Write "</tr>"
                end if
                rs2.close
                set rs2= nothing
            
                response.Write "<tr>"
                response.Write "<td valign=top bgcolor=gainsboro><u>HP</u><BR>Equipment<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(EquipmentStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(EquipmentStatus) & "</span></td>"
                response.Write "<td valign=top>" & strHPEquipment & "&nbsp;</td>"
                'response.Write "<td valign=center align=center rowspan=4 colspan=7><img style=""height:" & imageHeight & ";width:" & imageWidth & """ id=myPic src=""../temp/rad060ED.tmp.gif""></td>"
                response.Write "<td valign=center align=center rowspan=4 colspan=7>" 

                response.Write "</td>"
                response.Write "</tr>"
            
                response.Write "<tr>"
                response.Write "<td valign=top bgcolor=gainsboro><u>Supplier</u><BR>People&nbsp;&<br>Process<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(SupplierPeopleProcessStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(SupplierPeopleProcessStatus) & "</span></td>"
                response.Write "<td valign=top>" & strSupplierPeopleProcess & "&nbsp;</td>"
                response.Write "</tr>"

                response.Write "<tr>"
                response.Write "<td valign=top bgcolor=gainsboro><u>Supplier</u><BR>Software<BR><BR><span style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & StatusColorArray(SupplierDeliverablesStatus) & "; padding-right: 3px; padding-left: 3px;"">" & StatusNameArray(SupplierDeliverablesStatus) & "</span></td>"
                response.Write "<td valign=top>" & strSupplierDeliverables & "&nbsp;</td>"
                response.Write "</tr>"

                response.Write "<tr>"
                response.Write "<td valign=top bgcolor=gainsboro>Actions<BR><BR><span id=ActionStatus" & trim(rs("ID")) & " style=""width:62px;text-align: center;border-style: solid; border-width: 1px; background-color: " & ActionStatusColorArray(ActionStatus) & "; padding-right: 3px; padding-left: 3px;"">" & ActionNameArray(ActionStatus) & "</span></td>"
                response.Write "<td valign=top colspan=1 onmouseover=""javascript:ActionMouseOver(" & trim(rs("ID")) & ");"" onclick=""javascript:ActionClick(" & trim(rs("ID")) & ");"" id=""ActionRow" & rs("ID") & """>" & strAction & "&nbsp;</td>"
                response.Write "</tr>"
            end if
            'response.Write "</table><table>"
            rs.movenext
        loop
        rs.close
        response.Write "</table>"
      end if
    end if

    If Err.number <> 0 and SI_error <> 1 Then
        response.Write "<table><tr><td style='color:red'>"
        response.Write ("There is a problem displaying the report. Please try again.<br />If problem persists, please contact Pulsar Support.<br /><br />")
        response.write ("Error Number: " & Err.number & "<br />Description: '" & Err.Description & "'<br />" )
        response.Write "</td></tr></table>"
    End If
    
    On Error GoTo 0

%>
</BODY>
</HTML>

<%@ Language=VBScript %>
<%
	  Response.Buffer = True
	  Response.ExpiresAbsolute = Now() - 1
	  Response.Expires = 0
	  Response.CacheControl = "no-cache"
%>

<html>
<head>
<META name=VI60_defaultClientScript content=JavaScript>
<title>Product Readiness Report - Confidential</title>
<STYLE>
td
{
    FONT-FAMILY: verdana;
    FONT-SIZE: xx-small;	
}
A:link,A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
}


function ROW_onmouseover() {
	event.srcElement.style.cursor="hand";

	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;

	if (srcElem.className =="Row")
		srcElem.style.backgroundColor = "Thistle";


}

function ROW_onmouseout() {
	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;

	if (srcElem.className =="Row")
		srcElem.style.backgroundColor = "White";
	
}


function OTSROW_onclick(ID){
	var strResult;
	strResult = window.open("search/ots/Report.asp?txtReportSections=1&txtObservationID=" + ID,"_blank","width=700, height=400,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes") 

}

function DelROW_onclick(ID, RootID){
	var strResult;
	strResult = window.showModalDialog("WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + ID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 

}

//-->
</SCRIPT>
</head>
<%if request("TableOnly") = "1" then%>
    <STYLE>
    .AlertTable TD
    {
	    BACKGROUND-COLOR: white;
    }
    .AlertHeader TD
    {
	    BACKGROUND-COLOR: gainsboro;
    }
    .AlertNone TD
    {
	    BACKGROUND-COLOR: #ebf5db;
    }
    </STYLE>
<%else%>
    <STYLE>
    .AlertTable TD
    {
	    BACKGROUND-COLOR: ivory;
    }
    .AlertHeader TD
    {
	    BACKGROUND-COLOR: beige;
    }
    .AlertNone TD
    {
	    BACKGROUND-COLOR: #ebf5db;
    }
    </STYLE>
<%end if%>
<body LANGUAGE=javascript onload="return window_onload()">
<p align=center>
<font face=verdana size=3>
<%

    dim ProdID
    dim RTMID
    dim VersionID
    dim RootID
    dim strProductName
    dim strDeliverableName
    dim SEPMID
    dim strVersion

    dim strBuildLevel
    dim strDistribution
    dim strCertification
    dim strAvailability
    dim strWorkflow
    dim strDevApproval
    dim SectionArray
    dim strSection
    
    dim OutputArray
    dim RowCount
    dim OutputArraySize
    dim OutputArrayGrow
    
    OutputArraySize = 300
    OutputArrayGrow = 100
    
    strBuildLevel = ""
    strDistribution = ""
    strCertification = ""
    strAvailability = ""
    strWorkflow = ""
    strDevApproval = ""
    
    if trim(request("Sections")) = "" then
        SectionArray = split("1,2,3,4,5,6,7,8,10",",")
    else
        SectionArray = split(request("Sections"),",")
    end if
    
    ProdID = clng(request("ProdID"))
    if clng(request("RTMID"))>0 then
        RTMID = clng(request("RTMID"))
    else
        RTMID=0
    end if
    
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
	
	dim cm
	dim cn
	dim p
	dim rs
	dim i

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
	
	if ProdID <> "" then
		rs.Open "spGetProductVersionName " & ProdID,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strproductName = ""
			SEPMID =  ""
		else
			strproductName = rs("Name") & ""
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
	else
		strproductName = ""
	end if

	if trim(strProductName) = "" then
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to find the selected product.</font>"
		set rs = nothing
		set cn = nothing
	else
	    response.write "<font family=Verdana size=2><b>"
	    if request("TableOnly") <> "1" then
            if request("ReportType") = "2"  then
                response.write strproductname & " Alerts<BR><font size=1></b>(System BIOS Deliverables Only)</font><BR><BR><b>"
            elseif request("ReportType") = "3" then
                response.write strproductname & " Alerts<BR><font size=1></b>(HW Deliverables Only)</font><BR><BR><b>"
            else
                response.write strproductname & " Alerts<BR><font size=1></b>(SW/FW/Doc Deliverables Only)</font><BR><BR><b>"
            end if
	    end if
	    response.write "</b></font></p>"

            if request("ReportType") = "2" then
                rs.open "spListDeliverableAlertDetailsAll " & ProdID & ",2" ,cn,adOpenForwardOnly
            elseif request("ReportType") = "3" then
                if trim(request("TeamID")) = "" then
                    rs.open "spListDeliverableAlertDetailsAll " & ProdID & ",3" ,cn,adOpenForwardOnly
                else    
                    rs.open "spListDeliverableAlertDetailsAll " & ProdID & ",3," & clng(request("TeamID")) ,cn,adOpenForwardOnly
                end if
            else
                rs.open "spListDeliverableAlertDetailsAll " & ProdID ,cn,adOpenForwardOnly
            end if
            do while not rs.eof 
                strVersion = rs("Version") & ""
                if trim(rs("Revision") & "") <> "" then
                    strversion = strVersion &  "," & rs("Revision")
                end if
                if trim(rs("Pass") & "") <> "" then
                    strversion = strVersion &  "," & rs("Pass")
                end if

                if rs("targeted") and  trim(rs("Patch") & "") = "0" and trim(rs("Preinstall") & "") <> "True" and trim(rs("preload") & "") <> "True" and trim(rs("DropInBox") & "") <> "True" and rs("web") & "" <> "True" and trim(rs("SelectiveRestore") & "") <> "True" and trim(rs("arcd") & "") <> "True" and trim(rs("drdvd") & "") <> "True" and trim(rs("racd_EMEA") & "") <> "True" and trim(rs("racd_APD") & "") <> "True" and trim(rs("Racd_Americas") & "") <> "True" and trim(rs("doccd") & "") <> "True" and trim(rs("oscd") & "") <> "True" then
                    strDistribution = strDistribution & "<tr>"
                    strDistribution = strDistribution & "<td>" & rs("ID") & "</td>"
                    strDistribution = strDistribution & "<TD>No&nbsp;Distribution&nbsp;&nbsp;</TD>"
                    strDistribution = strDistribution & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strDistribution = strDistribution & "</tr>"

                end if

                if trim(rs("CertificationStatus") & "") = "0" then
                    strWHQLStatus = "Required"
                elseif trim(rs("CertificationStatus") & "") = "1" then
                    strWHQLStatus = "Submitted"
                elseif trim(rs("CertificationStatus") & "") = "2" then
                    strWHQLStatus = "Approved"
                elseif trim(rs("CertificationStatus") & "") = "3" then
                    strWHQLStatus = "Failed"
                elseif trim(rs("CertificationStatus") & "") = "4" then
                    strWHQLStatus = "Waiver"
                else
                    strWHQLStatus = "Required"
                end if
                
                if trim(rs("LevelID") & "") = "3" or trim(rs("LevelID") & "") = "9" or trim(rs("LevelID") & "") = "10" or trim(rs("LevelID") & "") = "11" then
                    strBuildLevel = strBuildLevel & "<tr>"
                    strBuildLevel = strBuildLevel & "<td>" & rs("ID") & "</td>"
                    strBuildLevel = strBuildLevel & "<TD>Alpha&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "</tr>"
		        end if

                if trim(rs("LevelID") & "") = "4" or trim(rs("LevelID") & "") = "12" or trim(rs("LevelID") & "") = "13" or trim(rs("LevelID") & "") = "14" then
                    strBuildLevel = strBuildLevel & "<tr>"
                    strBuildLevel = strBuildLevel & "<td>" & rs("ID") & "</td>"
                    strBuildLevel = strBuildLevel & "<TD>Beta&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "</tr>"
		        end if

                if trim(rs("LevelID") & "") = "2" then
                    strBuildLevel = strBuildLevel & "<tr>"
                    strBuildLevel = strBuildLevel & "<td>" & rs("ID") & "</td>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>Pre-Production&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "</tr>"
		        elseif request("ReportType") = "3" and trim(rs("LevelID") & "") = "" then
                    strBuildLevel = strBuildLevel & "<tr>"
                    strBuildLevel = strBuildLevel & "<td>" & rs("ID") & "</td>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>Level Not Specified&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strBuildLevel = strBuildLevel & "</tr>"
                end if

                if trim(rs("CertificationStatus") & "") <> "2" and trim(rs("CertificationStatus") & "") <> "4" and trim(rs("CertRequired") & "") = "1" and (trim(rs("LevelID") & "") = "7" or trim(rs("LevelID") & "") = "15" or trim(rs("LevelID") & "") = "16" or trim(rs("LevelID") & "") = "17"  or trim(rs("LevelID") & "") = "18") then 'RC or GM, Requires WHQL, WHQL Status <> 2 or 4
                    strCertification = strCertification & "<tr>"
                    strCertification = strCertification & "<td>" & rs("ID") & "</td>"
                    strCertification = strCertification & "<TD>" & strWHQLStatus & "</TD>"
                    strCertification = strCertification & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strCertification = strCertification & "</tr>"
		        end if

                if isdate(rs("EOLDate")) and clng(rs("ProductStatusID")) < 4  then
    		        if datediff("d",rs("EOLDate"),now) < 365 then
                    strAvailability = strAvailability & "<tr>"
                    strAvailability = strAvailability & "<td>" & rs("ID") & "</td>"
                    strAvailability = strAvailability & "<TD>" & rs("EOLDate") & "</TD>"
                    strAvailability = strAvailability & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strAvailability = strAvailability & "</tr>"
                    end if
		        end if
    		    
		        if trim(rs("DeveloperNotification") & "") = "1" and (trim(rs("DeveloperNotificationStatus") & "") = "0" or trim(rs("DeveloperNotificationStatus") & "") = "2") then
		            
                    strDevApproval = strDevApproval & "<tr >"
                    strDevApproval = strDevApproval & "<td>" & rs("ID") & "</td>"
                    if trim(rs("DeveloperNotificationStatus") & "") = 2 then
                        strDevApproval = strDevApproval & "<TD>Disapproved&nbsp;</TD>"
                    else
                        strDevApproval = strDevApproval & "<TD>Awaiting&nbsp;Approval&nbsp;</TD>"
                    end if
                    strDevApproval = strDevApproval & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strDevApproval = strDevApproval & "</tr>"
		        end if

		        if trim(rs("Location") & "") <> "Workflow Complete" then
                    strWorkflow = strWorkflow & "<tr>"
                    strWorkflow = strWorkflow & "<td>" & rs("ID") & "</td>"
                    strWorkflow = strWorkflow & "<TD>" & replace(replace(rs("Location")& "","Workflow Complete","Complete")," ","&nbsp;") & "</TD>"
                    strWorkflow = strWorkflow & "<TD nowrap>" & rs("name") & "&nbsp;[" & strversion & "]&nbsp;&nbsp;</TD>"
                    strWorkflow = strWorkflow & "</tr>"
    		        
    		    end if
            
                rs.movenext
            loop
            rs.close
%>

        
		<%
		
		
    for each strSection in SectionArray
        select case trim(strSection)
        case 1
       	   ' if request("ReportType") <> "3" or request("TableOnly") = "1" then
       	        if request("TableOnly") <> "1" then
    		        response.Write "<font size=2 face=verdana><B>Build Level Alerts</B></font><BR><font size=1 face=verdana>Deliverables with an Alpha or Beta build level</font>"
                end if
                if strBuildLevel = "" then
    		        strOutput = "<table class=AlertTable ID=BuildLevelAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
    		        response.Write strOutput
           	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (1,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
	                end if
		        else
    		        strOutput = "<table class=AlertTable ID=BuildLevelAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Build&nbsp;Level&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strBuildLevel & "</Table><BR>"
    		        response.Write strOutput
           	        if request("TableOnly") = "1"  and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (1,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
	                end if
		        end if
          '  end if
	    case 2
	        if (request("ReportType") <> "2" and request("ReportType") <> "3") or request("TableOnly") = "1" then
       	        if request("TableOnly") <> "1" then
        	        response.Write "<font size=2 face=verdana><B>Distribution Alerts</B></font><BR><font size=1 face=verdana>Deliverables that are targeted but have no distributions defined.</font>"
                end if
                if strDistribution = "" then
                    strOutput = "<table class=AlertTable ID=DistributionAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
        		    response.Write strOutput
           	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (2,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                    end if
		        else
        		    strOutput = "<table class=AlertTable ID=DistributionAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Alert&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strDistribution & "</Table><BR>"
        		    response.Write strOutput
           	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (2,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                    end if
		        end if
            end if
        case 3		
	        if (request("ReportType") <> "2" and request("ReportType") <> "3") or request("TableOnly") = "1" then
       	        if request("TableOnly") <> "1" then
    		        response.Write "<font size=2 face=verdana><B>Certification Alerts</B></font><BR><font size=1 face=verdana>Deliverables with a WHQL status other than ""Approved"" or ""Waiver"".</font>"
                end if
                if strCertification = "" then
    		        strOutput = "<table class=AlertTable ID=CertificationAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
    		        response.Write strOutput
           	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (3,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                    end if
		        else
    		        strOutput =  "<table class=AlertTable ID=CertificationAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>WHQL&nbsp;Status&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strCertification & "</Table><BR>"
    		        response.Write strOutput
           	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                    	RecordsUpdated = SaveAlertToDB (3,CurrentUserID,ProdID,strOutput,RTMID)
    	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                    end if
                end if
            end if
        case 4        
       	    if request("TableOnly") <> "1" then
    		    response.Write "<font size=2 face=verdana><B>Workflow Alerts</B></font><BR><font size=1 face=verdana>Deliverables that are not ""Workflow Complete"".</font>"
            end if
            if strWorkflow = "" then
    		    strOutput = "<table class=AlertTable ID=WorkflowAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (4,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
		    else
    		    strOutput =  "<table class=AlertTable ID=WorkflowAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Workflow&nbsp;Step&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strWorkflow & "</Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (4,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
            end if		
        case 5        
       	    if request("TableOnly") <> "1" then
    		    response.Write "<font size=2 face=verdana><B>Availability Alerts</B></font><BR><font size=1 face=verdana>Deliverables that can not be used after the reported date.</font>"
            end if
            if strAvailability = "" then
    		    strOutput = "<table class=AlertTable ID=AvailabilityAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (5,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
		    else
    		    strOutput =  "<table class=AlertTable ID=AvailabilityAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Use&nbsp;Until&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strAvailability & "</Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (5,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
            end if
        case 6       
       	    if request("TableOnly") <> "1" then
    		    response.Write "<font size=2 face=verdana><B>Developer Alerts</B></font><BR><font size=1 face=verdana>Deliverables that have not been approved by the development team.</font>"
            end if
            if strDevApproval = "" then
    		    strOutput = "<table class=AlertTable ID=DevApprovalAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (6,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
		    else
    		    strOutput =  "<table class=AlertTable ID=DevApprovalAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Dev.&nbsp;Approval&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>" & strDevApproval & "</Table><BR>"
   		        response.Write strOutput
       	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
                  	RecordsUpdated = SaveAlertToDB (6,CurrentUserID,ProdID,strOutput,RTMID)
   	                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
                end if
            end if
                		
		case 7 'Include roots with no versions?
       	    if request("TableOnly") <> "1" then
                response.Write "<font size=2 face=verdana><B>Root Deliverable Alerts</B></font><BR><font size=1 face=verdana>Root deliverables that have no versions targeted.</font>"
	        end if
	        if request("ReportType") = "2" then
		        rs.open "spListRootsWithNoTargetedVersions2 " & ProdID & ",2", cn,adOpenForwardOnly
	        elseif request("ReportType") = "3" then
                if trim(request("TeamID")) = "" then
    		        rs.open "spListRootsWithNoTargetedVersions2 " & ProdID & ",3", cn,adOpenForwardOnly
                else
    		        rs.open "spListRootsWithNoTargetedVersions2 " & ProdID & ",3," & clng(trim(request("TeamID"))), cn,adOpenForwardOnly
	            end if
	        else
		        rs.open "spListRootsWithNoTargetedVersions2 " & ProdID, cn,adOpenForwardOnly
		    end if

            Redim OutputArray(OutputArraySize)
            RowCount=0

		    if rs.eof and rs.bof then
		        OutputArray(RowCount) = "<table class=AlertTable ID=RootAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table>"
		    else
                OutputArray(RowCount) = "<table class=AlertTable ID=RootAlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertHeader><TD><b>ID&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD><b>Alert&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Deliverable</b></TD></TR>"
		    end if
            RowCount=1
		    do while not rs.eof
		        strOutput = "<TR><td>" & rs("id") & "</td>"
		        strOutput = strOutput &  "<td>No&nbsp;Versions&nbsp;Targeted&nbsp;</td>"
		        strOutput = strOutput & "<td>" & rs("name") & "</td></tr>"
                If RowCount > UBound(OutputArray) Then
			        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow + 1)
        		End If
                OutputArray(RowCount)  =  strOutput
                RowCount = RowCount +1 
		        rs.movenext 
		    loop
		    rs.close
                If RowCount > UBound(OutputArray) Then
			        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        		End If

            OutputArray(RowCount) = "</table><BR>"

	        response.Write join(OutputArray,"")
   	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
              	RecordsUpdated = SaveAlertToDB2 (7,CurrentUserID,ProdID,OutputArray,RTMID)
                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
            end if

		case 8 ' OTS on targeted products
            rs.open "select dbo.ufn_IsLinkedServerEnabled('Housireport01') as value", cn
            if not (rs.eof and rs.bof) then
                IsLinkedServerEnabled = trim(rs("value"))
            end if
            rs.close

            Redim OutputArray(OutputArraySize)
            RowCount=0

            if IsLinkedServerEnabled = "False" then
                OutputArray(RowCount) = "<table border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><td>Warning: Pulsar is unable to establish connection with Nebula</td></tr></table>"
                RowCount=1  
            else
                if request("ReportType") = "2" then    
	               rs.open "spListOTSTargeted4Product " &  ProdID & ",2" , cn,adOpenForwardOnly
                elseif request("ReportType") = "3" then    
                    if trim(request("TeamID")) = "" then
                        rs.open "spListOTSTargeted4Product " &  ProdID & ",3", cn,adOpenForwardOnly
                    else
                        rs.open "spListOTSTargeted4Product " &  ProdID & ",3," & clng(trim(request("TeamID"))) , cn,adOpenForwardOnly
	                end if
	            else
	               rs.open "spListOTSTargeted4Product " &  ProdID  , cn,adOpenForwardOnly
	            end if
       	        if request("TableOnly") <> "1" then
		        %>
                    <font face=verdana size=2><b>OTS Alerts - <%=strproductName%> Primary</b><br></font>
                    <font size=1 face=verdana>Open P0/P1 Observations written against <%=strproductName%> on any version of the targeted deliverables.</font><br></font>
                <%
                end if

                if not (rs.eof and rs.bof) then
                    OutputArray(RowCount) = "<table class=AlertTable width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro><tr class=AlertHeader><td><b>OTS&nbsp;ID</b></td><td><b>Deliverable</b></td><td><b>Pr</b></td><td><b>State</b></td><td><b>Milestone</b></td><td><b>Owner</b></td><td width=100><b>Summary</b></td></tr>"
                else
    		        OutputArray(RowCount) = "<table border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table>"
                end if
		        RowCount=1
		        do while not rs.eof
		            strOutput = "<tr><td><a target=""_blank"" href=""search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a></td>"
                    strOutput = strOutput & "<td width=""50%"">" & rs("OTSdeliverable") & " [" & rs("OTSComponentVersion") & "]</td>"
	                strOutput = strOutput & "<td>" & rs("Priority") & "</td>"
	                strOutput = strOutput & "<td>" & rs("State") & "</td>"
	                strOutput = strOutput & "<td>" & rs("GatingMilestone") & "&nbsp;</td>"
	                strOutput = strOutput & "<td nowrap>" & replace(replace(replace(rs("OwnerName") & "","VENDOR- ",""),"VENODR- ",""),"VENDOR - ","") & "</td>"
	                strOutput = strOutput & "<td width=""50%"">" & rs("Summary") & "</td></tr>"
                    If RowCount > UBound(OutputArray) Then
			            ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow + 1)
        	    	End If
                    OutputArray(RowCount)  =  strOutput
                    RowCount = RowCount +1 
                    rs.movenext
                loop
                rs.close
            end if

            If RowCount > UBound(OutputArray) Then
                ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
            End If

            OutputArray(RowCount) = "</table><br>"
            response.Write join(OutputArray,"")

            response.write "<font size=1 face=verdana>Observations Displayed: " & RowCount-1 & "<BR><BR><BR></font>"
   	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
              	RecordsUpdated = SaveAlertToDB2 (8,CurrentUserID,ProdID,OutputArray,RTMID)
                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
            end if

        case 9 'OTS on Affected products
            if request("ReportType") = "2" then
	            rs.open "spListOTS4OtherProductsRelated " &  ProdID & ",2", cn,adOpenForwardOnly
	        elseif request("ReportType") = "3" then
                if trim(request("TeamID")) = "" then
    	            rs.open "spListOTS4OtherProductsRelated " &  ProdID & ",3" , cn,adOpenForwardOnly
                else
    	            rs.open "spListOTS4OtherProductsRelated " &  ProdID & ",3," & clng(trim(request("TeamID"))) , cn,adOpenForwardOnly
	            end if
	        else
	            rs.open "spListOTS4OtherProductsRelated " &  ProdID , cn,adOpenForwardOnly
	        end if
       	    if request("TableOnly") <> "1" then
%>		    
                <font face=verdana size=2><b>OTS Alerts - Related Products</b><br></font>
                <font size=1 face=verdana>Open P0/P1 Observations written against other products (real products in same division) on any version of the deliverables targeted on <%=strproductName%> that are 'Untested*','Test Required','Waiver Requested', or 'Affected' on <%=strProductname%>.</font><br></font>
 <%         end if
 
            Redim OutputArray(OutputArraySize)
            RowCount=0
            
            if not (rs.eof and rs.bof) then
                OutputArray(RowCount) = "<table class=AlertTable width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro><tr class=AlertHeader><td><b>OTS&nbsp;ID</b></td><td><b>Product</b></td><td><b>Deliverable</b></td><td><b>Pr</b></td><td nowrap><b>" & strProductname & "</b></td><td><b>State</b></td><td><b>Milestone</b></td><td><b>Owner</b></td><td width=100><b>Summary</b></td></tr>"
            else
    		    OutputArray(RowCount) = "<table border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table>"
            end if
		    RowCount=1
		    do while not rs.eof
		        strOutput  = "<tr><td><a target=""_blank"" href=""search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a></td>"
                strOutput = strOutput & "<td nowrap>" & rs("Product") & "</td>"
                strOutput = strOutput & "<td width=""50%"">" & rs("OTSdeliverable") & " [" & rs("OTSComponentVersion") & "]</td>"
	            strOutput = strOutput & "<td>" & rs("Priority") & "</td>"
	            strOutput = strOutput & "<td>" & rs("AffectedState") & "</td>"
	            strOutput = strOutput & "<td>" & rs("State") & "</td>"
	            strOutput = strOutput & "<td>" & rs("GatingMilestone") & "&nbsp;</td>"
	            strOutput = strOutput & "<td nowrap>" & replace(replace(replace(rs("OwnerName") & "","VENDOR- ",""),"VENODR- ",""),"VENDOR - ","") & "</td>"
	            strOutput = strOutput & "<td width=""50%"">" & rs("Summary") & "</td></tr>"
                If RowCount > UBound(OutputArray) Then
			        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        		End If
                OutputArray(RowCount)  =  strOutput
                RowCount = RowCount +1 
                rs.movenext
            loop
            rs.close

            If RowCount > UBound(OutputArray) Then
    	        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
       		End If
            OutputArray(RowCount) = "</table><br>"
            response.Write join(OutputArray,"")
            response.write "<font size=1 face=verdana>Observations Displayed: " & RowCount-1 & "<BR><BR><BR></font>"
   	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
              	RecordsUpdated = SaveAlertToDB2 (9,CurrentUserID,ProdID,OutputArray,RTMID)
                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
           end if

        case 10 'OTS Count Summary
		    if request("Priority") <> "" or request("AffectedState") <> "" then
                '--------Show an expanded OTS section


	            set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                cm.commandtimeout=120
	            cm.CommandType = 4
            	cm.CommandText = "spListOTS4OtherProductsRelated2"
	            
                Set p = cm.CreateParameter("@ProdID", 3, &H0001)
	            p.Value = ProdID
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@ReportType",3, &H0001)
                if trim(request("ReportType")) = "2" then
                    p.Value = 2
                elseif trim(request("ReportType")) = "3" then
                    p.Value = 3
                else
                    p.Value = 1
                end if
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@TeamID",3, &H0001)
                if trim(request("TeamID")) = "" then
                    p.Value = null
                else
                    p.Value = clng(request("TeamID"))
                end if
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Priority",3, &H0001)
                if trim(request("Priority")) = "" then
                    p.Value = null
                else
                    p.Value = clng(request("Priority"))
                end if
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@AffectedState",200, &H0001,50)
                if trim(request("AffectedState")) = "" then
                    p.Value = null
                else
                    p.Value = request("AffectedState")
                end if
	            cm.Parameters.Append p

	            rs.CursorType = adOpenForwardOnly
	            rs.LockType=AdLockReadOnly
	            Set rs = cm.Execute 

	            set cm=nothing

       	        if request("TableOnly") <> "1" then
                    if request("Priority") = "1" then
                        strPriority = "<b>P0/P1</b> "
                    elseif request("Priority") = "2" then
                        strPriority = "<b>P2</b> "
                    elseif request("Priority") = "3" then
                        strPriority = "<b>P3/P4/P5</b> "
                    else
                        strPriority = ""
                    end if
                    if request("AffectedState") = "" then
                        strAffectedState=""
                    else
                        strAffectedState=" that are <b>" & server.HTMLEncode(request("AffectedState")) & "</b> on " & strProductname
                    end if
    %>		    
                    <font face=verdana size=2><b>OTS Alerts - Related Products</b><br></font>
                    <font size=1 face=verdana>Open <%=strPriority%>Observations written against other products (real products in same division) on any version of the deliverables targeted on <%=strproductName%><%=strAffectedState%>.</font><br><br>
     <%         end if
 
                Redim OutputArray(OutputArraySize)
                RowCount=0
            
                if not (rs.eof and rs.bof) then
                    OutputArray(RowCount) = "<table class=AlertTable width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro><tr class=AlertHeader><td><b>OTS&nbsp;ID</b></td><td><b>Product</b></td><td><b>Deliverable</b></td><td><b>Pr</b></td><td nowrap><b>" & strProductname & "</b></td><td><b>State</b></td><td><b>Milestone</b></td><td><b>Owner</b></td><td width=100><b>Summary</b></td></tr>"
                else
    		        OutputArray(RowCount) = "<table border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table>"
                end if
		        RowCount=1
		        do while not rs.eof
		            strOutput  = "<tr><td><a target=""_blank"" href=""search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a></td>"
                    strOutput = strOutput & "<td nowrap>" & rs("Product") & "</td>"
                    strOutput = strOutput & "<td width=""50%"">" & rs("OTSdeliverable") & " [" & rs("OTSComponentVersion") & "]</td>"
	                strOutput = strOutput & "<td>" & rs("Priority") & "</td>"
	                strOutput = strOutput & "<td>" & rs("AffectedState") & "</td>"
	                strOutput = strOutput & "<td>" & rs("State") & "</td>"
    	            strOutput = strOutput & "<td>" & rs("GatingMilestone") & "&nbsp;</td>"
	                strOutput = strOutput & "<td nowrap>" & replace(replace(replace(rs("OwnerName") & "","VENDOR- ",""),"VENODR- ",""),"VENDOR - ","") & "</td>"
	                strOutput = strOutput & "<td width=""50%"">" & rs("Summary") & "</td></tr>"
                    If RowCount > UBound(OutputArray) Then
			            ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        		    End If
                    OutputArray(RowCount)  =  strOutput
                    RowCount = RowCount +1 
                    rs.movenext
                loop
                rs.close

                If RowCount > UBound(OutputArray) Then
			        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        		End If

                OutputArray(RowCount) = "</table><br>"
                response.Write join(OutputArray,"")
                response.write "<font size=1 face=verdana>Observations Displayed: " & RowCount-1 & "<BR><BR><BR></font>"
   	            if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
              	    RecordsUpdated = SaveAlertToDB2 (9,CurrentUserID,ProdID,OutputArray,RTMID)
                    response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
               end if






            else '--------Show table of OTS counts
	            if request("ReportType") = "2" then
	                rs.open "spCountOTS4OtherProductsRelated " &  ProdID & ",2", cn,adOpenForwardOnly
	            elseif request("ReportType") = "3" then
                    if trim(request("TeamID")) = "" then
    	                rs.open "spCountOTS4OtherProductsRelated " &  ProdID & ",3" , cn,adOpenForwardOnly
                    else
    	                rs.open "spCountOTS4OtherProductsRelated " &  ProdID & ",3," & clng(trim(request("TeamID"))) , cn,adOpenForwardOnly
	                end if
	            else
	                rs.open "spCountOTS4OtherProductsRelated " &  ProdID , cn,adOpenForwardOnly
	            end if
       	        if request("TableOnly") <> "1" then
    %>		    
                    <font face=verdana size=2><b>OTS Alerts - Related Products</b><br></font>
                    <font size=1 face=verdana>Open Observations written against other products (real products in same division) on any version of the deliverables targeted on <%=strproductName%>.</font><br><br>
     <%         end if

                dim strLastState
                dim StatePriorityArray
                dim StateTotalsArray
                strLastState = ""
                StatePriorityArray = split(",0,0,0,0",",")
                StateTotalsArray = split(",0,0,0,0",",")
                if not (rs.eof and rs.bof) then
                    response.write "<table cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro bgcolor=ivory>"
                    response.write "<tr bgcolor=beige>" 
                    response.write "<td><b>Affected&nbsp;State&nbsp;&nbsp;&nbsp;</b></td>" 
                    response.write "<td><b>&nbsp;&nbsp;P0/P1&nbsp;&nbsp;</b></td>" 
                    response.write "<td><b>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>" 
                    response.write "<td><b>&nbsp;&nbsp;P3/P4/P5&nbsp;&nbsp;</b></td>" 
                    response.write "<td><b>&nbsp;&nbsp;Total&nbsp;&nbsp;</b></td>" 
                    response.write "</tr>" 
                end if
                do while not rs.EOF
                    if strLastState <> rs("AffectedState") and strLastState <> "" then
                        response.write "<tr><td>" & strLastState & "&nbsp;&nbsp;</td>"
                        for i  = 1 to 3
                            response.write "<td align=center>" & MakeSection10Link(StatePriorityArray(i),i,strLastState) & "</td>"
                            StateTotalsArray(i) = clng(StateTotalsArray(i)) + clng(StatePriorityArray(i))
                        next
                        response.write "<td align=center>" & MakeSection10Link(clng(StatePriorityArray(1)) + clng(StatePriorityArray(2)) + clng(StatePriorityArray(3)),4,strLastState) & "</td>"
                        StateTotalsArray(4) = clng(StateTotalsArray(4))  + clng(StatePriorityArray(1)) + clng(StatePriorityArray(2)) + clng(StatePriorityArray(3))
                        for i  = 1 to 3
                            StatePriorityArray(i) = "0"
                        next
                        response.write "</tr>"
                    end if
                    strLastState = rs("AffectedState")  & ""
                    StatePriorityArray(rs("Priority")) = rs("OTSCount")
                    rs.MoveNext
                loop
                if strLastState <> "" then
                    response.write "<tr><td>" & strLastState & "&nbsp;&nbsp;&nbsp;</td>"
                    for i  = 1 to 3
                        response.write "<td align=center>" & MakeSection10Link(StatePriorityArray(i),i,strLastState) & "</a></td>"
                        StateTotalsArray(i) = clng(StateTotalsArray(i)) + clng(StatePriorityArray(i))
                    next
                    response.write "<td align=center>" & MakeSection10Link(clng(StatePriorityArray(1)) + clng(StatePriorityArray(2)) + clng(StatePriorityArray(3)),4,strLastState) & "</a></td></tr>"
                    StateTotalsArray(4) = clng(StateTotalsArray(4))  + clng(StatePriorityArray(1)) + clng(StatePriorityArray(2)) + clng(StatePriorityArray(3))

                    response.write "<tr bgcolor=lavender><td>Total&nbsp;&nbsp;&nbsp;</td>"
                    for i  = 1 to 4
                        response.write "<td align=center>" & MakeSection10Link(StateTotalsArray(i),i,strLastState) & "</td>"
                    next
                    response.write "</tr>"
                end if
                if not (rs.eof and rs.bof) then
                    response.write "</table>"
                end if
                rs.close

 
            end if
        case 11 'OTS Primary and Affected products combined
            if request("ReportType") = "2" then
	            rs.open "spListOTS4OtherProductsRelated3 " &  ProdID & ",2", cn,adOpenForwardOnly
	        elseif request("ReportType") = "3" then
                if trim(request("TeamID")) = "" then
    	            rs.open "spListOTS4OtherProductsRelated3 " &  ProdID & ",3" , cn,adOpenForwardOnly
                else
    	            rs.open "spListOTS4OtherProductsRelated3 " &  ProdID & ",3," & clng(trim(request("TeamID"))) , cn,adOpenForwardOnly
	            end if
	        else
	            rs.open "spListOTS4OtherProductsRelated3 " &  ProdID , cn,adOpenForwardOnly
	        end if
       	    if request("TableOnly") <> "1" then
%>		    
                <font face=verdana size=2><b>OTS Alerts - Primary and Related Products</b><br></font>
                <font size=1 face=verdana>Open P0/P1 Observations written against other products (real products in same division) on any version of the deliverables targeted on <%=strproductName%> that are 'Untested*','Test Required','Waiver Requested', or 'Affected' on <%=strProductname%>.</font><br></font>
 <%         end if
 
            Redim OutputArray(OutputArraySize)
            RowCount=0
            
            if not (rs.eof and rs.bof) then
                OutputArray(RowCount) = "<table class=AlertTable width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro><tr class=AlertHeader><td><b>OTS&nbsp;ID</b></td><td><b>Product</b></td><td><b>Deliverable</b></td><td><b>Pr</b></td><td nowrap><b>" & strProductname & "</b></td><td><b>State</b></td><td><b>Milestone</b></td><td><b>Owner</b></td><td width=100><b>Summary</b></td></tr>"
            else
    		    OutputArray(RowCount) = "<table border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr class=AlertNone><TD>None Found</td></tr></Table>"
            end if
		    RowCount=1
		    do while not rs.eof
		        strOutput  = "<tr><td><a target=""_blank"" href=""search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a></td>"
                strOutput = strOutput & "<td nowrap>" & rs("Product") & "</td>"
                strOutput = strOutput & "<td width=""50%"">" & rs("OTSdeliverable") & " [" & rs("OTSComponentVersion") & "]</td>"
	            strOutput = strOutput & "<td>" & rs("Priority") & "</td>"
	            strOutput = strOutput & "<td>" & rs("AffectedState") & "</td>"
	            strOutput = strOutput & "<td>" & rs("State") & "</td>"
	            strOutput = strOutput & "<td>" & rs("GatingMilestone") & "&nbsp;</td>"
	            strOutput = strOutput & "<td nowrap>" & replace(replace(replace(rs("OwnerName") & "","VENDOR- ",""),"VENODR- ",""),"VENDOR - ","") & "</td>"
	            strOutput = strOutput & "<td width=""50%"">" & rs("Summary") & "</td></tr>"
                If RowCount > UBound(OutputArray) Then
			        ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        		End If
                OutputArray(RowCount)  =  strOutput
                RowCount = RowCount +1 
                rs.movenext
            loop
            rs.close
            If RowCount > UBound(OutputArray) Then
			    ReDim Preserve OutputArray(UBound(OutputArray) + OutputArrayGrow +1)
        	end If
            OutputArray(RowCount) = "</table><br>"
            response.Write join(OutputArray,"")
            response.write "<font size=1 face=verdana>Observations Displayed: " & RowCount-1 & "<BR><BR><BR></font>"
   	        if request("TableOnly") = "1" and request("RTMSignoff") = "1" then
              	RecordsUpdated = SaveAlertToDB2 (9,CurrentUserID,ProdID,OutputArray,RTMID)
                response.Write "<label id=RecordID style=""Display:none"">" & RecordsUpdated & "</label>"
           end if

        end Select
        next
        
	    set rs = nothing
	    set cn = nothing

  	    if request("TableOnly") <> "1" then
%>

        <br>
        <br>
        <font face=verdana size="1">Report Generated <%=formatdatetime(date(),vblongdate) %></font>
        <br>
        <br>
        <font Size="2" Color="red"><p><strong>HP&nbsp;Confidential</strong></p></font>
<% 
        end if
 end if


function MakeSection10Link(strID, Priority, State)
    dim strState 
    strState = ""

    if State  <> "" then
        strState = "AffectedState=" & State & "&"
    end if

    if strID = 0 then
        MakeSection10Link = strID
    elseif Priority=4 then
        MakeSection10Link = "<a target=_blank href=""ReadinessReport.asp?Sections=10&" & strState & "ProdID=" & request("ProdID") & """>" & strID & "</a>"
    else
        MakeSection10Link = "<a target=_blank href=""ReadinessReport.asp?Sections=10&" & strState & "Priority=" & Priority & "&ProdID=" & request("ProdID") & """>" & strID & "</a>"
    end if
end function

function SaveAlertToDB(strSectionID, strUserID, strProductID, strAlertHTML,strRTMID)
    dim recordsupdated
    

    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.NamedParameters = True
	cm.CommandText = "spAddProductRTMAlert"
                    
    Set p = cm.CreateParameter("@ReportSectionID",3, &H0001)
	p.Value = clng(strSectionID)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@UserID",3, &H0001)
	p.Value = clng(strUserID)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@ProductID", 3, &H0001)
    p.Value = clng(strProductID)
    cm.Parameters.Append p

  	Set p = cm.CreateParameter("@AlertHTML", 201, &H0001, 2147483647)
    p.Value = strAlertHTML
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@RTMID", 3, &H0001)
    p.Value = clng(strRTMID)
    cm.Parameters.Append p
 
    Set p = cm.CreateParameter("@NewID", 3, &H0002)
    cm.Parameters.Append p

    cm.Execute recordsupdated
    
    if recordsupdated <> 1 then
        SaveAlertToDB = 0
    else
        SaveAlertToDB = cm("@NewID")
    end if
   	set cm=nothing
   	
    
end function

function SaveAlertToDB2(strSectionID, strUserID, strProductID, AlertArray,strRTMID)
    dim recordsupdated
    

    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.NamedParameters = True
	cm.CommandText = "spAddProductRTMAlert"
                    
    Set p = cm.CreateParameter("@ReportSectionID",3, &H0001)
	p.Value = clng(strSectionID)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@UserID",3, &H0001)
	p.Value = clng(strUserID)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@ProductID", 3, &H0001)
    p.Value = clng(strProductID)
    cm.Parameters.Append p

  	Set p = cm.CreateParameter("@AlertHTML", 201, &H0001, 2147483647)
    p.Value = join(alertarray,"")
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@RTMID", 3, &H0001)
    p.Value = clng(strRTMID)
    cm.Parameters.Append p
 
    Set p = cm.CreateParameter("@NewID", 3, &H0002)
    cm.Parameters.Append p

    cm.Execute recordsupdated
    
    if recordsupdated <> 1 then
        SaveAlertToDB2 = 0
    else
        SaveAlertToDB2 = cm("@NewID")
    end if
    
   	set cm=nothing
   	
    
end function
%>
</font>
</body>
</html>

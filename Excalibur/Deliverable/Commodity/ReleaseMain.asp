<%@ Language="VBScript" %>
<%Option Explicit %>
<html>
<head>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onload() {
//    if (typeof(frmMain.txtEmail.length) == "undefined")
//        frmMain.txtEmail.value = frmMain.txtEmail.tag;
//    else
//        for (i=0;i<frmMain.txtEmail.length;i++)
//            frmMain.txtEmail[i].value = frmMain.txtEmail[i].tag;
}
//-->
</SCRIPT>
</head>
<STYLE>
TD{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
}
BODY{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
}
H1{
    FONT-FAMILY: Verdana;
    FONT-SIZE: small;
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<% if request("txtFunction") = "2" then %>
    <H1>Cancel Deliverable Versions</H1>
<%else%>
    <H1>Release Deliverables Versions</H1>
<%end if%>
<form id=frmMain method=post action="ReleaseSave.asp">

<%

	dim cn
	dim rs
	dim rs2
	dim strSQL
	dim strVersion
	dim strNotify
	dim strFromMilestone
	dim strToMilestone
	dim strFromMilestoneID
	dim strToMilestoneID
	dim MilestoneCount
	dim strTesterEmail
    dim strExecutionEngineerID
    dim strExecutionEngineerEmail
    dim strExecutionEngineerName
    dim strTestLeadList
    dim strPMEmailList
    dim strDevManagerEmail
    dim strDeveloperEmail

  
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    strSQL= "Select e2.email as DevManagerEmail, e1.email as DeveloperEmail, r.notifications as AlsoNotify, v.hfcn, r.testerid, r.developerid, r.devmanagerid, v.tts, v.wwanfailureconfirmed,c.RequiresTTS,v.CommercialReleaseStatus,v.ConsumerReleaseStatus,v.imagepath, r.typeid, v.id, v.deliverablename, v.Version, v.Revision, v.pass, v.partnumber, v.modelnumber, vd.name as vendor " & _
            "From DeliverableVersion v with (NOLOCK), vendor vd with (NOLOCK), deliverableroot r with (NOLOCK), deliverablecategory c with (NOLOCK), employee e1 with (NOLOCK), employee e2 with (NOLOCK) " & _
            "Where vd.id = v.vendorid " & _
            "and v.deliverablerootid = r.id " & _
            "and r.categoryid = c.id " & _
            "and e1.id = v.Developerid " & _
            "and e2.id = r.devmanagerid " & _
            "and v.id in (" & scrubsql(request("VersionID")) & ")"


            
    rs.Open strSQL,cn,adOpenKeyset
    if not (rs.EOF and rs.BOF) then
        response.Write "<TABLE width=100% border=1 cellpadding=2 cellspacing=0><TR bgcolor=beige>"    
        response.Write "<TD>&nbsp;</TD>"
        response.Write "<TD><b>ID</b></TD>"
        response.Write "<TD><b>Name</b></TD>"
        response.Write "<TD><b>Version</b></TD>"
        response.Write "<TD><b>Part</b></TD>"
        response.Write "<TD><b>Model</b></TD>"
        response.Write "<TD><b>Next&nbsp;Step</b></TD>"
        response.Write "</TR>"
    end if
    do while not rs.EOF
        strFromMilestone = ""
        strToMilestone = ""
        strDeveloperEmail = rs("DeveloperEmail") & ""
        strDevManagerEmail = rs("DevManagerEmail") & ""
        'Get Workflow Step and notifications	
    	set rs2 = server.CreateObject("ADODB.recordset")
	    rs2.Open "spGetWorkflowStepsInProgress " & rs("ID"),cn,adOpenForwardOnly
    	MilestoneCount  = 0
  		strFromMilestoneID=0
	    strToMilestoneID=0

    	strNotify = ""
        if rs2.EOF and rs2.BOF then
            response.write "<TR bgcolor=gainsboro><TD colspan=7>Deliverable " & rs("ID") & " can not be processed on this page because it has completed its workflow.</td></TR>"
    	elseif lcase(trim(rs("wwanfailureconfirmed") & "")) = "false" and lcase(trim(rs("TTS"))) = "failed" and rs("RequiresTTS") then
            response.write "<TR bgcolor=gainsboro><TD colspan=7>Deliverable " & rs("ID") & " can not be released to the next workflow step until it is reviewed by the WWAN Engineers because it failed TTS.</td></TR>"
    	else
    		do while not rs2.EOF
		        if rs2("ReportMilestone") <= 2 or rs2("ReportMilestone")=6 then 'no one should release from reportmilestone 3 (release team)
    				strFromMilestone = rs2("Milestone")
	                strFromMilestoneID = rs2("ID")
    				
    				strNotify = rs2("Notify") & ""
			    	if rs("ConsumerReleaseStatus") > 0  and trim(rs2("NotifyConsumerSpecific") & "") <> "" then
				    	strNotify = strNotify & ";" & rs2("NotifyConsumerSpecific")
				    end if
				    if rs("CommercialReleaseStatus") > 0 and trim(rs2("NotifyCommercialSpecific") & "") <> "" then
					    strNotify = strNotify & ";" & rs2("NotifyCommercialSpecific")
				    end if
				    MilestoneCount = MilestoneCount + 1
			    end if
			    rs2.Movenext
    	    loop
    	    if MilestoneCount <> 1 then
                response.write "<TR bgcolor=gainsboro><TD colspan=7>Deliverable " & rs("ID") & " can not be released because Excalibur can not determine the current workflow step.</td></TR>"
            end if
        end if
        rs2.Close
        set rs2 = nothing

        if strFromMilestoneID <> "" and strFromMilestoneID <> "0" then
        	set rs2 = server.CreateObject("ADODB.recordset")
    		rs2.Open "spGetNextMilestone " & rs("ID") & "," & clng(strFromMilestoneID),cn,adOpenForwardOnly
	    	if rs2.EOF and rs2.BOF then
		    	strToMilestoneID = 0		
			    strToMilestone = "Complete"
		    else
			    strToMilestoneID = rs2("ID")		
			    strToMilestone = rs2("Milestone")
		    end if
		    rs2.Close
            set rs2 = nothing
        else
	    	strToMilestoneID = 0		
		    strToMilestone = "Complete"
        end if


        if strToMilestone = "" then
            strToMilestone = "Complete"
        end if
        if trim(rs("Typeid")) <> "1" then
            response.write "<TR bgcolor=gainsboro><TD colspan=7>Deliverable " & rs("ID") & " can not be processed on this page because it is not a hardware deliverable.</td></TR>"
        elseif MilestoneCount = 1 then
            'Lookup Tester Email
            strTesterEmail = ""
	        if trim(rs("TesterID") & "" ) <> "" and trim(rs("TesterID") & "" ) then
            	set rs2 = server.CreateObject("ADODB.recordset")
		        rs2.Open "spGetEmployeeByID " & clng(rs("TesterID")),cn,adOpenStatic
		        if not (rs2.EOF and rs2.BOF) then
			        strTesterEmail = rs2("Email") & ""
		        end if
		        rs2.Close
	            set rs2 = nothing
	        end if        
            if strTesterEmail <> "" then
                strNotify = strNotify & ";" & strTesterEmail 
            end if
            
            if lcase(trim(strFromMilestone)) = "core team" then
            	set rs2 = server.CreateObject("ADODB.recordset")
                rs2.open "spGetExecutionEngineer "  & rs("ID"),cn,adOpenForwardOnly
                if rs2.eof and rs2.bof then
    	            strExecutionEngineerID = 0
	                strExecutionEngineerEmail = ""
	                strExecutionEngineerName= ""
                else
    	            strExecutionEngineerID = trim(rs2("ID") & "")
	                strExecutionEngineerEmail = rs2("email") & ""
	                strExecutionEngineerName= rs2("name") & ""
                end if
                rs2.close
                set rs2 = nothing
            else
  	            strExecutionEngineerID = 0
                strExecutionEngineerEmail = ""
                strExecutionEngineerName= ""
            end if            
            
            'Get PM Email addresses            
           	set rs2 = server.CreateObject("ADODB.recordset")
		    rs2.Open "spListCommodityPMs4Version " &  rs("ID"),cn,adOpenForwardOnly
		    strPMEmailList = ""
		    do while not rs2.EOF
    			strPMEmailList = strPMEmailList & ";" & rs2("Email")
			    rs2.MoveNext
		    loop
		    rs2.Close
		    set rs2 = nothing
		    if strPMEmailList <> "" then
    			strPMEmailList = mid(strPMEmailList,2)
		    end if
		
            'Get Test Lead Email addresses            
           	set rs2 = server.CreateObject("ADODB.recordset")
		    rs2.Open "spListTestLeads4Version " &  rs("ID"),cn,adOpenForwardOnly
		    strTestLeadList = ""
		    do while not rs2.EOF
    			strTestLeadList = strTestLeadList & ";" & rs2("Email")
			    rs2.MoveNext
		    loop
		    rs2.Close
		    set rs2 = nothing
		    if strTestLeadList <> "" then
    			strTestLeadList = mid(strTestLeadList,2)
		    end if		
            
            'Built Full Notification List
		    if left(trim(strNotify),1) = ";" then
			    strNotify = mid(strNotify,2)
		    end if
		    if rs("HFCN") and strToMilestone = "Complete" then
			    if trim(strNotify) = "" then
    				strNotify = "NBSCSWEngrs@hp.com"
	    		else
		    		strNotify = "NBSCSWEngrs@hp.com;" & strNotify
			    end if
		    end if
    		strNotify = replace(strNotify,"[PM]",strPMEmailList)
	    	strNotify = replace(strNotify,"[CommodityPM]",strPMEmailList)
		    strNotify = replace(strNotify,"[DevManager]",strDevManagerEmail)
		    strNotify = replace(strNotify,"[Developer]",strDeveloperEmail)
		    strNotify = replace(strNotify,"[TestLeads]",strTestLeadList)

            do while left(strNotify,1) = ";" and len(strNotify) > 2
                strNotify = mid(strNotify,2)
            loop
            
		    if trim(strTesterEmail) <> "" then 
			    if trim(strNotify) = "" then
				    strNotify = strTesterEmail
			    else
				    strNotify = strTesterEmail & ";" & strNotify 
			    end if
		    end if
		
		    if trim(strExecutionEngineerEmail) <> "" then 
			    if trim(strNotify) = "" then
				    strNotify = strExecutionEngineerEmail
			    else
				    strNotify = strExecutionEngineerEmail & ";" & strNotify
			    end if
		    end if
		
		
		    if instr(strNotify,strDeveloperEmail)=0 then 'Add the developer if they are not in the list already
			    if trim(strNotify) = "" then
				    strNotify = strDeveloperEmail
			    else
				    strNotify = strDeveloperEmail & ";" & strNotify
    			end if
	    	end if

		    if trim(rs("AlsoNotify") & "") <> "" then
			    if strNotify="" then
				    strNotify = trim(rs("AlsoNotify") & "")
			    elseif right(strNotify,1)=";" then
				    strNotify = strNotify & trim(rs("AlsoNotify") & "")
			    else
				    strNotify = strNotify & ";" & trim(rs("AlsoNotify") & "")
			    end if
		    end if

            
            'Get Deliverable Details
            strVersion = rs("Version")
            if trim(rs("Revision") & "") <> "" then
                strversion = strVersion & "," & rs("Revision")
            end if
            if trim(rs("Pass") & "") <> "" then
                strversion = strVersion & "," & rs("Pass")
            end if
            response.write "<TR>"
            response.Write "<TD valign=top rowspan=4><input checked id=""chkID"" name=""chkID"" type=""checkbox"" value=""" & rs("ID") & """></TD>"
            response.Write "<TD>" & rs("ID") & "</TD>"
            response.Write "<TD>" & rs("Vendor") & " " & rs("DeliverableName") & "</TD>"
            response.Write "<TD>" & strversion & "</TD>"
            response.Write "<TD>" & rs("PartNumber") & "&nbsp;</TD>"
            response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
            if trim(request("txtFunction")) = "2" then
                response.Write "<TD>Cancelled</TD>"
            else
                response.Write "<TD>" & strToMilestone & "</TD>"
            end if
            response.write "</TR>"
            response.write "<TR>"
            response.Write "<TD><b>Email:</b></TD>"
            response.Write "<TD colspan=5><input  id=""txtEmail"" name=""txtEmail" & rs("ID") & """ style=""width:100%"" type=""text"" tag=""" & strNotify & """value=""" & strNotify & """></TD>"
            response.write "</TR>"
            response.write "<TR>"
            response.Write "<TD><b>Location:</b></TD>"
            response.Write "<TD colspan=5><font size=1 color=green>Enter path to files or tell how to get hardware update.</font><BR><input id=""txtLocation"" name=""txtLocation" & rs("ID") & """ type=""text"" style=""width:100%"" value=""" & server.htmlencode(rs("ImagePath")&"") & """></TD>"
            response.write "</TR>"
            response.write "<TR>"
            response.Write "<TD valign=top><b>Comments:</b></TD>"
            response.Write "<TD colspan=5><font size=1 color=green>Appended to previous comments.</font><BR><textarea id=""txtComments"" name=""txtComments" & rs("ID") & """ rows=3 type=""text"" style=""width:100%""></textarea>"
            response.Write "<input id=""txtExecutionEngineer"" name=""txtExecutionEngineer" & rs("ID") & """ type=""hidden"" value=""" & trim(strExecutionEngineerID) & """>"
            response.Write "<input id=""txtExecutionEngineerName"" name=""txtExecutionEngineerName" & rs("ID") & """ type=""hidden"" value=""" & trim(strExecutionEngineerName) & """>"
            response.Write "<input id=""SelectedMilestoneName"" name=""SelectedMilestoneName" & rs("ID") & """ type=""hidden"" value=""" & strFromMilestone & """>"
            response.Write "<input id=""SelectedMilestoneID"" name=""SelectedMilestoneID" & rs("ID") & """ type=""hidden"" value=""" & strFromMilestoneID & """>"
            response.Write "<input id=""NextMilestoneID"" name=""NextMilestoneID" & rs("ID") & """ type=""hidden"" value=""" & strToMilestoneID & """>"
            response.write "</TD>"
            response.write "</TR>"
        end if
        rs.MoveNext
    loop
    if not (rs.EOF and rs.BOF) then
        response.Write "</TABLE>"    
    end if
    rs.Close
    
    set rs = nothing
    cn.close
    set cn = nothing


	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i

	'	strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

%>
    <input id="txtFunction" name="txtFunction" type="hidden" value="<%=request("txtFunction")%>">
</form>
</body>
</html>


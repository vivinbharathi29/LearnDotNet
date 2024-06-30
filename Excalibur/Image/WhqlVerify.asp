<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Image MDA Verification</title>
<style type="text/css">
td
{
    font-size: x-small;
    font-family: verdana;
	background-color: ivory;
}
th
{
    font-size: x-small;
    font-family: verdana;
	background-color: cornsilk;
}
.SummaryTH
{
    font-size: x-small;
    font-family: verdana;
	background-color: lightsteelblue;
}
.SummaryTD
{
    font-size: x-small;
    font-family: verdana;
	background-color: gainsboro;
}

.Failed
{
    background-color: red;
}

.Passed
{
    background-color: springgreen;
}
</style>
<script type="text/javascript" language="javascript" src="../_ScriptLibrary/sort.js"></script>
<script id="clientEventHandlersJS" type="text/javascript">
<!--
function window_onload() {
	lblProcessing.style.display = "none";
	//ReportTitle.style.display = "";
}

function DisplayTargetIssues(){
	TargetIssuesRow.style.display = "";	
}

function CompareLines(strTable){
	var i;
	document.all("frmCompare" + strTable).submit();
}
//-->
</script>
</head>
<body onload="return window_onload()">
<span id="lblProcessing" style="font:bold x-small verdana ">Processing. This may take several minutes.  Please wait...</span>
<%
Server.ScriptTimeout = 1200

dim StartDate
StartDate = now()

dim cn
dim cn2
dim cm
dim p
dim rs
dim rs2
	
	
TableCount =0
if request("ProdID") = "" then
	Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
	Response.End
End If

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")
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
	p.Value = request("ProdID")
	cm.Parameters.Append p
	rs.CursorType = adOpenStatic
	Set rs = cm.Execute 
	Set cm=nothing

	strPreinstallTeam = 0

	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
	else
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../NoAccess.asp?Level=0"
			end if
		end if
		
		Response.Write "<h2 style=""font-family:verdana,arial;"">" & rs("Name") & "" & " MDA Verification</h2>"
		
		rs.close
		
		
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "usp_GetProductCertificationCount"
        Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
	    p.Value = request("ProdID")
	    cm.Parameters.Append p
	    rs.CursorType = adOpenStatic
	    Set rs = cm.Execute 
	    Set cm=nothing        
        
        if rs.BOF and rs.EOF Then
            totalTargeted = 0
            totalWhql = 0
            totalOEMReady = 0
            totalNA = 0
            totalExceptions = 0
        else
            totalTargeted = rs("TargetedCount")
            totalWhql = rs("WhqlCount")
            totalOEMReady = rs("OemReadyCount")
            totalNA = rs("NotRequiredCount")
            totalExceptions = rs("ExceptionCount")
            
            rs.close
        end if
        
        'Get Targeted Deliverable Count
        'Get Deliverables that require WHQL
        'Get Deliverables that require OEM Ready
        'Get Deliverables that 
        'Get Deliverables that are not ready.
        'Get Deliverables with exceptions/waivers	
        
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "usp_SelectTargetedDeliverablesNotCertified"
        Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
	    p.Value = request("ProdID")
	    cm.Parameters.Append p
	    rs.CursorType = adOpenStatic
	    Set rs = cm.Execute 
	    Set cm=nothing   
	    
	    Dim strDeliverableID, strDeliverableName, strVersion, strRevision, strPass
	    Dim strOemReadyStatus, blnOemReadyRequired, strOemReadyComments, strOemReadyStatusTxt
	    Dim strCertificationStatus, blnCertificationRequired, strCertificationComments, strCertificationStatusTxt
	    Dim strClass
	    
	    totalNotCertified = 0
	    
	    strMissingTargetRows = ""
	    Do Until rs.EOF
	        totalNotCertified = totalNotCertified + 1
	        
	        strDeliverableID = rs("id") & ""
            strDeliverableName = rs("deliverablename") & ""
            strVersion = rs("version") & ""
            strRevision = rs("revision") & ""
            strPass = rs("pass")
            strOemReadyStatus = rs("oemreadystatus") & ""
            blnOemReadyRequired = rs("oemreadyrequired") & ""
            strOemReadyComments = rs("oemreadycomments") & ""
            strCertificationStatus = rs("certificationstatus") & ""
            blnCertificationRequired = CBOOL(rs("CertRequired") & "")
            strCertificationComments = rs("certificationcomments") & ""
            
            If Not blnOemReadyRequired Then
                strOemReadyStatusTxt = "Not Required"
                strOemReadyComments = rs("oemreadyexception") & ""
            Else
                Select Case strOemReadyStatus
                    Case "0"
                        strOemReadyStatusTxt = "Required"
                    Case "1"
                        strOemReadyStatusTxt = "Submitted"
                    Case "2"
                        strOemReadyStatusTxt = "Approved"
                    Case "3"
                        strOemReadyStatusTxt = "Failed"
                    Case "4"
                        strOemReadyStatusTxt = "Exempt"
                    Case Else
                        strOemReadyStatusTxt = "Required"
                End Select        
            End If

            If Not blnCertificationRequired Then	    
                strCertificationStatusTxt = "Not Required"
            Else
                Select Case strCertificationStatus
                    Case "0"
                        strCertificationStatusTxt = "Required"
                    Case "1"
                        strCertificationStatusTxt = "Submitted"
                    Case "2"
                        strCertificationStatusTxt = "Approved"
                    Case "3"
                        strCertificationStatusTxt = "Failed"
                    Case "4"
                        strCertificationStatusTxt = "Waivered"
                    Case Else
                        strCertificationStatusTxt = "Required"
                End Select      
              End If
              
	        strMissingTargetRows = strMissingTargetRows & "<tr>"
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strDeliverableName & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strVersion & "," & strRevision & "," & strPass & "</td>"
	        strClass = ""
            If (blnOemReadyRequired And (strOemReadyStatus = "2" Or strOemReadyStatus = "4")) Then
                strClass = "Passed"
            ElseIf ((Not blnOemReadyRequired) And InStr(LCase(strOemReadyComments), "nitial default value") = 0) Then
                strClass = "Passed"
            Else
                strClass = "Failed"
            End If
	        
	        strMissingTargetRows = strMissingTargetRows & "<td class=""" & strClass & """>" & strOemReadyStatusTxt
            If strOemReadyComments <> "" Then
                strMissingTargetRows = strMissingTargetRows & "<br />" & strOemReadyComments
            End If
            strMissingTargetRows = strMissingTargetRows & "</td>"
	        
            If (blnCertificationRequired And (strCertificationStatus = "2" Or strCertificationStatus = "4")) Or (Not blnCertificationRequired) Then
                strClass = "Passed"
            Else
                strClass = "Failed"
            End If

            strMissingTargetRows = strMissingTargetRows & "<td class=""" & strClass & """>" & strCertificationStatusTxt
            If strOemReadyComments <> "" Then
                strMissingTargetRows = strMissingTargetRows & "<br />" & strCertificationComments
            End If
	        strMissingTargetRows = strMissingTargetRows & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "</tr>"
	        rs.MoveNext
	    Loop
        rs.close
        	        

        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "usp_SelectTargetedDeliverablesWithWaiver"
        Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
	    p.Value = request("ProdID")
	    cm.Parameters.Append p
	    rs.CursorType = adOpenStatic
	    Set rs = cm.Execute 
	    Set cm=nothing   
	    
	    totalExceptions = 0
	    
	    strExemptRows = ""
	    Do Until rs.EOF
	        totalExceptions = totalExceptions + 1
	        
	        strDeliverableID = rs("id") & ""
            strDeliverableName = rs("deliverablename") & ""
            strVersion = rs("version") & ""
            strRevision = rs("revision") & ""
            strPass = rs("pass")
            strOemReadyStatus = rs("oemreadystatus") & ""
            blnOemReadyRequired = rs("oemreadyrequired") & ""
            strOemReadyComments = rs("oemreadycomments") & ""
            strCertificationStatus = rs("certificationstatus") & ""
            blnCertificationRequired = rs("CertRequired") & ""
            strCertificationComments = rs("certificationcomments") & ""
            
            If blnOemReadyRequired = "0" Then
                strOemReadyStatusTxt = "Not Required"
            Else
                Select Case strOemReadyStatus
                    Case "0"
                        strOemReadyStatusTxt = "Required"
                    Case "1"
                        strOemReadyStatusTxt = "Submitted"
                    Case "2"
                        strOemReadyStatusTxt = "Approved"
                    Case "3"
                        strOemReadyStatusTxt = "Failed"
                    Case "4"
                        strOemReadyStatusTxt = "Exempt"
                    Case Else
                        strOemReadyStatusTxt = "Required"
                End Select        
            End If

            If blnCertificationRequired = "0" Then	    
                strCertificationStatusTxt = "Not Required"
            Else
                Select Case strCertificationStatus
                    Case "0"
                        strCertificationStatusTxt = "Required"
                    Case "1"
                        strCertificationStatusTxt = "Submitted"
                    Case "2"
                        strCertificationStatusTxt = "Approved"
                    Case "3"
                        strCertificationStatusTxt = "Failed"
                    Case "4"
                        strCertificationStatusTxt = "Waivered"
                    Case Else
                        strCertificationStatusTxt = "Required"
                End Select      
              End If
              
	        strExemptRows = strExemptRows & "<tr>"
	        strExemptRows = strExemptRows & "<td>" & strDeliverableName & "</td>"
	        strExemptRows = strExemptRows & "<td>" & strVersion & "," & strRevision & "," & strPass & "</td>"
            If (blnOemReadyRequired = "1" And (strOemReadyStatus = "2" Or strOemReadyStatus = "4")) Or (blnOemReadyRequired = "0") Then
	            strExemptRows = strExemptRows & "<td class=""Passed"">" & strOemReadyStatusTxt
                If strOemReadyComments <> "" Then
                    strExemptRows = strExemptRows & "<br />" & strOemReadyComments
                End If
	            strExemptRows = strExemptRows & "</td>"
	        else
	            strExemptRows = strExemptRows & "<td class=""Failed"">" & strOemReadyStatusTxt
                If strOemReadyComments <> "" Then
                    strExemptRows = strExemptRows & "<br />" & strOemReadyComments
                End If
	            strExemptRows = strExemptRows & "</td>"
	        End If
	        
            If (blnCertificationRequired = "1" And (strCertificationStatus = "2" Or strCertificationStatus = "4")) Or (blnCertificationRequired = "0") Then
	            strExemptRows = strExemptRows & "<td class=""Passed"">" & strCertificationStatusTxt
                If strOemReadyComments <> "" Then
                    strExemptRows = strExemptRows & "<br />" & strCertificationComments
                End If
	            strExemptRows = strExemptRows & "</td>"
	        else
	            strExemptRows = strExemptRows & "<td class=""Failed"">" & strCertificationStatusTxt
                If strOemReadyComments <> "" Then
                    strExemptRows = strExemptRows & "<br />" & strCertificationComments
                End If
	            strExemptRows = strExemptRows & "</td>"
	        End If
	        strExemptRows = strExemptRows & "</tr>"
	        rs.MoveNext
	    Loop
        rs.close

	end if
%>
<span style="font:x-small verdana;">
<%
Response.Write "<font size=2 face=verdana><u><b>Summary</b></u></font><BR><BR>"
if totalErrorCount > TotalCompared then
	totalErrorCount = TotalCompared
end if
if totalErrorCount = "" then
	totalErrorCount = "0"
end if
if TotalCompared = "" then
	TotalCompared = "0"
end if
%>
</span>

<table border="0" style="border-color:indigo;" cellspacing="1" cellpadding="2">
<tr><td class="SummaryTH">Deliverables&nbsp;Targeted:</td><td class="SummaryTD">&nbsp;<%=totalTargeted%>&nbsp;</td></tr>
<tr><td class="SummaryTH">Deliverables Requiring WHQL:</td><td class="SummaryTD">&nbsp;<%=totalWhql%>&nbsp;</td></tr>
<tr><td class="SummaryTH">Deliverables Requiring OEM Ready:</td><td class="SummaryTD">&nbsp;<%=totalOEMReady%>&nbsp;</td></tr>
<tr><td class="SummaryTH">Deliverables with Exceptions/Waivers:</td><td class="SummaryTD">&nbsp;<%=totalExceptions%>&nbsp;</td></tr>
<tr><td class="SummaryTH">Discrepancies Found:</td><td class="SummaryTD">&nbsp;<%=totalNotCertified%>&nbsp;</td></tr>
</table><br />


<%
	if strHideTables <> "" then
		strHideTables = mid(strHideTables,2)
		strHideTables=""
	end if
%>
<span style="font: x-small verdana;">
<%
Response.Write "<font size=2 face=verdana><u><b><BR><BR>Potential " & strProduct & " Issues</b></u></font><BR><BR>"
%>
</span>
<table id="TargetIssuesRow" style="display:; border-color:tan; background-color:ivory;" width="100%" border="1">
<thead>
<tr>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 0,4,1);">Deliverable</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 1,4,1);">Version</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 2,4,1);">OEM Ready</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 3,4,1);">WHQL</a></th>
</tr></thead>
<%=strMissingTargetRows%>
</table>
<span style="font: x-small verdana;">
<%
Response.Write "<font size=2 face=verdana><u><b><BR><BR>Exempt/Wavered Deliverables</b></u></font><BR><BR>"
%>
</span>
<table id="Table1" style="display:; border-color:tan; background-color:ivory;" width="100%" border="1">
<tr>
<th style="text-align:left; font-weight:bold;">Deliverable</th>
<th style="text-align:left; font-weight:bold;">Version</th>
<th style="text-align:left; font-weight:bold;">OEM Ready</th>
<th style="text-align:left; font-weight:bold;">WHQL</th>
</tr>
<%=strExemptRows%>
</table>
<input type="hidden" id="txtHideTables" name="txtHideTables" value="<%=strHideTables%>" />
</body>
</html>

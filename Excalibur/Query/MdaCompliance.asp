<%@ Language=VBScript %>
<% Option Explicit %>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"	

	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
	end if
%>
<!-- #include file ="../includes/ReportFunctions.asp" -->
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
dim StartDate
StartDate = now()

dim cn
dim cn2
dim cm
dim p
dim rs
dim rs2
	
Dim strProductIds
Dim strProductIdsPulsar
Dim TableCount
Dim CurrentUser
Dim CurrentUserPartner
Dim totalTargeted, totalWhql, totalOEMReady, totalNA, totalExceptions, totalNotCertified
Dim strMissingTargetRows, totalErrorCount, TotalCompared, strHideTables, strProduct
Dim strExemptRows

strProductIdsPulsar = ""

strProductIds = request("lstProducts")
 ' wgomero: PBI 18749 add Pulsar products to the list if any
strProductIdsPulsar = request("lstProductsPulsar")
 dim productReleases
 dim productPulsarIds
 dim productReleaseIds
 dim productReleaseId
    
    productPulsarIds=""
    productReleaseIds=""
    productReleaseId=""

    productReleases = split(request("lstProductsPulsar"),",")
    dim productRelease
    for each productRelease in productReleases
      if instr(productRelease,":")>0 then
		productReleaseId = split(productRelease,":")
        productPulsarIds = productPulsarIds + "," + productReleaseId(0)
        productReleaseIds = productReleaseIds + "," + productReleaseId(1)
      end if
    next


if strProductIds <> "" and productPulsarIds <> "" then
        strProductIds = strProductIds + "," + productPulsarIds
    end if

if strProductIds = "" and productPulsarIds <> "" then
        strProductIds =  productPulsarIds
    end if

if strProductIds <> "" and productPulsarIds = "" then
        strProductIds =  strProductIds
    end if
' end

If request.Form("lstProductGroups") <> "" Then
strProductIds = strProductIds & ProcessProductGroupListGetIds(request("lstProductGroups"))
strProductIds = Replace(strProductIds, ",,", ",")
Else
strProductIds = strProductIds
End If

TableCount =0
if strProductIds = "" then
	Response.Write "<BR><font size=2 face=verdana>No Products were selected.</font><BR><BR>"
	Response.End
End If

    If Left(strProductIds, 1) = "," Then
        strProductIds = Mid(strProductIds, 2, Len(strProductIds) - 1)
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
	
'	set cm = server.CreateObject("ADODB.Command")
'	Set cm.ActiveConnection = cn
'	cm.CommandType = 4
'	cm.CommandText = "spGetProductVersionName"
'	Set p = cm.CreateParameter("@ID", 3, &H0001)
'	p.Value = request("ProdID")
'	cm.Parameters.Append p
'	rs.CursorType = adOpenStatic
'	Set rs = cm.Execute 
'	Set cm=nothing

'	strPreinstallTeam = 0

	If True Then
'		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
'	else
		'Verify Access is OK
'		if trim(CurrentUserPartner) <> "1" then
'			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
'				set rs = nothing
'				set cn=nothing
				
'				Response.Redirect "../NoAccess.asp?Level=0"
'			end if
'		end if
		
'		Response.Write "<h2 style=""font-family:verdana,arial;"">" & rs("Name") & "" & " MDA Verification</h2>"
		Response.Write "<h2 style=""font-family:verdana,arial;"">MDA Verification</h2>"
		
'		rs.close
		
		
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "usp_GetProductsCertificationCount"
        Set p = cm.CreateParameter("@p_ProductVersionID", 200, 1, 40000)
	    p.Value = strProductIds
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
        cm.CommandType = 1
        cm.CommandText = "select distinct dv.ID, v.name as VendorName, DeliverableName, Version, Revision, Pass, VendorVersion, OEMReadyStatus, OEMReadyComments, OEMReadyException, OEMReadyRequired, CertificationStatus, CertificationComments, CertRequired " & _
                        "from product_deliverable pdv with (NOLOCK) inner join deliverableversion dv with (NOLOCK) on pdv.deliverableversionid = dv.id inner join deliverableroot dr with (NOLOCK) on dv.deliverablerootid = dr.id inner join vendor v with (NOLOCK) on dr.vendorid = v.id " & _
                        "where dr.typeid = 2 and targeted = 1 and productversionid in (" & strProductIds & ") " & _
                        "order by DeliverableName, Version, Revision, Pass"
	    Set rs = cm.Execute 
	    Set cm=nothing   
	    
        'strProductIds

	    
	    Dim strDeliverableID, strDeliverableName, strVersion, strRevision, strPass, strVendorVersion, strVendorName
	    Dim strOemReadyStatus, blnOemReadyRequired, strOemReadyComments, strOemReadyStatusTxt
	    Dim strCertificationStatus, blnCertificationRequired, strCertificationComments, strCertificationStatusTxt
	    Dim strClass
	    
	    totalNotCertified = 0
	    
	    strMissingTargetRows = ""
	    Do Until rs.EOF
	        totalNotCertified = totalNotCertified + 1
	        
	        strDeliverableID = rs("id") & ""
	        strVendorName = rs("VendorName") & ""
	        strVendorVersion = rs("VendorVersion") & ""
            strDeliverableName = rs("deliverablename") & ""
            strVersion = rs("version") & ""
            strRevision = rs("revision") & ""
            strPass = rs("pass")
            strOemReadyStatus = rs("oemreadystatus") & ""
            If Not IsNull(rs("oemreadyrequired").Value) Then
                 blnOemReadyRequired = CBOOL(rs("oemreadyrequired") & "")
            Else
                 blnOemReadyRequired = CBOOL(0 & "")
            End If
            strOemReadyComments = rs("oemreadycomments") & ""
            strCertificationStatus = rs("certificationstatus") & ""
             If Not IsNull(rs("CertRequired").Value) Then
                 blnCertificationRequired = CBool(rs("CertRequired") & "")
            Else
                 blnCertificationRequired = CBool(0 & "")
            End If
            strCertificationComments = rs("certificationcomments") & ""
            
            If strVendorVersion = "" Then
                strVendorVersion = "&nbsp;"
            End If
            
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
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strVendorName & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strDeliverableName & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strVersion & "," & strRevision & "," & strPass & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "<td>" & strVendorVersion & "</td>"
	        strClass = ""
	        
	        If strOemReadyComments = "" Then strOemReadyComments = "&nbsp;"
	        
            If (blnOemReadyRequired And (strOemReadyStatus = "2" Or strOemReadyStatus = "4")) Then
                strClass = "Passed"
            ElseIf ((Not blnOemReadyRequired) And InStr(LCase(strOemReadyComments), "nitial default value") = 0) Then
                strClass = "Passed"
            Else
                strClass = "Failed"
            End If
	        
	        strMissingTargetRows = strMissingTargetRows & "<td class=""" & strClass & """>" & strOemReadyStatusTxt
            If strOemReadyComments <> "" Then
                strMissingTargetRows = strMissingTargetRows & "</td><td>" & strOemReadyComments
            End If
            strMissingTargetRows = strMissingTargetRows & "</td>"
	        
            If (blnCertificationRequired And (strCertificationStatus = "2" Or strCertificationStatus = "4")) Or (Not blnCertificationRequired) Then
                strClass = "Passed"
            Else
                strClass = "Failed"
            End If

            If strCertificationCOmments = "" Then strCertificationComments = "&nbsp;"

            strMissingTargetRows = strMissingTargetRows & "<td class=""" & strClass & """>" & strCertificationStatusTxt
            If strOemReadyComments <> "" Then
                strMissingTargetRows = strMissingTargetRows & "</td><td>" & strCertificationComments
            End If
	        strMissingTargetRows = strMissingTargetRows & "</td>"
	        strMissingTargetRows = strMissingTargetRows & "</tr>"
	        rs.MoveNext
	    Loop
        rs.close
        	        
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 1
        cm.CommandText = "Select f.Name, v.Version From ProductFamily f with (NOLOCK) INNER JOIN ProductVersion v with (NOLOCK) ON f.ID = v.ProductFamilyID WHERE v.ID IN (" & strProductIds & ") Order By f.Name, v.Version"
	    Set rs = cm.Execute 
	    Set cm=nothing   
        
        Dim strProductNames
        strProductNames = ""
        Do Until rs.EOF
            strProductNames = strProductNames & ", " & rs("Name") & "&nbsp;" & rs("Version")
            rs.MoveNext
        Loop
        rs.close
        set rs = nothing
        
        If Left(strProductNames, 1) = "," Then
            strProductNames = Mid(strProductNames, 3, Len(strProductNames) -2)
        End If

'        Set cm = Server.CreateObject("ADODB.Command")
'        Set cm.ActiveConnection = cn
'        cm.CommandType = 4
'        cm.CommandText = "usp_SelectTargetedDeliverablesWithWaiver"
'        Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
'	    p.Value = request("ProdID")
'	    cm.Parameters.Append p
'	    rs.CursorType = adOpenStatic
'	    Set rs = cm.Execute 
'	    Set cm=nothing   
	    
'	    totalExceptions = 0
	    
'	    strExemptRows = ""
'	    Do Until rs.EOF
'	        totalExceptions = totalExceptions + 1
'	        
'	        strDeliverableID = rs("id") & ""
'            strDeliverableName = rs("deliverablename") & ""
'            strVersion = rs("version") & ""
'            strRevision = rs("revision") & ""
'            strPass = rs("pass")
'            strOemReadyStatus = rs("oemreadystatus") & ""
'            blnOemReadyRequired = rs("oemreadyrequired") & ""
'            strOemReadyComments = rs("oemreadycomments") & ""
'            strCertificationStatus = rs("certificationstatus") & ""
'            blnCertificationRequired = rs("CertRequired") & ""
'            strCertificationComments = rs("certificationcomments") & ""
'            
'            If blnOemReadyRequired = "0" Then
'                strOemReadyStatusTxt = "Not Required"
'            Else
'                Select Case strOemReadyStatus
'                    Case "0"
'                        strOemReadyStatusTxt = "Required"
'                    Case "1"
'                        strOemReadyStatusTxt = "Submitted"
'                    Case "2"
'                        strOemReadyStatusTxt = "Approved"
'                    Case "3"
'                        strOemReadyStatusTxt = "Failed"
'                    Case "4"
'                        strOemReadyStatusTxt = "Exempt"
'                    Case Else
'                        strOemReadyStatusTxt = "Required"
'                End Select        
'            End If
'
'            If blnCertificationRequired = "0" Then	    
'                strCertificationStatusTxt = "Not Required"
'            Else
'                Select Case strCertificationStatus
'                    Case "0"
'                        strCertificationStatusTxt = "Required"
'                    Case "1"
'                        strCertificationStatusTxt = "Submitted"
'                    Case "2"
'                        strCertificationStatusTxt = "Approved"
'                    Case "3"
'                        strCertificationStatusTxt = "Failed"
'                    Case "4"
'                        strCertificationStatusTxt = "Waivered"
'                    Case Else
'                        strCertificationStatusTxt = "Required"
'                End Select      
'              End If
              
'	        strExemptRows = strExemptRows & "<tr>"
'	        strExemptRows = strExemptRows & "<td>" & strDeliverableName & "</td>"
'	        strExemptRows = strExemptRows & "<td>" & strVersion & "," & strRevision & "," & strPass & "</td>"
'            If (blnOemReadyRequired = "1" And (strOemReadyStatus = "2" Or strOemReadyStatus = "4")) Or (blnOemReadyRequired = "0") Then
'	            strExemptRows = strExemptRows & "<td class=""Passed"">" & strOemReadyStatusTxt
'                If strOemReadyComments <> "" Then
'                    strExemptRows = strExemptRows & "<br />" & strOemReadyComments
'                End If
'	            strExemptRows = strExemptRows & "</td>"
'	        else
'	            strExemptRows = strExemptRows & "<td class=""Failed"">" & strOemReadyStatusTxt
'                If strOemReadyComments <> "" Then
'                    strExemptRows = strExemptRows & "<br />" & strOemReadyComments
'                End If
'	            strExemptRows = strExemptRows & "</td>"
'	        End If
	        
'            If (blnCertificationRequired = "1" And (strCertificationStatus = "2" Or strCertificationStatus = "4")) Or (blnCertificationRequired = "0") Then
'	            strExemptRows = strExemptRows & "<td class=""Passed"">" & strCertificationStatusTxt
'                If strOemReadyComments <> "" Then
'                    strExemptRows = strExemptRows & "<br />" & strCertificationComments
'                End If
'	            strExemptRows = strExemptRows & "</td>"
'	        else
'	            strExemptRows = strExemptRows & "<td class=""Failed"">" & strCertificationStatusTxt
'                If strOemReadyComments <> "" Then
'                    strExemptRows = strExemptRows & "<br />" & strCertificationComments
'                End If
'	            strExemptRows = strExemptRows & "</td>"
'	        End If
'	        strExemptRows = strExemptRows & "</tr>"
'	        rs.MoveNext
'	    Loop
'        rs.close
'
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
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 0,4,1);">Vendor</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 1,4,1);">Deliverable</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 2,4,1);">HP Version</a></th>
<th style="text-align:left; font-weight:bold;"><a href="javascript: SortTable( 'TargetIssuesRow', 3,4,1);">Vendor Version</a></th>
<th style="text-align:left; font-weight:bold; width:10%;"><a href="javascript: SortTable( 'TargetIssuesRow', 4,4,1);">OEM Ready</a></th>
<th style="text-align:left; font-weight:bold; width:10%;"><a href="javascript: SortTable( 'TargetIssuesRow', 5,4,1);">OEM Ready Comments</a></th>
<th style="text-align:left; font-weight:bold; width:10%;"><a href="javascript: SortTable( 'TargetIssuesRow', 6,4,1);">WHQL</a></th>
<th style="text-align:left; font-weight:bold; width:10%;"><a href="javascript: SortTable( 'TargetIssuesRow', 7,4,1);">WHQL Comments</a></th>
</tr></thead>
<%=strMissingTargetRows%>
</table>
<!--<span style="font: x-small verdana;">
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
</table>-->
<input type="hidden" id="txtHideTables" name="txtHideTables" value="<%=strHideTables%>" />
<p style="font-family:Verdana; font-size:xx-small"><strong>Products Selected:&nbsp;</strong><%= strProductNames %></p>
<font Size="2" Color="red"><p><strong>Confidential</strong></p></font>
</body>
</html>


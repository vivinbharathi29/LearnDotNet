<%@  language="VBScript" %>
<% Option Explicit %>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<%
    dim dtTime0, dtTime1, dtTime2, dtTime3
    dtTime0 = Now() 

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
    <title>MDA Compliance 2015</title>
    <style type="text/css">
        td {
            font-size: x-small;
            font-family: verdana;
            background-color: ivory;
        }

        th {
            font-size: x-small;
            font-family: verdana;
            background-color: cornsilk;
        }

        .SummaryTH {
            font-size: x-small;
            font-family: verdana;
            background-color: lightsteelblue;
        }

        .SummaryTD {
            font-size: x-small;
            font-family: verdana;
            background-color: gainsboro;
        }

        .Failed {
            background-color: red;
        }

        .Passed {
            background-color: springgreen;
        }
    </style>
    <script type="text/javascript" language="javascript" src="../_ScriptLibrary/sort.js"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
<!--
    function window_onload() {
        lblProcessing.style.display = "none";
    }

    function DisplayTargetIssues() {
        TargetIssuesRow.style.display = "";
    }

    function CompareLines(strTable) {
        var i;
        document.all("frmCompare" + strTable).submit();
    }

    function SortByCol(sortOnColumn) {
        SortTable('TargetIssuesRow', sortOnColumn, 4, 1);
    }

    //-->
    </script>
</head>
<body onload="return window_onload()">
    <span id="lblProcessing" style="font: bold x-small verdana">Processing. This may take several minutes.  Please wait...</span>
    <%
Server.ScriptTimeout = 290

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
Dim strMissingTargetRows
Dim totalErrorCount, TotalCompared, strHideTables, strProduct
Dim strExemptRows
Dim strProductGroups

strProductGroups = ""
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
    strProductGroups = request("lstProductGroups")

    strProductIds = strProductIds & ProcessProductGroupListGetAllIds(strProductGroups)
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
    cn.CommandTimeout = 290
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
    cm.CommandTimeout = 290
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
	

	If True Then

		Response.Write "<h2 style=""font-family:verdana,arial;"">MDA Compliance 2015</h2>"

        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "usp_GetProductsMDACount"
        Set p = cm.CreateParameter("@p_ProductVersionID", 200, 1, 4000)
	    p.Value = strProductIds
	    cm.Parameters.Append p
	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
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
        
        dtTime1=now()
        
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandTimeout = 290
        cm.CommandType = 4
        cm.CommandText = "usp_SelectDeliverableMDA"
        Set p = cm.CreateParameter("@strProductVersionIDs",201,1,8000)
	    p.Value = strProductIds
	    cm.Parameters.Append p
	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 
	    Set cm=nothing   
	    
        dtTime2 = now()

	    Dim strDeliverableID, strDeliverableName, strVersion, strRevision, strPass, strVendorVersion, strVendorName
        Dim strCategory, strPartNumber, strIRSPartNumber, strHWID
	    Dim strCertificationStatus, blnCertificationRequired, strCertificationComments, strCertificationStatusTxt
        Dim strSoftpaqNumber, strSoftpaqTitle, strSoftpaqStatus, strSoftPaqXML
	    Dim strClass

        Dim strProductName, strCycleName, strSysBoardId, strReleasesInProduct
	    Dim strMdaId, strMdaDate, boolWhqlRequired, strMdaStatus

	    totalNotCertified = 0
	    
	    strMissingTargetRows = ""
	    Do Until rs.EOF
	        totalNotCertified = totalNotCertified + 1
	        
            strProductName = rs("ProductVersionName") & ""
            strCycleName = rs("CycleName") & ""
            strReleasesInProduct =  rs("ReleasesInProduct") & ""
            strSysBoardId = rs("SystemBoardID") & ""
	        strDeliverableID = rs("DeliverableVersionID") & ""
	        strVendorName = rs("VendorName") & ""
            strDeliverableName = rs("DriverName") & ""
            strCategory = rs("RootCategory") & ""
            strPartNumber = rs("PartNumber") & ""
            strIRSPartNumber = rs("IRSPartNumber") & ""
            strVersion = rs("version") & ""
            strRevision = rs("revision") & ""
            strPass = rs("pass") & ""
	        strVendorVersion = rs("VendorVersion") & ""
            strMdaId = rs("SubmissionId") & ""
            strMdaDate = rs("SubmissionDate") & ""
            boolWhqlRequired= CBOOL(rs("CertRequired") & "")
            strHWID = rs("PNPDevices") & "" 
            strSoftpaqNumber = rs("SoftpaqNumbers") & ""
            strSoftpaqTitle = rs("SoftpaqTitle") & ""
            strSoftPaqXML = rs("SoftpaqSatausXML") & ""
            strSoftpaqStatus = "" 
            strCertificationStatus = rs("CertificationStatus") & ""
            blnCertificationRequired = CBOOL(rs("CertRequired") & "")
            strCertificationComments = rs("CertificationComments") & ""
            
        'Set "Empty" to "Blank Space"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If trim(strSysBoardId) ="" Then strSysBoardId = "&nbsp;"
            If trim(strPartNumber) = "" Then strPartNumber = "&nbsp;"
            If trim(strIRSPartNumber) = "" Then strIRSPartNumber = "&nbsp;"
            If strMdaId = "" Then  strMdaId = "&nbsp;"
            If strMdaDate = "" Then strMdaDate = "&nbsp;"
            If strMdaStatus = "" Then  strMdaStatus = "&nbsp;"
            If strVendorVersion = "" Then strVendorVersion = "&nbsp;"
            If trim(strCategory ) = "" Then strCategory  = "&nbsp;"
            If trim(strHWID ) = "" Then strHWID  = "&nbsp;"
            If trim(strSoftpaqNumber ) = "" Then strSoftpaqNumber  = "&nbsp;"
            If trim(strSoftpaqTitle ) = "" Then strSoftpaqTitle  = "&nbsp;"
            If trim(strSoftpaqStatus ) = "" Then strSoftpaqStatus  = "&nbsp;"
            If strCertificationCOmments = "" Then strCertificationComments = "&nbsp;"
        'Set "Empty" to "Blank Space''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'fix strSysBoardId in Excel
            strSysBoardId = replace(strSysBoardId,",",", ")


            If strCycleName = "" Then
                strCycleName = "&nbsp;"
            ElseIf Right(strCycleName,2) = ", " Then
                strCycleName = Left(strCycleName, Len(strCycleName) -2)
            End If

            If strReleasesInProduct = "" Then
                strReleasesInProduct = "&nbsp;"
            ElseIf Right(strReleasesInProduct,2) = ", " Then
                strReleasesInProduct = Left(strReleasesInProduct, Len(strReleasesInProduct) -2)
            End If

            If boolWhqlRequired Then
                strMdaStatus = "Yes"
            Else
                strMdaStatus = "No"
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
              
        	strClass = ""
            If (blnCertificationRequired And (strCertificationStatus = "2" Or strCertificationStatus = "4")) Or (Not blnCertificationRequired) Then
                strClass = "Passed"
            Else
                strClass = "Failed"
            End If

            If strSoftPaqXML <> "" then 
                strSoftPaqXML = Replace(strSoftPaqXML, "</Status>" + strDeliverableID + ",<spnumber>","</td></tr><tr><td>")
                strSoftPaqXML = Replace(strSoftPaqXML, "spnumber","td")
                strSoftPaqXML = Replace(strSoftPaqXML, "Status","td")
                strSoftPaqXML = Replace(strSoftPaqXML, strDeliverableID + ",","<tr>")
                strSoftPaqXML = "<table id=""ID" + strDeliverableID + """ width=""100%""  border=""1"" style=""white-space: nowrap;"" >" + strSoftPaqXML + "</tr></table>"
                strSoftPaqXML = Replace(strSoftPaqXML, "<td>","<td width=""50%"" style=""text-wrap: none;"" >")
            Else
                strSoftPaqXML = "&nbsp;"
            End If

            strSoftpaqStatus = strSoftPaqXML


	        strMissingTargetRows = strMissingTargetRows + "<tr id=""" + CStr(totalNotCertified) + """>" _
                                    + "<td>" + strProductName + "</td>" _
                                    + "<td>" + strCycleName + "</td>" _
                                    + "<td>" + strReleasesInProduct + "</td>" _
                                    + "<td>" + strSysBoardId + "</td>" _
	                                + "<td>" + strVendorName + "</td>" _
	                                + "<td>" + strDeliverableName + "</td>" _
                                    + "<td>" + strCategory + "</td>" _
                                    + "<td>" + strDeliverableID + "</td>" _
                                    + "<!--<td IRSPartNumber=""" + strIRSPartNumber + """ PulsarPartNumber=""" + strPartNumber +  """>" + strIRSPartNumber + "</td> -->" _
	                                + "<td>" + strVersion + "," + strRevision + "," + strPass + "</td>" _
	                                + "<td>" + strVendorVersion + "</td>" _
                                    + "<td>" + strMdaId + "</td>" _
                                    + "<td>" + strMdaDate + "</td>" _
                                    + "<td>" + strMdaStatus + "</td>" _
                                    + "<td>" + strHWID + "</td>" _
                                    + "<td>" + strSoftpaqTitle + "</td>" _
                                    + "<td>" + strSoftpaqStatus + "</td>" _
                                    + "<!-- <td>" + strSoftpaqNumber + "</td> -->" _
                                    + "<!-- <td class=""" + strClass + """>" + strCertificationStatusTxt _
                                    + "</td> --> <!--<td>" + strCertificationComments _
	                                + "</td> -->" _
	                                + "</tr>" _
                                    + vbCrLf 
	        rs.MoveNext
	    Loop
        rs.close
        	       
        dtTime3 =now()
         
        Set cm = Server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 1
        cm.CommandText = "select DOTSName from ProductVersion with (NOLOCK) where ID in ( select * from dbo.Split('"  & strProductIds & "')) order by DOTSName "
	    Set rs = cm.Execute 
	    Set cm=nothing   
        
        Dim strProductNames
        strProductNames = ""
        Do Until rs.EOF
            strProductNames = strProductNames & ", " & rs(0) 
            rs.MoveNext
        Loop
        rs.close
        set rs = nothing
        
        If Left(strProductNames, 1) = "," Then
            strProductNames = Mid(strProductNames, 3, Len(strProductNames) -2)
        End If
        
	end if
    %>
    <span style="font: x-small verdana;">
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

    <table border="0" style="border-color: indigo;" cellspacing="1" cellpadding="2">
        <tr>
            <td class="SummaryTH">Drivers&nbsp;Targeted:</td>
            <td class="SummaryTD">&nbsp;<%=totalNotCertified%>&nbsp;</td>
        </tr>
        <tr>
            <td class="SummaryTH">Drivers Requiring WHQL:</td>
            <td class="SummaryTD">&nbsp;<%=totalWhql%>&nbsp;</td>
        </tr>

    </table>
    <br />


    <%
	if strHideTables <> "" then
		strHideTables = mid(strHideTables,2)
		strHideTables=""
	end if
    %>
    <span style="font: x-small verdana;">
        <%
Response.Write "<font size=2 face=verdana><u><b><BR><BR>Driver Tracking</b></u></font><BR><BR>"
        %>
    </span>
    <table id="TargetIssuesRow" style="display: ; border-color: tan; background-color: ivory;" width="100%" border="1">
        <thead>
            <tr style="text-align: left; font-weight: bold;">
                <th><a href="#" onclick="SortByCol(0);">Product</a></th>
                <th><a href="#" onclick="SortByCol(1);">Cycle</a></th>
                <th><a href="#" onclick="SortByCol(2);">Product_Releases </a></th>
                <th><a href="#" onclick="SortByCol(3);">Sys Board ID</a></th>
                <th><a href="#" onclick="SortByCol(4);">Vendor</a></th>
                <th><a href="#" onclick="SortByCol(5);">Driver</a></th>
                <th><a href="#" onclick="SortByCol(6);">Category</a></th>
                <th><a href="#" onclick="SortByCol(7);">Component Version ID</a></th>
                <!-- <th><a href="#" onclick="SortByCol(7);">Part Number</a></th> -->
                <th><a href="#" onclick="SortByCol(8);">HP Version</a></th>
                <th><a href="#" onclick="SortByCol(9);">Vendor Version</a></th>
                <th><a href="#" onclick="SortByCol(10);">WHQL ID</a></th>
                <th><a href="#" onclick="SortByCol(11);">Date Submitted</a></th>
                <th><a href="#" onclick="SortByCol(12);">WHQL Certification Required</a></th>
                <th><a href="#" onclick="SortByCol(13);">HWID </a></th>
                <th><a href="#" onclick="SortByCol(14);">SoftPaq Title</a></th>
                <th>
                    <table border="1" style="width: 100%; white-space: pre; border-spacing: 0;" cellspacing="0">
                        <tr style="text-align: left; font-weight: bold;">
                            <th style="width: 50%; white-space: pre; padding: 4px;"><a href="#">Softpaq Number</a></th>
                            <th style="width: 50%; white-space: pre; padding: 4px;"><a href="#">Softpaq Status</a></th>
                        </tr>
                    </table>
                </th>
                <!-- <th><a href="#" onclick="SortByCol(15);">Softpaq Number</a></th> -->
                <!-- <th><a href="#" onclick="SortByCol(14);">WHQL</a></th> -->
                <!-- <th><a href="#" onclick="SortByCol(15);">WHQL Comments</a></th> -->
            </tr>
        </thead>

        <%=strMissingTargetRows%>
    </table>

    <input type="hidden" id="txtHideTables" name="txtHideTables" value="<%=strHideTables%>" />
    <p style="font-family: Verdana; font-size: xx-small"><strong>Products Selected:&nbsp;</strong><%= strProductNames %></p>
    <p><font size="2" color="red"><strong>Confidential</strong></font></p>
    <!-- ProductGroups: <%=strProductGroups %> -->
    <!--ProductIds=    
        <%=strProductIds %> 
        -->
    <!--   Time    -->
    <!-- Server Start: <%=dtTime0 %> -->
    <!-- 1: <%=DateDiff("s",dtTime0,dtTime1) %>' -->
    <!-- 2: <%=DateDiff("s",dtTime1,dtTime2) %>' -->
    <!-- 3: <%=DateDiff("s",dtTime2,dtTime3) %>' -->
    <!-- Server End: <%=Now() %>  :::  <%= DateDiff("s",dtTime0,Now()) %>'  -->
</body>
</html>


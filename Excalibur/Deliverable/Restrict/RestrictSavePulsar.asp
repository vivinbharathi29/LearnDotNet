<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                //close window
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    parent.window.parent.closeExternalPopup();
                }
                else
                    window.parent.Close(true);
            }
        }
    }

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>');">
<%
	dim cn
	dim cm
	dim rs
	dim strRestriction
	dim strBody
	dim strTO
	dim strFrom
	dim strSubject
	dim strCC
	dim strSuccess
	dim LogIDList
	dim LogIDArray
	dim LogIDArrayParts
	dim LogIDPair
	dim LogText
	
	strCC = request("txtCC")
	
	if request("optRestriction")= "0" then
		strRestriction = "1"
		strSubject = "Supply Chain Restriction Added"
		LogText = "Restriction Added"
	else
		strRestriction = "0"
		strSubject = "Supply Chain Restriction Removed"
		LogText = "Restriction Removed"
	end if


	dim strSQL
	dim strVersionIDList
    dim strPulsarProductList
    dim strLegacyProductList
    dim strReleaseList

	strVersionIDList = request("chkVersions")
    strPulsarProductList = request("txtPulsarProductList")
	strLegacyProductList = request("txtLegacyProductList")
	
    if len(strPulsarProductList) > 0 then
	    strPulsarSQlMain = "UPDATE pdr " & _
			         "SET pdr.SupplyChainRestriction = " & strRestriction & " " & _
                     "FROM Product_Deliverable pd WITH (NOLOCK) " &_
                     "INNER JOIN Product_Deliverable_Release pdr WITH (NOLOCK) on pdr.ProductDeliverableID = pd.ID " &_
                     "inner join productversion_release pvr with (NOLOCK) on pvr.ProductVersionID = pd.ProductVersionID and pvr.ReleaseID = pdr.ReleaseID " &_
			         "WHERE pvr.ID in (" & strPulsarProductList & ") " & _
			         "AND DeliverableVersionID in (" & strVersionIDList & ") "
    end if

    if len(strLegacyProductList) > 0 then
        strLegacySQlMain = "UPDATE Product_Deliverable " & _
			                 "SET SupplyChainRestriction = " & strRestriction & " " & _
			                 "WHERE ProductVersionID in (" & strLegacyProductList & ") " & _
			                 "AND DeliverableVersionID in (" & strVersionIDList & ") "
    end if
	
	set cn = server.CreateObject("ADODB.connection")
	set rs = server.CreateObject("ADODB.recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
		
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

	Set rs = cm.Execute 

	set cm=nothing	

	if (rs.EOF and rs.BOF) then
		'Bug 26641/ Task 26642 - Harris, Valerie
		Response.Write("Your user name could not be found, can not save. Please try again.")
        Response.End()

	else
		strFrom = rs("Email")
		currentuserid = rs("ID")
	end if 
	strTo = strFrom & ";APJ-RCTO.SC@hp.com;TWNPDCNBCommodityTechnology@hp.com;kidwell.proceng@hp.com"
	rs.Close

	'Get Dev Team
	strSQL = "Select distinct e.email " & _
			 "from deliverableversion v with (NOLOCK), employee e with (NOLOCK) " & _
			 "where v.id in (" & strVersionIDList & ") " & _
			 "and v.developerid = e.id  " & _
			 "Union  " & _
			 "Select e.email " & _
			 "from deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), employee e with (NOLOCK) " & _
			 "where v.id in (" & strVersionIDList & ") " & _
			 "and r.devmanagerid = e.id " & _
			 "and r.id = v.deliverablerootid"
	rs.Open strSQL,cn,adOpenStatic
	do while not rs.EOF
		strTO = strTo & ";" & rs("Email")
		rs.MoveNext
	loop
	rs.Close
	
	'Get ODM PMs
    if len(strPulsarProductList) > 0 then 
	    strSQL = "Select distinct e.email " & _
			     "from productversion v with (NOLOCK) " &_
                 "inner join employee e with (NOLOCK) on v.pdeid = e.id " & _
                 "inner join productversion_release pvr with (NOLOCK) on pvr.ProductVersionID = v.ID " &_
			     "where pvr.id in (" & strPulsarProductList & ") " 
			 
	    rs.Open strSQL,cn,adOpenStatic
	    do while not rs.EOF
		    strTO = strTo & ";" & rs("Email")
		    rs.MoveNext
	    loop
	    rs.Close
	end if

    if len(strLegacyProductList) > 0 then 
	    strSQL = "Select distinct e.email " & _
			     "from productversion v with (NOLOCK), employee e with (NOLOCK) " & _
			     "where v.id in (" & strLegacyProductList & ") " & _
			     "and v.pdeid = e.id"
			 
	    rs.Open strSQL,cn,adOpenStatic
	    do while not rs.EOF
		    strTO = strTo & ";" & rs("Email")
		    rs.MoveNext
	    loop
	    rs.Close
    end if

	'Format Email Body		
	strSQL = "Select v.id, r.name, v.Version, v.revision, v.pass, v.modelnumber, v.partnumber "& _
			 "FROM deliverableroot r with (NOLOCK), deliverableversion v with (NOLOCK) " & _
			 "WHERE r.id = v.deliverablerootid " & _
			 "and v.id in (" & strVersionIDList & ") " & _
			 "order by r.name, v.id"
	rs.Open strSQL,cn,adOpenStatic
	do while not rs.EOF
		strVersion = rs("Version") & ""
		if rs("Revision")&"" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass")&"" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if
		strBody = strBody & "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("ID") & """>" & rs("ID") & "</a></TD>"
		strBody = strBody & "<TD>" & rs("Name") & "</TD>"
		strBody = strBody & "<TD nowrap>" & strVersion & "&nbsp;</TD>"
		strBody = strBody & "<TD>" & rs("Modelnumber") & "&nbsp;</TD>"
		strBody = strBody & "<TD nowrap>" & rs("Partnumber") & "&nbsp;</TD>"
		strBody = strBody & "</TR>"
		rs.MoveNext
	loop
	rs.Close	

	strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1 style=""WIDTH:100%""><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD>" & strBody & "</table>"
	strBody = "<font size=2 face=verdana color=black><b>" & request("txtActionText") & "</b></font><BR><BR>" & strBody
	
	strBody = strBody & "<BR><TABLE  style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>Product List</b></TD></TR><TR><TD>" & request("txtProductNames") & "</TD></TR></TABLE>"
	
	Dim RowsUpdated
	strSuccess = "1"
	cn.BeginTrans

    if len(strPulsarSQLMain) > 0 then
	    cn.Execute strPulsarSQLMain
	    if cn.Errors.count > 0	then
		    strSuccess = "0"
	    else
            'Load IDs to Log
	        LogIDList = ""
            if len(strPulsarProductList) > 0 then
	            strSQl = "Select pd.ProductVersionID, pd.DeliverableVersionID, PVR.ID " & _
			             "FROM Product_Deliverable pd with (NOLOCK) " & _
                         "inner join Product_Deliverable_Release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.ID " &_ 
                         "inner join ProductVersion_Release PVR with (NOLOCK) on PVR.ReleaseID = pdr.ReleaseId and PVR.ProductVersionID = pd.ProductVersionID " &_
			             "WHERE pd.DeliverableVersionID in (" & strVersionIDList & ") " &_
                         "and PVR.ID in (" & strPulsarProductList & ") "

	            rs.Open strSQL, cn,adOpenKeyset
	            do while not rs.eof
		            LogIDList = LogIDList & ":" & rs("ProductVersionID") & "," & rs("DeliverableVersionID") & "," & rs("ID")
		            rs.MoveNext
	            loop
	            rs.Close

                if LogIDList <> "" then
		            LogIDList = mid(LogIDList ,2)
	            end if
            end if

		    LogIDArray = split(LogIDList,":")
		    for each LogIDPair in LogIDArray
			    if trim(LogIDPair) <> "" then
				    LogIDArrayParts = split(LogIDPair,",")
				    cn.Execute "INSERT ActionLog (ActionID, UserID,Updated, ProductVersionID, DeliverableVersionID, UserName, Details, ProductVersionReleaseID) " & _
						       "VALUES (31," & currentuserid & ",GetDate()," & LogIDArrayParts(0) & "," & LogIDArrayParts(1) & ",'" & CurrentDomain & "_" & CurrentUser & "','" & LogText & "'," & LogIDArrayParts(2) & ");" ,RowsUpdated
				    if RowsUpdated <> 1 then
					     strSuccess = "0" 
					     exit for
				    end if
			    end if
		    next
	    end if
    end if

    if len(strLegacySQLMain) > 0 then
	    cn.Execute strLegacySQLMain
	    if cn.Errors.count > 0	then
		    strSuccess = "0"
	    else
            'Load IDs to Log
	        LogIDList = ""
            if len(strLegacyProductList) > 0 then
	            strSQl = "Select pd.ProductVersionID, pd.DeliverableVersionID, ID = 0 " & _
			             "FROM Product_Deliverable pd with (NOLOCK) " & _
			             "WHERE pd.DeliverableVersionID in (" & strVersionIDList & ") " &_
                         "and pd.ProductVersionID in (" & strLegacyProductList & ")" 

	            rs.Open strSQL, cn,adOpenKeyset
	            do while not rs.eof
		            LogIDList = LogIDList & ":" & rs("ProductVersionID") & "," & rs("DeliverableVersionID") & "," & rs("ID")
		            rs.MoveNext
	            loop
	            rs.Close

                if LogIDList <> "" then
		            LogIDList = mid(LogIDList ,2)
	            end if
            end if

		    LogIDArray = split(LogIDList,":")
		    for each LogIDPair in LogIDArray
			    if trim(LogIDPair) <> "" then
				    LogIDArrayParts = split(LogIDPair,",")
				    cn.Execute "INSERT ActionLog (ActionID, UserID,Updated, ProductVersionID, DeliverableVersionID, UserName, Details, ProductVersionReleaseID) " & _
						       "VALUES (31," & currentuserid & ",GetDate()," & LogIDArrayParts(0) & "," & LogIDArrayParts(1) & ",'" & CurrentDomain & "_" & CurrentUser & "','" & LogText & "'," & LogIDArrayParts(2) & ");" ,RowsUpdated
				    if RowsUpdated <> 1 then
					     strSuccess = "0" 
					     exit for
				    end if
			    end if
		    next
	    end if
    end if

	if cn.Errors.count > 0 or strSuccess = "0" then
		cn.RollbackTrans
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to update the selected versions.</font>"
	else
		cn.CommitTrans
	end if



	set rs = nothing
	cn.close
	set cn = nothing

	if strSuccess = "1" then
	
		if strCC <> "" then
			strTO = strTo & ";" & strCC 
		end if
	
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")	
		oMessage.From = strFrom
		'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
			oMessage.To = strTo 
		oMessage.Subject = strSubject
				
		oMessage.HTMLBody = strBody '& "<BR><BR><BR><BR><font size=1 face=verdana>" & strTO & "</font>"
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	end if

%>
	<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>
<%@ Language=VBScript %>
<!-- #include file="includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="includes/client/jquery.min.js" type="text/javascript"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
//
// Old function replaced by jquery document on ready function.
//
function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
			window.returnValue = txtStatusName.value;			
			window.parent.close();
			}
		}
}

$(function(){
	var returnValue = $("#txtSuccess").val();
	if (returnValue == "1"){
	    var iframeName = parent.window.name;
	    if (iframeName != '') {
	        parent.window.parent.ClosePropertiesDialog(returnValue);
	    } else {
	        window.returnValue = returnValue;
	        window.close();
	    }
	}
});

//-->
</SCRIPT>

</HEAD>
<BODY>
<%
	dim cn
	dim rs
	dim strSuccess
	dim IDArray
	dim strID
	
	IDArray = split(request("txtPDID"),",")
		
	strSuccess = "1"

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


		'Get User
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
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") 
			CurrentUserEmail = rs("EMail") 
		end if
		rs.Close
		set cm = nothing

	cn.BeginTrans
	
	for each strid in IDArray
	    'Get Email data if needed
        if trim(request("cboDevStatus")) = "2" and trim(request("txtTypeID")) <> "2" then  'Developer Notification Status Updated

            response.write "<BR>" & strID
            rs.open "spGetDeliverableVersionProductPM " & clng(strid),cn
            if rs.eof and rs.bof then
                strTo = ""
                strCC = ""
                strDeliverableName = ""
                strProductName = ""
            else
                response.write "<BR>Deliverable Found"
                if trim(rs("TypeID")) = "1" then
                    strTo = trim(rs("Email") & "")
                    strCC = trim(rs("SystemTeamEmail") & "")
                    strDeliverableName = trim(rs("Deliverablename") & "")
                    strProductName = trim(rs("Productname") & "")
                    strSubject = "Deliverable Version Support Request Rejected"
                    strBody = "The request to support this version of " & strDeliverableName & " on " & strProductName & " was rejected." & "<BR><BR>"
                    if request("txtComments") <> "" then
                        strBody = strBody & "Reason Rejected: " & request("txtComments") & "<BR><BR>"
                    end if
                    strBody = strBody & "Vendor: " & rs("Vendor") & "<BR>"
                    strBody = strBody & "Hardware Version: " & rs("Version") & "<BR>"
                    strBody = strBody & "Firmware Version: " & rs("Revision") & "<BR>"
                    strBody = strBody & "Revision: " & rs("Pass") & "<BR>"
                    strBody = strBody & "Part Number: " & rs("PartNumber") & "<BR>"
                    strBody = strBody & "Model Number: " & rs("ModelNumber") & "<BR>"
                    response.write "<BR>Email Body Prepared"
                else
                    strTo = ""
                    strCC = ""
                    strDeliverableName = ""
                    strProductName = ""
                    response.write "<BR>" & trim(rs("TypeID"))
                end if
            end if
            rs.Close            


        end if

	    set cm = server.CreateObject("ADODB.Command")
	    '&H0004 means stored procedure
	    cm.CommandType =  &H0004
	    cm.ActiveConnection = cn
	    if trim(request("txtTypeID")) = "2" then
		    cm.CommandText = "spUpdateDeveloperTestStatus"
	    else
		    cm.CommandText = "spUpdateDeveloperNotificationStatus"
	        response.write "<BR>Updating " & strID
        end if
    						
	    Set p = cm.CreateParameter("@PDID",3, &H0001)
	    p.Value = clng(strID)
	    cm.Parameters.Append p
    		
	    Set p = cm.CreateParameter("@StatusID",16, &H0001)
	    p.Value = request("cboDevStatus")
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@DeveloperTestNotes", 200,  &H0001,200)
	    p.Value = left(request("txtComments"),200)
	    cm.Parameters.Append p

	    if trim(request("txtTypeID")) <> "2" then
		    Set p = cm.CreateParameter("@Type",16, &H0001)
		    p.Value = 2
		    cm.Parameters.Append p

    	    Set p = cm.CreateParameter("@UserID",3, &H0001)
	        p.Value = clng(CurrentUserID)
	        cm.Parameters.Append p

	    end if
    				
	    cm.Execute rowschanged
    	Set cm=nothing
	    if cn.Errors.count > 0 then
		    strSuccess = "0"
		    exit for
	    end if	

    	if trim(request("cboDevStatus")) = "2" and trim(request("txtTypeID")) <> "2" then 
            'Send Email to System Team if needed
            if (trim(strTo) <> "" or trim(strCC) <> "") then
                response.write "<BR>Preparing Email"
                if strDeliverableName = "" or strproductName = "" then
                    strBody = "ProdDelRootID: " &  clng(strID) & "<BR>"
                    strBody = strBody & "To: " &  strTo & "<BR>"
                    strBody = strBody & "Product: " & strProductName & "<BR>"
                    strBody = strBody & "Root: " &  strDeliverableName & "<BR>"
                    strTo = "max.yu@hp.com"
                    strCC = "max.yu@hp.com;"
                else
                    if strTo="" then
                        strTO = CurrentUserEmail
                    end if
                    strCC = CurrentUserEmail & ";" & strCC 
                    if lcase(trim(strProductname)) = "test product 1.0" then
                        strBody = "This update occured on Test Product 1.0.  The following notifications would have been sent to <BR>TO: " & strTo & "<BR>CC: " & strCC & ".<BR><BR>"  & strBody
                        strTo = "max.yu@hp.com" 
                        strCC = "max.yu@hp.com"
                    end if
                end if

                response.write "<BR>Email Ready: " & strTO & "_" & strCC & "_" & strBody
                if (strTo <> "" or strCC <> "" )  then
                    response.write "<BR>Sending Email"

                    Set oMessage = New EmailWrapper 
	                oMessage.From = CurrentUserEmail
		
	                if strTo="" then
		                oMessage.To = CurrentUserEmail
               		else
		                oMessage.To = strTo 
	                end if
                    oMessage.CC = strCC
                    omessage.bcc = "max.yu@hp.com"
		
	                oMessage.Subject = strSubject
	                oMessage.HTMLBody = "<font size=2 face=verdana color=black>" & strBody & "</font>"
	                oMessage.DSNOptions = cdoDSNFailure
	                oMessage.Send 
	                Set oMessage = Nothing 			
                end if

            end if

        end if

    next
	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	Set p=nothing


	dim strTO
    dim strCC
	dim strFrom
	dim strSubject
	dim strNewStatus
	dim strPMID
	dim strBody
	dim strVersion
	dim strRows
	dim strProductID
    dim strDeliverableName
    dim strProductName
	

    response.write "Checking to see if emails are required."
	if trim(request("txtTypeID")) = "2" then 
        'Test Status Updated
    	'Send Email of the change to the HW PM
		rs.open "spGetDevProductionReleaseEmailInfo " & clng(request("txtPDID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strTo = "max.yu@hp.com"
			strPMID = 0
			strProductID = 0
		else
			strPMID = rs("PMFieldName") & ""
			strPMID = rs(strPMID) & ""
			if rs("DeveloperTestStatus") = 1 then
				strNewStatus = "Approved for Production"
			elseif rs("DeveloperTestStatus") = 2 then
				strNewStatus = "Not Approved for Production"
			else
				strNewStatus = "TBD"
			end if
			
			strVersion = rs("Version") & ""
			if rs("revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision") 
			end if
			if rs("pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass") 
			end if
			
			strProductID = rs("ProductID") & ""
			
			strBody = "<B>" & rs("Vendor") & " " & rs("Deliverable") & " [" & strVersion & "] set to '" & strNewStatus & "' on " & rs("Product") & "</b>" 
			strRows = "<TR>"
			strRows = strRows & "<TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
			strRows = strRows & "<TD>" & rs("Product") & "</TD>"
			strRows = strRows & "<TD>" & rs("Vendor") & "</TD>"
			strRows = strRows & "<TD>" & rs("Deliverable") & "</TD>"
			strRows = strRows & "<TD>" & strVersion & "</TD>"
			strRows = strRows & "<TD>" & rs("PartNumber") & "</TD>"
			strRows = strRows & "<TD>" & rs("ModelNumber") & "</TD>"
			strRows = strRows & "<TD>" & rs("DeveloperTestNotes") & "</TD>"
			strRows = strRows & "</TR>"
			strBody = strBody & "<BR><BR><STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Dev.&nbsp;Comments</b></TD></tr>" & strRows & "</table></font>"
		end if
		rs.Close

		'Lookup PM
		if trim(strPMID) <> "0" and trim(strPMID) <> "" then
			rs.open "spGetEmployeeByID " & clng(strPMID),cn,adOpenStatic
			if rs.EOF and rs.BOF then
				strTo = "max.yu@hp.com"
			else
				strTo = rs("Email") 
			end if
			rs.Close
		else
			strTo = "max.yu@hp.com"
		end if
		
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		if CurrentUserEmail <> "" then
			oMessage.From = CurrentUserEmail 
		else
			oMessage.From = "max.yu@hp.com" 
		end if
		if trim(strProductID) = "100" or CurrentUserID=31 then
			oMessage.To = "max.yu@hp.com"
		else
			oMessage.To =  strTO 
		end if
		oMessage.Subject = "Developer Final Approval Updated"
		oMessage.HTMLBody = "<font size=2 face=verdana color=black>" & strBody & "</font>"
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	end if
'	
	cn.Close
	set rs = nothing
	set cn = nothing

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=txtDevStatus name=txtDevStatus value="<%=request("cboDevStatus")%>">
<INPUT type="hidden" id=txtStatusName name=txtStatusName value="<%=request("txtStatusName")%>">

</BODY>
</HTML>

<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<%if request("Type") = 1 then%>
<TITLE>Target Version</TITLE>
<%else%>
<TITLE>Reject Version</TITLE>
<%end if%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	        if (txtSuccess.value=="1")
	        {
	            window.parent.Close(hdnRowID.value);
		    }
		}
	
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<BR>
<table width=100%><TR><TD align=center>
<font face=verdana size =2>Processing Request.  Please wait...</font>
</td></tr></table>
<%

	dim strVersionID
	dim strProductID
	dim strType
	dim strRejected
	dim cn
	dim rs
	dim IDArray
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
	dim cm
	dim TargetArray
	dim strTO
	dim strFrom
	dim strSubject
	dim strNewStatus
	dim strPMID
	dim strBody
	dim strVersion
	dim strRows
    dim RowIDs

	if request("txtMultiID") = "" then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0

	else
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
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
			CurrentUserEmail = rs("Email") 
		end if
		rs.Close
	
		Response.Write "<BR>UserID:" & CurrentUserID & "<BR>"

		cn.BeginTrans
		
		ProcessArray = split(request("txtMultiID"),",")
		
		Response.Write "<BR>Processing:" & "<BR>"
		strTo = ""
		strBody = ""
		
     

		for i = 0 to ubound(ProcessArray)
            set cm = server.CreateObject("ADODB.Command")
			cm.ActiveConnection = cn

			strID = trim(ProcessArray(i))	
    	
			Response.Write strID & " - " & clng(request("NewValue")) & "<BR>"
			IDArray = split(strID,"_")

            if RowIDs <> "" then
                RowIDs = RowIDs + ","
            end if

            RowIDs = RowIDs + IDArray(1) + "_" + IDArray(2)

			cm.CommandText = "spUpdateDeveloperTestStatus"
			cm.CommandType = &H0004
						
			Set p = cm.CreateParameter("@PDID",3, &H0001)
			p.Value = clng(IDArray(1))
			cm.Parameters.Append p
			
			Set p = cm.CreateParameter("@StatusID",16, &H0001)
			p.Value = clng(request("NewValue"))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@DeveloperTestNotes", 200, &H0001, 256)
			p.Value = trim(left(request("txtComments"),256))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@PDRID",3, &H0001)
			p.Value = clng(IDArray(2))
			cm.Parameters.Append p			
					
			cm.Execute
			Set cm = Nothing

            Response.Write "<BR>preparing data ...<BR>"

	    next

    	if cn.Errors.count > 0 then
			Response.Write "<BR>Failed.<BR>"
			strSuccess = "0"
			cn.RollbackTrans
			Response.Write "<BR>Records were not saved correctly.<BR>"
		else
			strSuccess = "1"
            cn.CommitTrans
            if cn.Errors.count > 0 then
                strSuccess = "0"
                Response.Write "<BR>Records commited error.<BR>"
            else
                strSuccess = "1"
                Response.Write "<BR>Records commited and saved successfully.<BR>"
            end if
        end if

        
            
		if strSuccess = "1" then
                for i = 0 to ubound(ProcessArray)
                    strID = trim(ProcessArray(i))	
    	
			        IDArray = split(strID,"_")

                    if clng(IDArray(2)) = 0 then 
				        rs.open "spGetDevProductionReleaseEmailInfo " & clng(IDArray(1)),cn,adOpenStatic
                    else 
                        rs.open "spGetDevProductionReleaseEmailInfoPulsar " & clng(IDArray(2)),cn,adOpenStatic
                    end if

				    if rs.EOF and rs.BOF then
					    if instr(strTo,"max.yu@hp.com")=0 then
						    strTo = strTo & ";max.yu@hp.com"
					    end if
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

					    strBody = "<B>The following deliverables have been set to '" & strNewStatus & "' on the selected products.</b>" 

					    strRows = strRows & "<TR>"
					    strRows = strRows & "<TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
					    strRows = strRows & "<TD>" & rs("Product") & "</TD>"
					    strRows = strRows & "<TD>" & rs("Vendor") & "</TD>"
					    strRows = strRows & "<TD>" & rs("Deliverable") & "</TD>"
					    strRows = strRows & "<TD>" & strVersion & "</TD>"
					    strRows = strRows & "<TD>" & rs("PartNumber") & "</TD>"
					    strRows = strRows & "<TD>" & rs("ModelNumber") & "</TD>"
					    strRows = strRows & "<TD>" & rs("DeveloperTestNotes") & "</TD>"
					    strRows = strRows & "</TR>"
				    end if
				    rs.Close

				    'Lookup PM
				    if trim(strPMID) <> "0" and trim(strPMID) <> "" then
					    rs.open "spGetEmployeeByID " & clng(strPMID),cn,adOpenStatic
					    if rs.EOF and rs.BOF then
						    if instr(strTo,"max.yu@hp.com")=0 then
							    strTo = strTo & ";max.yu@hp.com"
						    end if
					    else
						    if instr(strTo,rs("Email")&"" )=0 then
							    strTo = strTo & ";" & rs("Email") 
						    end if
					    end if
					    rs.Close
				    else
					    if instr(strTo,"max.yu@hp.com")=0 then
						    strTo = strTo & ";max.yu@hp.com"
					    end if
				    end if
                next

                if cn.Errors.count > 0 then
                    Response.Write "<BR>Getting email information was error.<BR>"
                else
                    Response.Write "<BR>Getting email information was successful.<BR>"
                end if
		end if
		



	    'Send Email of the change to the HW PM
	    if strRows <> "" then 
		    if strTO = "" then
			    strTo = "max.yu@hp.com"
		    else
			    strTo = mid(strTo,2)
		    end if
		
		    strBody = strBody & "<BR><BR><STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Dev.&nbsp;Comments</b></TD></tr>" & strRows & "</table></font>"

		    Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		    'Set oMessage.Configuration = Application("CDO_Config")		
		    if CurrentUserEmail <> "" then
			    oMessage.From = CurrentUserEmail 
		    else
			    oMessage.From = "max.yu@hp.com" 
		    end if
		    if trim(strProductID) = "100" then
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
					            
		set rs = nothing
		set cn = nothing

	end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input type="hidden" id="hdnRowID" name="hdnRowID" value="<%=RowIDs%>"
</BODY>
</HTML>

<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onload() {
	if (typeof(txtSuccess) != "undefined")
	{
	    if (txtSuccess.value != "0") {
	        
	            if (txtAction.value != "0")
	                window.open("../Actions/Action.asp?TicketID=" + txtAction.value, "", "status=0,toolbar=0,location=0,menubar=0,directories=0,resizable=1,height=650,width=655");
	            //window.returnValue = txtSuccess.value;
	            //window.parent.opener='X';
	            //window.parent.open('', '_parent', '');
	            //window.parent.close();	
	            if (IsFromPulsarPlus()) {
	                window.parent.parent.parent.popupCallBack(txtSuccess.value);
	                ClosePulsarPlusPopup();
	            } else {
	            parent.window.parent.ShowTicket_return(txtSuccess.value);	            parent.window.parent.modalDialog.cancel(false);
	        }
	    }
		}
	
}
//-->
</SCRIPT>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</STYLE>
</HEAD>



<BODY  onload="window_onload();">

<%
    dim cn, rs, strNewOwner

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserEmail
	dim CurrentUserID

    dim strNotificationList

    dim strSubject
    dim strBody

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
		CurrentUserID = rs("ID") & ""
        CurrentUserEmail = rs("Email") & ""
	else
		CurrentUserID = ""
        CurrentUserEmail = ""
	end if
	
	rs.Close

    if request("cboCategory") = "" or request("chkCopyTeam") <> "1" then
        strNotificationList = ""
    else
        rs.open "spSupportIssuesDefaultOwnerSelect " & clng(request("cboCategory")),cn,adOpenStatic
        if rs.eof and rs.bof then
            strNotificationList = ""
        else    
            strNotificationList = trim(rs("NotificationList") & "")
        end if
        rs.Close
    end if

    if request("chkCopyMe") = "1" and CurrentUserEmail <> "" then
        if trim(strNotificationList) = "" then
            strNotificationList = CurrentUserEmail
        else
            strNotificationList = CurrentUserEmail & ";" & strNotificationList
        end if
    end if

    if request("txtID") <> "" then
	    cn.BeginTrans

	    set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
	    cm.CommandText = "spSupportIssueUpdate"
	    cm.CommandType = &H0004

	    Set p = cm.CreateParameter("@ID",adInteger, &H0001)
	    p.Value = cint(request("txtID"))
	    cm.Parameters.Append p
	      
	    Set p = cm.CreateParameter("@Summary",200, &H0001,2000)
	    p.Value = left(request("txtSubject"),2000)
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@Details",200, &H0001,2147483647)
	    p.Value = request("txtDetails")
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@OwnerID",adInteger, &H0001)
	    p.Value = cint(request("cboOwner"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@StatusID",adInteger, &H0001)
	    p.Value = cint(request("cboStatus"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ProjectID",adInteger, &H0001)
	    p.Value = clng(request("cboProject"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@CategoryID",adInteger, &H0001)
	    p.Value = clng(request("cboCategory"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@TypeID",adInteger, &H0001)
	    p.Value = clng(request("cboType"))
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@Resolution",200, &H0001,2147483647)
	    p.Value = request("txtResolution")
	    cm.Parameters.Append p


        cm.Execute RowsEffected
        
	    Set cm = Nothing
	    if rowseffected <> 1 then
    		strSuccess = "0"
	    	cn.RollbackTrans
    	else
        	strSuccess = "1"
	    	cn.CommitTrans
	    end if

    end if
    
    'Send the "reassigned" email
    if trim(request("tagOwner")) <> trim(request("cboOwner")) then
            'Lookup new owner Email
            rs.open "spGetEmployeeByID " & clng(request("cboOwner")),cn
            if rs.eof and rs.bof then
                 strNewOwner = "max.yu@hp.com"
            else
                 strNewOwner = rs("Email") & ""
            end if
            rs.Close

            'Prepare EMail

            if request("txtProjectName") = "" then
                strSubject = "Mobile Tools Support Request - (ID: " & request("txtID") & ") - Reassigned to you."
            else
                strSubject =  request("txtProjectName") & " Support Request - (ID: " & request("txtID") & ") - Reassigned to you."
            end if

            strBody = "<style>td{ font-family:verdana; font-size: xx-small;}</style><font size=2 face=verdana>"
            strBody = strBody & "<table bgcolor=ivory cellpadding=2 cellspacing=0 border=1 bordercolor=LightGrey width=""100%"">"
            strBody = strBody & "<tr><td><b>Ticket:</b></td><td width=""100%""><a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/support/ticket.asp?ID=" & server.HTMLEncode(request("txtID")) & """>" &  server.HTMLEncode(request("txtID")) & "</a></td></tr>"
            strBody = strBody & "<tr><td><b>Subject:</b></td><td> " & request("txtSubject") & "</td></tr>"

            strBody = strBody & "<tr><td><b>Request&nbsp;Type:</b></td><td>"
            select case clng(request("cboType"))
            case 1
                strBody = strBody & "Ask a Question"
            case 2
                strBody = strBody & "Report an Issue"
            case 3
                strBody = strBody & "Make a Suggestion"
            case 4
                strBody = strBody & "Request Admin Updates"
            end select
            strBody = strBody & "</td></tr>"

            strBody = strBody & "<tr><td nowrap><b>Working&nbsp;Notes:&nbsp;&nbsp;&nbsp;</b></td><td>" &  replace(request("txtDetails"),chr(13),"<BR>") & "</td></tr>"
            strBody = strBody & "</table></font>"

            'This is temporary
           ' strBody = strBody & "<BR><BR>TO:" & strNewOwner

			Set oMessage = New EmailWrapper 		
	
			if CurrentUserEmail <> "" then
				oMessage.From = CurrentUserEmail
			else
				oMessage.From = "max.yu@hp.com"
			end if
			
			if strNewOwner <> "" then
				oMessage.To= strNewOwner 
			else
				oMessage.To= "max.yu@hp.com;"
			end if
			
			oMessage.Subject = strSubject
									
			oMessage.HTMLBody = strBody

			oMessage.Send 
			Set oMessage = Nothing 

    end if


    'Send closed email.
    if (trim(strNotificationList) <> "" or trim(request("txtNotify")) <> "" ) and trim(request("tagStatus")) <> "2" and trim(request("cboStatus")) = "2" then

            if request("txtProjectName") = "" then
                strSubject = "Mobile Tools Support Request - (ID: " & request("txtID") & ") - Closed"
            else
                strSubject =  request("txtProjectName") & " Support Request - (ID: " & request("txtID") & ") - Closed"
            end if
            
            strBody = "<style>td{ font-family:verdana; font-size: xx-small;}</style><font size=2 face=verdana>"
            if request("txtResolution") <> "" then
                strBody = strBody &  replace(request("txtResolution"),chr(13),"<BR>") & "<BR><BR>"
            end if
            
            strBody = strBody & "<table bgcolor=ivory cellpadding=2 cellspacing=0 border=1 bordercolor=LightGrey width=""100%"">"
            strBody = strBody & "<tr><td><b>Ticket:</b></td><td width=""100%""><a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/support/ticket.asp?ID=" & server.HTMLEncode(request("txtID")) & """>" &  server.HTMLEncode(request("txtID")) & "</a></td></tr>"
            strBody = strBody & "<tr><td><b>Subject:</b></td><td> " & request("txtSubject") & "</td></tr>"
            strBody = strBody & "<tr><td align=top nowrap><b>Working&nbsp;Notes:&nbsp;&nbsp;&nbsp;</b></td><td>" &  replace(request("txtDetails"),chr(13),"<BR>") & "</td></tr>"
            strBody = strBody & "</table></font>"

            'This is temporary
            'strBody = strBody & "<BR><BR>TO:" & server.htmlencode(request("txtNotify"))
            'strBody = strBody & "<BR>CC:" & strNotificationList

			Set oMessage = New EmailWrapper 		
	
			if CurrentUserEmail <> "" then
				oMessage.From = CurrentUserEmail
			else
				oMessage.From = "max.yu@hp.com"
			end if
			
			if trim(request("txtNotify")) <> "" then
				oMessage.To= server.htmlencode(request("txtNotify")) 
			else
				oMessage.To= "max.yu@hp.com"
			end if
			
            if trim(strNotificationList) <> "" then
                oMessage.CC = strNotificationList
            end if

			oMessage.Subject = strSubject
									
			oMessage.HTMLBody = strBody

			oMessage.Send 
			Set oMessage = Nothing 
    end if

    set rs = nothing
    cn.Close
    set cn = nothing

    dim strAction
    if trim(request("chkActionItem")) = "1" and trim(request("txtID")) <> "" then 'and cint(request("cboStatus")) = 2 then
        strAction = trim(request("txtID"))
    else
        strAction = 0
    end if

%>

<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="text" style="display:none" id=txtAction name=txtAction value="<%=strAction%>">
<INPUT type="text" style="display:none" id=txtUpdateBy name=txtUpdateBy value="<%=trim(lcase(Session("LoggedInUser")))%>">

</BODY>
</HTML>





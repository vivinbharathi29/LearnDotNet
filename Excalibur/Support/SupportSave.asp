<%@ Language="VBScript" %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file="../includes/emailwrapper.asp" -->
<html>
<head>
<title></title>
    <script src="../Scripts/PulsarPlus.js"></script>
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--
    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value != "0") {
                if (IsFromPulsarPlus()) {
                    window.parent.parent.parent.popupCallBack(1);
                    ClosePulsarPlusPopup();
                }
                else {
                    if (CheckOpener() === false) {
                        if (typeof parent.window.parent.ClosePropertiesDialog !== "undefined") {
                            parent.window.parent.ClosePropertiesDialog();
                        } else if (typeof parent.window.parent.modalDialog.cancel !== "undefined") {
                            parent.window.parent.AddTicket_return(txtSuccess.value);
                            parent.window.parent.modalDialog.cancel(false);
                        }
                    } else {
                        window.returnValue = txtSuccess.value;
                        window.parent.opener = 'X';
                        window.parent.open('', '_parent', '');
                        window.parent.close();
                    }
                }
            }
        }
    }

    function CheckOpener() {
        //If True, page opened with showModalDialog
        //if False, page opened with JQuery Modal Dialog
        var oWindow = window.dialogArguments;
        return (oWindow == null) ? false : true;
    }
//-->
</script>

<style type="text/css">
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</style>
</head>

<body  onload="window_onload();">

<%
    'Get User
    Dim blnTrackTickets
    Dim strSuccess
    Dim CurrentUser
    Dim CurrentDomain
    Dim CurrentUserID
    Dim CurrentUserEmail
    Dim CurrentUserPartner
    Dim strDetails
    Dim strDefaultOwnerID
    Dim strDefaultOwnerName
    Dim strDefaultOwnerEmail
    Dim strNotificationList
    Dim NewID
    Dim PathArray
    Dim strSubject
    Dim strBody
    Dim strAttachments
    Dim Rowseffected 
    Dim cn
    Dim rs
    Dim cm
    Dim p
	set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") 

	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


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
        CurrentUserPartner = rs("PartnerID") & ""
	else
		CurrentUserID = ""
        CurrentUserEmail = ""
        CurrentUserPartner = ""
	end if
	
	rs.Close

    strDetails = trim(request("txtDetails"))
    if strDetails <> "" then
        strDetails = strDetails & vbcrlf  & vbcrlf & request("txtRequired")
    else
        strDetails = request("txtRequired")
    end if

    rs.open "spSupportIssuesDefaultOwnerSelect " & clng(request("cboCategory")),cn,adOpenStatic
    if rs.eof and rs.bof then
        strDefaultOwnerID = 31
        strDefaultOwnerName = "Max, Yu"
        strDefaultOwnerEmail = "max.yu@hp.com"
        blnTrackTickets = true
        strNotificationList = ""
    else
        strDefaultOwnerID = rs("ID")
        strDefaultOwnerName = rs("Name")
        strDefaultOwnerEmail = rs("Email")
        blnTrackTickets = rs("TrackTickets")
        strNotificationList = trim(rs("NotificationList") & "")
    end if
    rs.Close
    if request("chkCopyMe") <> "" and CurrentUserEmail <> "" then
        if trim(strNotificationList) = "" then
            strNotificationList = CurrentUserEmail
        else
            strNotificationList = CurrentUserEmail & ";" & strNotificationList
        end if
    end if

    NewID = ""
    if request("txtSubject") <> "" and blnTrackTickets then

	    cn.BeginTrans

	    set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
	    cm.CommandText = "spSupportIssuesInsert"
	    cm.CommandType = &H0004
	      
	    Set p = cm.CreateParameter("@Summary",200, &H0001,500)
	    p.Value = left(request("txtSubject"),500)
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@Details",200, &H0001,2147483647)
	    p.Value = strDetails
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@OwnerID",adInteger, &H0001)
	    p.Value = strDefaultOwnerID
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@SubmitterID",adInteger, &H0001)
	    p.Value = clng(CurrentUserID)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ProjectID",adInteger, &H0001)
	    p.Value = clng(request("cboProject"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@CategoryID",adInteger, &H0001)
	    p.Value = clng(request("cboCategory"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@TypeID",adInteger, &H0001)
	    p.Value = clng(request("optType"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Attachment1",200, &H0001,500)
	    p.Value = left(request("txtAttachmentPath1"),500)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Attachment2",200, &H0001,500)
	    p.Value = left(request("txtAttachmentPath2"),500)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Attachment3",200, &H0001,500)
	    p.Value = left(request("txtAttachmentPath3"),500)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@NewID",adInteger, &H0002)
	    cm.Parameters.Append p

        cm.Execute RowsEffected
        
        NewID = cm("@NewID")

	    Set cm = Nothing
	    if rowseffected <> 1 then
    		strSuccess = "0"
	    	cn.RollbackTrans
    	else
        	strSuccess = NewID
	    	cn.CommitTrans
	    end if

    end if
    

    if request("txtSubject") <> "" then
            'Lookup Default Owners Name
            dim strNewOwner

            if strDefaultOwnerID <> "" then
                rs.open "spGetEmployeeByID " & clng(strDefaultOwnerID),cn
                if rs.eof and rs.bof then
                     strNewOwner = ""
                else
                     strNewOwner = rs("Name") & ""
                end if
                rs.Close
                if instr(strNewOwner,", ") > 0 then
                    strNewOwner = mid(strnewOwner,instr(strNewOwner,",")+2)
                end if
            else
                strNewOwner = ""
            end if
            'Create Email
            if trim(request("txtProjectName")) = "" then
                strSubject = "Excalibur Support Request - (ID: " & NewID & ") - "
            else
                strSubject = request("txtProjectName") & " Support Request - (ID: " & NewID & ") - "
            end if
            strSubject = strSubject & "Opened"
            if trim(strNewOwner) <> "" then
                strSubject = strSubject & " and assigned to " & strNewOwner
            end if
            strBody = ""
            if trim(NewID) <> "" then
                if trim(CurrentUserPartner) <> "1" then
                    strBody = strBody & "TICKET#: " & newid & " - <a target=_blank href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/support/ticket.asp?ID=" & newid & """>HP Link</a> | <a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/support/ticket.asp?ID=" & newid & """>Partner Link</a><BR><BR>"
                else
                    strBody = strBody & "TICKET#: <a target=_blank href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/support/ticket.asp?ID=" & newid & """>" & newid & "</a><BR><BR>"
                end if
            end if
            strBody = strBody & request("txtSubject") & "<BR><BR>"
            strBody = strBody & "REQUEST TYPE: "
            select case clng(request("optType"))
            case 0
                strBody = strBody & "Ask a Question"
            case 1
                strBody = strBody & "Report an Issue"
            case 2
                strBody = strBody & "Make a Suggestion"
            case 3
                strBody = strBody & "Request Admin Updates"
            end select
            strBody = strBody & "<BR>"
            strBody = strBody & "PROJECT: " & request("txtProjectName") & "<BR>"
            strBody = strBody & "FEATURE/ISSUE: " & request("txtCategoryName") & "<BR>"

            if trim(request("txtRequired")) <> "" then
                strBody = strBody & replace(request("txtRequired"),chr(13),"<BR>") & "<BR><BR>"
            end if

            if trim(request("txtDetails")) <> "" then
                strBody = strBody & replace(request("txtDetails"),chr(13),"<BR>") & "<BR><BR>"
            end if

            'PathArray = split(request("txtAttachmentPath1"),"\")
            'strBody = strBody & request("txtAttachmentPath1") '& "|" &  PathArray(ubound(PathArray))

            strAttachments = ""
            if trim(request("txtAttachmentPath1")) <> "" then
                PathArray = split(request("txtAttachmentPath1"),"\")
        	   ' oMessage.AddRelatedBodyPart request("txtAttachmentPath1"), PathArray(ubound(PathArray)), cdoRefTypeId
               ' if instr(lcase(PathArray(ubound(PathArray))),".gif")> 0 or instr(lcase(PathArray(ubound(PathArray))),".jpg")> 0 or instr(lcase(PathArray(ubound(PathArray))),".bmp")> 0 or instr(lcase(PathArray(ubound(PathArray))),".png")> 0 then
                '    strAttachments = strAttachments & "<IMG SRC=""" & PathArray(ubound(PathArray)) & """><br>"
                'else
                    strAttachments = strAttachments & "<a href=""" & request("txtAttachmentPath1") & """>" & PathArray(ubound(PathArray)) & "</a>" & "<BR>"
			    'end if
			end if

            if trim(request("txtAttachmentPath2")) <> "" then
                PathArray = split(request("txtAttachmentPath2"),"\")
        	  '  oMessage.AddRelatedBodyPart request("txtAttachmentPath2"), PathArray(ubound(PathArray)), cdoRefTypeId
               ' if instr(lcase(PathArray(ubound(PathArray))),".gif")> 0 or instr(lcase(PathArray(ubound(PathArray))),".jpg")> 0 or instr(lcase(PathArray(ubound(PathArray))),".bmp")> 0 or instr(lcase(PathArray(ubound(PathArray))),".png")> 0 then
               '     strAttachments = strAttachments & "<IMG SRC=""" & PathArray(ubound(PathArray)) & """><br>"
               ' else
                    strAttachments = strAttachments & "<a href=""" & request("txtAttachmentPath2") & """>" & PathArray(ubound(PathArray)) & "</a>" & "<BR>"
			   ' end if
			end if

            if trim(request("txtAttachmentPath3")) <> "" then
                PathArray = split(request("txtAttachmentPath3"),"\")
        	  '  oMessage.AddRelatedBodyPart request("txtAttachmentPath3"), PathArray(ubound(PathArray)), cdoRefTypeId
                'if instr(lcase(PathArray(ubound(PathArray))),".gif")> 0 or instr(lcase(PathArray(ubound(PathArray))),".jpg")> 0 or instr(lcase(PathArray(ubound(PathArray))),".bmp")> 0 or instr(lcase(PathArray(ubound(PathArray))),".png")> 0 then
                 '   strAttachments = strAttachments & "<IMG SRC=""" & PathArray(ubound(PathArray)) & """><br>"
                'else
                    strAttachments = strAttachments & "<a href=""" & request("txtAttachmentPath3") & """>" & PathArray(ubound(PathArray)) & "</a>" & "<BR>"
			    'end if
            end if
            if strAttachments <> "" then
                strAttachments = "<BR>ATTACHMENTS:<BR>" & strAttachments
            end if



            'Temp - Remove before releasing
           ' strAttachments = strAttachments & "<BR><BR>TO:" & server.htmlencode(strDefaultOwnerEmail)
           ' strAttachments = strAttachments & "<BR>CC:" & strNotificationList

			Set oMessage = New EmailWrapper 		
	
			if CurrentUserEmail <> "" then
				oMessage.From = CurrentUserEmail
			else
				oMessage.From = "max.yu@hp.com"
			end if
			
			if strDefaultOwnerEmail <> "" then
				oMessage.To= strDefaultOwnerEmail '"max.yu@hp.com" 'strDefaultOwnerEmail
			else
				oMessage.To= "max.yu@hp.com"
			end if
			
            if trim(strNotificationList) <> "" then
                'oMessage.CC = strNotificationList
                oMessage.BCC = strNotificationList
            end if

			oMessage.Subject = strSubject
									
			oMessage.HTMLBody = "<font face=verdana size=2>" & strBody & strAttachments & "</font>"

			oMessage.Send 
			Set oMessage = Nothing 
    end if

    set rs = nothing
    cn.Close
    set cn = nothing
%>

<input type="text" style="display:" id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>" />
<input type="text" style="display:none" id="txtUpdateBy" name="txtUpdateBy" value="<%=CurrentUser%>" />

</body>
</html>





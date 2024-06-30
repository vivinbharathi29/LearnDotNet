<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value != "0") {
                if (CheckOpener() === false) {
                    if (typeof parent.window.parent.ClosePropertiesDialog !== "undefined") {
                        parent.window.parent.AddTicket_return(txtSuccess.value);
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

    function CheckOpener() {
        //If True, page opened with showModalDialog
        //if False, page opened with JQuery Modal Dialog
        var oWindow = window.dialogArguments;
        return (oWindow == null) ? false : true;
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
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID
    dim NewID
    dim strDefaultOwnerID
    dim strDefaultOwnerName
    dim strDefaultOwnerEmail
    dim strDetails
    dim blnTrackTickets
    dim strNotificationList
    dim PathArray
    dim strSubject
    dim strBody
    dim strAttachments

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
	else
		CurrentUserID = ""
	end if
	
	rs.Close

    NewID = ""

   ' if false then
	    cn.BeginTrans

	    set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
	    cm.CommandText = "spSupportLogCancel"
	    cm.CommandType = &H0004
	      
	    Set p = cm.CreateParameter("@Summary",200, &H0001,500)
	    p.Value = left(request("txtCancelSummary"),500)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@WorkflowStep",adInteger, &H0001)
	    p.Value = clng(request("txtCancelStep"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ProjectName",200, &H0001,120)
	    p.Value = left(request("txtCancelProject"),120)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@CategoryName",200, &H0001,120)
	    p.Value = left(request("txtCancelCategory"),120)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
	    p.Value = CurrentUserID
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

   ' end if

    set rs = nothing
    cn.Close
    set cn = nothing
%>

<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>





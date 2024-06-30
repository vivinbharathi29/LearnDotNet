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
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value != "0")
			{
			//window.returnValue = txtSuccess.value;
		    //window.parent.close();

		    parent.window.parent.AddArticle_return(txtSuccess.value);		    parent.window.parent.modalDialog.cancel(false);

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



<BODY onload="window_onload();">

<%
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserEmail
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
        CurrentUserEmail = rs("Email") & ""
	else
		CurrentUserID = ""
        CurrentUserEmail = ""
	end if
	
	rs.Close

    if request("IsValidation") = "true" then
        rs.Open "Select * from SupportArticle where Title = '" + request("Title") + "' and ID <> '"+ request("Id") + "'" ,cn,adOpenForwardOnly

        if not rs.EOF Then
			Response.Write "TitleExists"
		else
			Response.Write "TitleDoesNotExist"
		end If
				        
		rs.Close
    else
        NewID = ""

	    cn.BeginTrans

	    set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
        if trim(request("txtID")) = "" then
	        cm.CommandText = "spSupportArticleInsert"
    	    cm.CommandType = &H0004
        else
	        cm.CommandText = "spSupportArticleUpdate"
    	    cm.CommandType = &H0004
            NewID = clng(request("txtID"))

	        Set p = cm.CreateParameter("@ID",adInteger, &H0001)
    	    p.Value = NewID
	        cm.Parameters.Append p

        end if
	      
	    Set p = cm.CreateParameter("@Title",200, &H0001,2000)
	    p.Value = left(request("txtTitle"),2000)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@AuthorID",adInteger, &H0001)
	    p.Value = clng(request("cboOwner"))
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@SupportCategoryID",adInteger, &H0001)
	    p.Value = clng(request("cboCategory"))
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@ArticleText",200, &H0001,2147483647)
	    p.Value = request("txtArticleText")
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ArticleURL",200, &H0001,2000)
	    p.Value = left(request("txtURL"),2000)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Keywords",200, &H0001,2000)
	    p.Value = left(request("txtKeywords"),2000)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@StatusID",adInteger, &H0001)
	    p.Value = clng(request("cboStatus"))
	    cm.Parameters.Append p


        if trim(request("txtID")) = "" then
    	    Set p = cm.CreateParameter("@NewID",adInteger, &H0002)
    	    cm.Parameters.Append p
        end if 

        cm.Execute RowsEffected
        
        if trim(request("txtID")) = "" then
            NewID = cm("@NewID")
        end if

	    Set cm = Nothing

	    if rowseffected <> 1 then
    	    strSuccess = "0"
	        cn.RollbackTrans
        else
            strSuccess = NewID
	        cn.CommitTrans
	    end if
    

        set rs = nothing
        cn.Close
        set cn = nothing
    end if
%>

<INPUT type="text" style="display:" id=txtSuccess name=txtSuccess value="<%=NewID%>">

</BODY>
</HTML>





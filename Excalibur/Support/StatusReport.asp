<%@ Language=VBScript %>
<!-- #include file = "../includes/noaccess.inc" -->
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

    function SetStartDate(strOldDate) {
        var strDate;
        strDate = window.showModalDialog("../mobilese/today/caldraw1.asp", strOldDate, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strDate) != "undefined") {
            window.location.href = ReplaceURLParameter(window.location.href, "StartDate", strDate);
        }
    }


    function SetEndDate(strOldDate) {
        var strDate;
        strDate = window.showModalDialog("../mobilese/today/caldraw1.asp", strOldDate, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strDate) != "undefined") {
            window.location.href = ReplaceURLParameter(window.location.href, "EndDate", strDate);
        }
    }

    function ReplaceURLParameter(strURL, strKey, strValue) {
        if (strURL.indexOf("?") == -1)
            strURL = strURL + "?";

        var blnFound = false;
        var NewParameters = "";
        var MyArray = strURL.split("?");
        var URL = MyArray[0];
        var Parameters = MyArray[1];

        if (URL) {
            var MyArray = Parameters.split("&");
            for (var i in MyArray) {
                NewParameters = NewParameters + "&";
                if (MyArray[i] == "") {
                    NewParameters = NewParameters + strKey + "=" + strValue;
                    blnFound = true;
                }
                else if (MyArray[i].indexOf(strKey) == -1)
                    NewParameters = NewParameters + MyArray[i];
                else {
                    NewParameters = NewParameters + strKey + "=" + strValue;
                    blnFound = true;
                }
            }
        }
        if (!blnFound)
            NewParameters = NewParameters + "&" + strKey + "=" + strValue;

        return URL + "?" + NewParameters.substr(1); 

    }
//-->
</SCRIPT>

<STYLE>
	h3
	{
	FONT-SIZE: small;
	FONT-FAMILY: Verdana;
	}
	
	LI
	{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	}
	BODY
	{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	}
	td
	{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	}
A:link
{
    COLOR: blue;
}
A:visited
{
    COLOR: blue;
}

A:hover
{
    COLOR: red;
}	
</STYLE>

</HEAD>


<BODY>
<h3>Status Report</h3>
<%

    dim StartDate, EndDate
    dim cn, rs, strSQL
    dim ResolutionArray
    dim strRow

    if trim(request("StartDate")) = "" then
        StartDate = formatdatetime(Now()-7,vbshortdate) 
    else
        StartDate = formatdatetime(cdate(request("StartDate")),vbshortdate) 
    end if
    if trim(request("EndDate")) = "" then
        EndDate = formatdatetime(Now(),vbshortdate)
    else
        EndDate = formatdatetime(cdate(request("EndDate")),vbshortdate)
    end if

    Response.write "<a href=""javascript:SetStartDate('" & StartDate & "');"">" & StartDate & "</a>" & " - " & "<a href=""javascript:SetEndDate('" & EndDate & "');"">" & EndDate & "</a>"

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    dim CurrentUserID
    dim CurrentUser
    dim CurrentDomain

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
	else
		CurrentUserID = "0"
	end if
	
	rs.Close

    if trim(request("UserID")) <> "" then
        strSQL = "spSupportStatusSelect '" & cdate(StartDate) & "','" & cdate(EndDate) & "'," & clng(request("UserID"))
    else    
        strSQL = "spSupportStatusSelect '" & cdate(StartDate) & "','" & cdate(EndDate) & "'," & clng(CurrentUserID)
    end if
    rs.open strSQL,cn
    if not (rs.eof and rs.bof) then
        response.write "<ul>"
    end if
    do while not rs.eof
        if lcase(trim(rs("StatusType") & "")) = "ticket" then
            response.write "<BR><li><a target=_blank href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/support/Ticket.asp?ID=" & rs("ID") & """>" & rs("StatusType") & " " & rs("ID") & "</a>: " & rs("Summary")  & "</li><ul>"
        else
            response.write "<BR><li><a target=_blank href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/mobilese/today/actionreportmain.asp?Action=0&ID=" & rs("ID") & "&Type=2"">" & rs("StatusType") & " " & rs("ID") & "</a>: " & rs("Summary")  & "</li><ul>"
        end if
        response.write "<table><tr><td><u>FROM:</u> " & rs("Submitter") & "&nbsp;&nbsp;&nbsp;</td><td><u>CREATED:</u> " & rs("Opened") & "&nbsp;&nbsp;&nbsp;</td><td><u>CLOSED:</u> " & rs("Closed") & "&nbsp;&nbsp;&nbsp;</td></tr></table><u>RESOLUTION:</u><br>"
        ResolutionArray = split(rs("Resolution"),vbcrlf)
        for each strRow in ResolutionArray
            if len(strRow) > 20 or len(rs("Resolution") & "") < 40 then
                response.write "<li>" & strRow & "</li>"
            end if
        next
        response.write "</ul>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
        response.write "</ul>"
    end if
    rs.close    

    set rs = nothing
    cn.Close
    set cn = nothing
%>
</BODY>
</HTML>





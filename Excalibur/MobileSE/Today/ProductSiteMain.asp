<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<TITLE>Choose Product RCTO Sites</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<%

if request("ID") = ""  then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim strLoaded
	
	strLoaded = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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
	

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
	end if
	rs.Close
	
%>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b>Choose RCTO Sites for 
<%
    rs.open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
    if rs.eof and rs.bof then
        response.write "Product"    
    else
        response.write rs("Name")  
    end if
    rs.close

'    rs.open "spGetRCTOSites4Product " & clng(request("ID")),cn,adOpenForwardOnly
'    if not (rs.eof and rs.bof) then
'        strLoaded = replace(rs("RCTOSites") & ""," ","")
'    end if
'    rs.close
 
    strLoaded = replace(request("Sites") & ""," ","")
%>
</b></font>
<form ID=frmMain method=post action="ProductSiteSave.asp">
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="150" nowrap valign=top><b>RCTO Sites:</b>&nbsp;</td>
		<td width="100%">
        <%  
            dim strName
            rs.open "spListRCTOSites",cn,adOpenForwardOnly
            do while not rs.eof
                if instr("," & strLoaded & ",","," & rs("Name") & ",") > 0 then 'strLoaded
                    response.write "<input checked id=""chkSite"" name=""chkSite"" type=""checkbox"" value=""" & rs("ID") & """ SiteName=""" & rs("Name") &  """> " & rs("Name") & "<BR>"
                else
                    response.write "<input id=""chkSite"" name=""chkSite"" type=""checkbox"" value=""" & rs("ID") & """ SiteName=""" & rs("Name") &  """> " & rs("Name") & "<BR>"
                end if
                rs.movenext
            loop
            rs.close
            if trim(strLoaded) <> "" then
                strLoaded = mid(strLoaded,2)
            end if
        %>
		</td>
	</tr>
</table>
    <INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
    <input id="txtSiteList" name="txtSiteList" type="hidden" value="">
    <input id="txtSiteLoaded" name="txtSiteLoaded" type="hidden" value="<%=strLoaded%>">
</form>
<%

	set rs = nothing
	set cn = nothing
end if


%>

</BODY>
</HTML>

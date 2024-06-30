<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	%>
	
<HTML>
<HEAD>
<TITLE>Reject Feature Request</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
}

function cmdOK_onclick() {
    if (RejectFeature.txtReason.value == "") {
        alert("You must enter a reason why you are rejecting this request.");
        RejectFeature.txtReason.focus();
    }
    else
        RejectFeature.submit();
}

function window_onload() {
    RejectFeature.txtReason.focus();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory onload="window_onload();">
<%
    

if request("ID") = "" then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
    strID = clng(request("ID"))

    dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserGroup
	
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
		CurrentUserGroup = rs("WorkgroupID") & ""
	end if
	rs.Close

    rs.open "spPULSAR_Today_GetFeatureName " & clng(strID),cn
    if rs.eof and rs.bof then
	    strFeatureName = "[NOT FOUND]"
    else
	    strFeatureName =  rs("FeatureName") & ""
    end if
    rs.close
    if trim(strFeatureName) = "" then
        strFeatureName = "[Feature Name Not Specified]"
    end if
    if trim(strFeatureName) = "[NOT FOUND]"  then
        response.Write "Unable to find the selected feature"
    else

%>
<link rel="stylesheet" type="text/css" href="../../Style/programoffice.css">
<form ID=RejectFeature method=post action="RejectFeatureRequestSave.asp">
<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
<table>
	<TR><TD><font size=3 face=verdana><B>Reject Feature Request:</b></font>&nbsp;
</td></tr><TR><TD>
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
			<td width="150" nowrap><b>Feature:</b>&nbsp;</td>
		<td width="100%"><%=strFeatureName%></td>
	</tr>

	
	<tr>
		<td width="150" nowrap valign=top><b>Reason Rejected:&nbsp;<font color="red">*</font></b></td>
		<td width="100%"><TEXTAREA rows=3 cols=60 id=txtReason name=txtReason></TEXTAREA>
		</td>
	</tr>
	
</table>
</td></tr>
<TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>

</table>
</form>
    <%
    end if
	set rs = nothing
	set cn = nothing
end if


%>

</BODY>
</HTML>

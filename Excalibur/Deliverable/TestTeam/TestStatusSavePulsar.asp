<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

$(document).ready(function () {
    if ($("#txtKeepItOpen").val() == "false") {
        if (parent.window.parent.loadDatatodiv != undefined) {
            parent.window.parent.closeExternalPopup();
        }
        else if (IsFromPulsarPlus()) {
            window.parent.parent.parent.TestStatusCallback($("#txtSuccess").val());
            ClosePulsarPlusPopup();
        } else {
            window.parent.SetNewStatus($("#txtSuccess").val(), $("#txtTodayPageSection").val(), $("#txtFieldID").val(), $("#txtRowID").val());
            window.parent.Close();
        }
    }
    else {
        window.parent.SetNewStatus($("#txtSuccess").val(), $("#txtTodayPageSection").val(), $("#txtFieldID").val(), $("#txtRowID").val());
        document.location = txtRedirect.value;
        window.parent.repositionParentWindow();
    }
});

//-->
</SCRIPT>

<BODY LANGUAGE=javascript>
<%
	dim cn
	dim cm
	dim blnErrors
	dim blnKeepItOpen
    dim strRedirect

    blnKeepItOpen = request("txtKeepItOpen")
    strRedirect = request("txtRedirect")

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	
	'Get User
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

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
    response.Write clng(request("txtProdDelRelID")) & "<BR>"
    response.write clng(request("txtProductID")) & "<BR>"
    response.write clng(request("txtVersionID")) & "<BR>"
    response.write clng(request("txtFieldID")) & "<BR>"
    response.write clng(CurrentUserID) & "<BR>"
    response.write left(CurrentDomain + "_" + Currentuser,80) & "<BR>"
    response.write clng(request("cboStatus")) & "<BR>"
    response.write left(request("txtNotes"),200) & "<BR>"
    response.write trim(request("txtReceived")) & "<BR>"
		
	blnErrors = false		
		
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "spUpdateTestLeadStatusPulsar"	

	Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
	p.Value = clng(request("txtProductID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DeliverableID", 3,  &H0001)
	p.Value = clng(request("txtVersionID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FieldID", 3,  &H0001)
	p.Value = clng(request("txtFieldID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UserID", 3,  &H0001)
	p.Value = clng(CurrentUserID)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Username", 200,  &H0001,80)
	p.Value = left(CurrentDomain + "_" + Currentuser,80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
	p.Value = clng(request("cboStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Notes", 200,  &H0001,200)
	p.Value = left(request("txtNotes"),200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UnitsReceived", 2,  &H0001)
	if trim(request("txtReceived")) = "" then
		p.Value = null
	else
		p.Value = clng(request("txtReceived"))
	end if
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
	p.Value = clng(request("txtProdDelRelID"))
	cm.Parameters.Append p

	cm.Execute rowschanged

	set cm=nothing
		
	cn.Close
	set rs = nothing
	set cn = nothing
	
	
	
	
%>

<input type="hidden" id="txtRedirect" name="txtRedirect" value="<%=strRedirect%>" />
<input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="<%=blnKeepItOpen%>" />
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=request("txtTodayPageSection")%>" />
<input type="hidden" id="txtFieldID" name="txtFieldID" value="<%=request("txtFieldID")%>" />
<input type="hidden" id="txtRowID" name="txtRowID" value="<%=request("txtRowID")%>" />
</BODY>
</HTML>

<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined")
	{
		if (txtSuccess.value == "1")
		{
		    //close window
		    if (parent.window.parent.loadDatatodiv != undefined) {
		        parent.window.parent.closeExternalPopup();
		    }
			//window.returnValue = txtSuccess.value;
			//window.parent.Close();
		    else if (IsFromPulsarPlus()) {
			    window.parent.parent.parent.TestStatusCallback(txtSuccess.value);
			    ClosePulsarPlusPopup();
			}
			else 
			{
			    window.parent.Close();
			}

		}
		//else
		//	document.write ("<BR><font size=2 face=verdana>Unable to update test status.</font>");
	}
	//else
	//	document.write ("<BR><font size=2 face=verdana>Unable to update test status.</font>");
}


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<%
	strSuccess = "1"

	dim cn
	dim cm
	dim blnErrors
	
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
	
    response.write clng(request("txtProductID")) & "<BR>"
    response.write clng(request("txtVersionID")) & "<BR>"
    response.write clng(request("txtFieldID")) & "<BR>"
    response.write clng(CurrentUserID) & "<BR>"
    response.write left(CurrentDomain + "_" + Currentuser,80) & "<BR>"
    response.write clng(request("cboStatus")) & "<BR>"
    response.write left(request("txtNotes"),200) & "<BR>"
    response.write trim(request("txtReceived")) & "<BR>"
		
	cn.BeginTrans
	blnErrors = false
		
		
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "spUpdateTestLeadStatus"	

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


	cm.Execute rowschanged

	set cm=nothing

		
	if  rowschanged <> 1 then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.CommitTrans
	end if
		
	cn.Close
	set rs = nothing
	set cn = nothing
	
	
	
	
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

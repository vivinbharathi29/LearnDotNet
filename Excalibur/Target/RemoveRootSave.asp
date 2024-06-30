<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                //close window
                if (parent.window.parent.loadDatatodiv != undefined) {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp('Deliverables');
                }
                else {
                    if (parent.window.parent.document.getElementById('modal_dialog')) {
                        parent.window.parent.RootRemoveResults('1');
                        parent.window.parent.modalDialog.cancel();
                    } else {
                        window.returnValue = 1;
                        window.close();
                    }
                }
            }
        }
    }

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	dim cn
	dim rs
	dim blnSuccess
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	blnSuccess = false
	
	
	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserEmail
	
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
		CurrentUserEmail = rs("email")
	end if
	if CurrentUserID = "" then
		CurrentUserID = 0
		CurrentUserEmail = ""
	end if
	rs.Close   		
	
	if request("txtDeliverableID") <> "" and request("txtproductID") <> "" and CurrentUserID <> 0 then

		Response.Write "saving"

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn	
		cm.CommandText = "spRemoveDelRootFromProd"	
			
		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = clng(request("txtproductID"))
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@RootID", 3,  &H0001)
		p.Value = clng(request("txtDeliverableID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
		p.Value = clng(CurrentUserID)
		cm.Parameters.Append p

		cm.Execute rowschanged
		
		Response.Write "<br><br>"
		if cn.Errors.count = 0 then
			blnSuccess = true
		end if
					
		set cm = nothing			
			
		
		if request("chkNotify")	<> "" then
			dim DevManagerEmail
			dim strRootDeliverableName
			
			DevManagerEmail = ""
			rs.Open "spGetRootProperties " &  clng(request("txtDeliverableID")) ,cn,adOpenForwardOnly
			if not (rs.EOF and rs.BOF) then
				strRootDeliverableName = rs("name")
				DevManagerEmail = rs("DevManagerEmail") & ""
			end if
			rs.close		
	
			
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")		
	
			if CurrentUserEmail <> "" then
				oMessage.From = CurrentUserEmail
			else
				oMessage.From = "max.yu@hp.com"
			end if
			
			if DevManagerEmail <> "" then
				oMessage.To= DevManagerEmail
			else
				oMessage.To= "max.yu@hp.com"
			end if
			
			oMessage.Subject = strRootDeliverableName & " has been removed from product " & request("txtProductName")
									
			oMessage.HTMLBody = "<font face=verdana size=2>Product: " & request("txtProductName") & "<BR>" & "Root Deliverable: " & strRootDeliverableName & "<BR>" & "Reason for Removal: " & request("txtReason") & "<BR><BR>" & "PM has removed " & request("txtProductName") & " from the supported list of " & strRootDeliverableName & "</font>"
			
			oMessage.Send 
			Set oMessage = Nothing 
		
		end if
	
	end if
	set rs=nothing
	cn.Close
	set cn = nothing
%>
<%if blnSuccess then%>
<INPUT type="text"  style="display:none" id=txtSuccess name=txtSuccess value="1">
<%else%>
<INPUT type="text"  style="display:none" id=txtSuccess name=txtSuccess value="0">
<%end if%>

</BODY>
</HTML>

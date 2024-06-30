<%@ Language=VBScript %>

<% Option Explicit%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--



    function window_onload(pulsarplusDivId) {
        var OutArray = new Array();

        if (txtSuccess.value == "1") {
            //close window
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                parent.window.parent.closeExternalPopup();
            }
            else {
                OutArray[0] = txtPart.value;
                OutArray[1] = txtImage.value;

                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    //save value and return to parent page: ---
                    parent.window.parent.DisplayPIPropertiesResults(OutArray);
                    parent.window.parent.modalDialog.cancel();
                } else {
                    window.returnValue = OutArray;
                    window.parent.close();
                }
            }
        } else {
            document.write("Unable to save Properties.");
        }
    }
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">


<%


	dim i
	dim cn
	dim cm
	dim p
	dim rowschanged
	dim FoundErrors
	dim strPartNumber	
	dim strImage
	dim rs	
	dim CurrentUser
	dim CurrentUserID

	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


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

	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") 
	end if
	rs.Close
	set rs = nothing


	cn.BeginTrans

	FoundErrors = false	

	
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "spUpdatePartNumber"	

	Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
	p.Value = request("txtVersionID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
	p.Value = request("txtProductID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PartNumber", 200, &H0001, 50)
	p.value = left(request("txtPartNumber"),50)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UserID", 3,  &H0001)
	p.Value = CurrentUserID
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@PreinstallInternalRevSkipped", 3,  &H0001)
	if request("chkSkip") = "" then
	    p.Value = null
    else
	    p.Value = clng(request("chkSkip"))
    end if
    cm.Parameters.Append p

	if request("chkInImage") <> request("chkInImageTag") then
		Set p = cm.CreateParameter("@InImage", 16,  &H0001)
		if request("chkInImage") = "on" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p
	end if
	
	cm.Execute rowschanged

	if rowschanged <> 1 then
		FoundErrors = true
	end if
		
	set cm = nothing
		

	if FoundErrors then
		cn.RollbackTrans
		%><INPUT type="text" id=txtSuccess name=txtSuccess value="0"><%
	else
		cn.CommitTrans
		%><INPUT type="text" id=txtSuccess name=txtSuccess value="1"><%
	end if
	
	cn.Close
	set cn = nothing
	set p = nothing
	
	if request("chkInImage") = "on" then
		strImage = "X"
	else
		strImage = "&nbsp;"
	end if
	
%>
<INPUT type="hidden" id=txtPart name=txtPart value="<%=request("txtPartNumber")%>">
<INPUT type="hidden" id=txtImage name=txtImage value="<%=strImage%>">
</BODY>
</HTML>

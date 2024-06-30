<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/jquery-1.10.2.js"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined"){
	    if (txtSuccess.value == "1") {
	        //close window
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	            parent.window.parent.reloadFromPopUp(pulsarplusDivId);
	        }
	        else if (IsFromPulsarPlus()) {
	            //window.parent.parent.parent.popupCallBack(1);
	            ClosePulsarPlusPopup();
	        }
	        else {
	            if (parent.window.parent.document.getElementById('modal_dialog')) {
	                //save value and return to parent page: ---
	                if (typeof parent.window.parent.AdvancedTargetResult == 'function') {
	                    parent.window.parent.AdvancedTargetResult('1');
	                    parent.window.parent.modalDialog.cancel();
	                } else {
	                    parent.window.parent.modalDialog.cancel(true);
	                }
	            } else {
	                window.returnValue = 1;
	                window.close();
	            }
	        }
		}
	}
	//else
	//	{
	//	document.write ("Unable to update Targets.  An unexpected error occurred.");
    //	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">

<%

	dim ProdArray
	dim StatusArray
	dim StatusTagArray
	dim VersionArray
	dim i
	dim blnPMAlert
	dim blnTargeted
	dim blnRejected
	dim cn
	dim rs
	dim CurrentUserID
	dim cm
	dim blnSuccess
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
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

	set cm=nothing
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") 
	end if
	rs.Close
	
	
	ProdArray = split(request("txtID"),",")
	VersionArray = split(request("txtVersionID"),",")
	StatusArray = split(request("cboStatus"),",")
	StatusTagArray = split(request("cboStatusTag"),",")

    dim xmlReleaseIDs : xmlReleaseIDs = ""

	cn.BeginTrans
	blnSuccess = true
	for i = 0 to ubound(ProdArray)
		blnTargeted = 0
		blnPMAlert=1
		blnRejected=0

		
		if (request("txtNotes" & trim(i)) <> request("txtNotesTag" & trim(i))) or (StatusArray(i) <> StatusTagArray(i)) then 
			select case trim(StatusArray(i))
			case "New"
				blnPMAlert=1
				blnTargeted = 0
				blnRejected=0
			case "Targeted"
				blnPMAlert=0
				blnTargeted = 1
				blnRejected=0
			case "Available"
				blnPMAlert=0
				blnTargeted = 0
				if trim(StatusTagArray(i)) = "New" then
					blnRejected = 1
				else
					blnRejected = 0
				end if
			end select

            'add the releases if there is only one product selected, other wise the release will be added from the release pop up screen
                   
            if request("txtNoOfReleases") = "1" then
                xmlReleaseIDs = "<?xml version='1.0' encoding='iso-8859-1' ?><ReleaseInfo><Release><ReleaseID>" & request("TargetedReleases_" & trim(i)) & "</ReleaseID><TargetNotes></TargetNotes></Release></ReleaseInfo>"
                
                set cm = server.CreateObject("ADODB.Command")

		        cm.ActiveConnection = cn
		        cm.CommandText = "usp_ProductDeliverable_TargetReleases"
		        cm.CommandType = &H0004
		
		        Set p = cm.CreateParameter("@p_intDeliverableVersionID",adInteger, &H0001)
			    p.Value = clng(trim(VersionArray(i)))
		        cm.Parameters.Append p
			
		        Set p = cm.CreateParameter("@p_intProductVersionID",adInteger, &H0001)
			    p.Value = clng(request("txtProductID"))
		        cm.Parameters.Append p   
		
			    Set p = cm.CreateParameter("@p_xmlReleaseInfo",adLongVarChar, &H0001,len(xmlReleaseIDs))
			    p.Value = xmlReleaseIDs
			    cm.Parameters.Append p

                Set p = cm.CreateParameter("@p_UserID",adInteger, &H0001)
			    p.Value = CurrentUserID
		        cm.Parameters.Append p

                response.flush
		        cm.Execute
			    Set cm = Nothing

                if cn.Errors.count > 0 then
				    blnSuccess = false
				    cn.RollbackTrans
				    exit for
			    end if
            end if

            if blnSuccess then
			    set cm = server.CreateObject("ADODB.Command")

		        cm.ActiveConnection = cn
		        cm.CommandText = "spTargetAdvanced"
		        cm.CommandType = &H0004
		
		        Set p = cm.CreateParameter("@DeliverableVersionID",adInteger, &H0001)
			    p.Value = clng(trim(VersionArray(i)))
		        cm.Parameters.Append p
			
		        Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
			    p.Value = clng(request("txtProductID"))
		        cm.Parameters.Append p
			
		        Set p = cm.CreateParameter("@TargetValue",adBoolean, &H0001)
			    p.Value = blnTargeted
		        cm.Parameters.Append p
			
		        Set p = cm.CreateParameter("@PMAlert",adBoolean, &H0001)
			    p.Value = blnPMAlert
		        cm.Parameters.Append p
			    
		        Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
			    p.Value = CurrentUserID
		        cm.Parameters.Append p
		
			    Set p = cm.CreateParameter("@Rejected",adBoolean, &H0001)
			    p.Value = blnRejected
			    cm.Parameters.Append p
			
			    Set p = cm.CreateParameter("@TargetNotes",adVarChar, &H0001,255)
			    p.Value = left(request("txtNotes" & trim(i)),255)
			    cm.Parameters.Append p

                response.flush
		        cm.Execute
			    Set cm = Nothing
	
			    if cn.Errors.count > 0 then
				    blnSuccess = false
				    cn.RollbackTrans
				    exit for
			    end if    
            end if    			
		end if
	next

	if blnSuccess then
		cn.CommitTrans
	end if	

    
    'Update ProductComponent link in Sudden Imact
    on error resume next
	for i = 0 to ubound(ProdArray)
	    if clng(request("txtProductID")) <> 100 and trim(StatusArray(i)) = "Targeted" then
    	    cn.execute "spUpdateSuddenImpactProdDel " & clng(request("txtProductID")) & "," &  clng(trim(VersionArray(i)))
		end if
    next

	set rs=nothing
	set cn=nothing
%>
<%if blnSuccess then%>
<INPUT type="text"  style="display:none" id=txtSuccess name=txtSuccess value="1">
<%else%>
<INPUT type="text"  style="display:none" id=txtSuccess name=txtSuccess value="0">
<%end if%>
</BODY>
</HTML>

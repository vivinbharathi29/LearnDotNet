<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
 <script src="../../Scripts/PulsarPlus.js"></script>
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            //close window
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.closeExternalPopup();
            }
            else if (IsFromPulsarPlus()) {
                window.parent.parent.parent.MultiTestStatusCallback(txtSuccess.value);
                ClosePulsarPlusPopup();
            }
            else {
                window.parent.close(txtTodayPageSection.value, txtRowIDs.value, txtIndex.value);
            }
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to update test statuses.</font>");
    }
    else {
        document.write("<BR><font size=2 face=verdana>Unable to update test statuses.</font>");
    }
}


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<%
	strSuccess = "1"

	dim cn
	dim cm
	dim blnErrors
	dim IDArray
	dim ItemID
	dim RowIDs
    
    IDArray = split(request("lstID"),",")
          
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
    
	cn.BeginTrans
	blnErrors = false

	for each ItemID in IDArray
        dim sp, productid, versionid, productdeliverableid, productdeliverablereleaseid

        if InStr(ItemID,"_") > 0 then
            arrIDs = split(ItemID, "_")
            versionid = arrIDs(0)
            productid = arrIDs(1)
            productdeliverableid = arrIDs(2)
            productdeliverablereleaseid = arrIDs(3)

            if productdeliverablereleaseid > 0 then
                sp = "spUpdateTestLeadStatusOnlyPulsar"
            else
                sp = "spUpdateTestLeadStatusOnly"
            end if

            if request("txtTodayPageSection") <> "" then
                if RowIDs <> "" then
                    RowIDs = RowIDs & ","
                end if

                RowIDs = RowIDs & productdeliverableid & "_" & productdeliverablereleaseid
            else 
                if RowIDs <> "" then
                    RowIDs = RowIDs & ","
                end if

                RowIDs = RowIDs & versionid
            end if
        end if
        
		if isnumeric(versionid) and trim(versionid) <> "" then
			
			if trim(request("cboIntegrationStatus")) <> "" and trim(request("cboIntegrationStatus")) <> "0" then
				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
					
				cm.CommandText = sp	

                if productdeliverablereleaseid > 0 then
				    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
				    p.Value = clng(productdeliverablereleaseid)
				    cm.Parameters.Append p
                end if 

				Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
				p.Value = clng(productid)
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@DeliverableID", 3,  &H0001)
				p.Value = clng(versionid)
				cm.Parameters.Append p
                
				Set p = cm.CreateParameter("@FieldID", 3,  &H0001)
				p.Value = 1
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@UserID", 3,  &H0001)
				p.Value = clng(CurrentUserID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Username", 200,  &H0001,80)
				p.Value = left(CurrentDomain + "_" + Currentuser,80)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
				p.Value = clng(request("cboIntegrationStatus"))
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Notes", 200,  &H0001,200)
				p.Value = left(request("txtIntegrationNotes"),200)
				cm.Parameters.Append p                
                
				cm.Execute rowschanged

				set cm=nothing
					
				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if

                
			end if 'End Integration
            
	        
			if strSuccess <> "0" and trim(request("cboODMStatus")) <> "" and trim(request("cboODMStatus")) <> "0" then
				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
					
				cm.CommandText = sp	

				if productdeliverablereleaseid > 0 then
				    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
				    p.Value = clng(productdeliverablereleaseid)
				    cm.Parameters.Append p
                end if 

				Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
				p.Value = clng(productid)
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@DeliverableID", 3,  &H0001)
				p.Value = clng(versionid)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@FieldID", 3,  &H0001)
				p.Value = 2
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@UserID", 3,  &H0001)
				p.Value = clng(CurrentUserID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Username", 200,  &H0001,80)
				p.Value = left(CurrentDomain + "_" + Currentuser,80)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
				p.Value = clng(request("cboODMStatus"))
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Notes", 200,  &H0001,200)
				p.Value = left(request("txtODMNotes"),200)
				cm.Parameters.Append p

				cm.Execute rowschanged

				set cm=nothing

					
				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if
			end if 'End ODM

			if strSuccess <> "0" and trim(request("cboWWANStatus")) <> "" and trim(request("cboWWANStatus")) <> "0" then
				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
					
				cm.CommandText = sp	

				if productdeliverablereleaseid > 0 then
				    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
				    p.Value = clng(productdeliverablereleaseid)
				    cm.Parameters.Append p
                end if 

				Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
				p.Value = clng(productid)
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@DeliverableID", 3,  &H0001)
				p.Value = clng(versionid)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@FieldID", 3,  &H0001)
				p.Value = 3
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@UserID", 3,  &H0001)
				p.Value = clng(CurrentUserID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Username", 200,  &H0001,80)
				p.Value = left(CurrentDomain + "_" + Currentuser,80)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
				p.Value = clng(request("cboWWANStatus"))
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Notes", 200,  &H0001,200)
				p.Value = left(request("txtWWANNotes"),200)
				cm.Parameters.Append p

				cm.Execute rowschanged

				set cm=nothing

					
				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if
			end if 'End WWAN
			
    		if strSuccess <> "0" and trim(request("cboDEVStatus")) <> "" and trim(request("cboDEVStatus")) <> "0" then
				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
					
				cm.CommandText = sp	

				if productdeliverablereleaseid > 0 then
				    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
				    p.Value = clng(productdeliverablereleaseid)
				    cm.Parameters.Append p
                end if 

				Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
				p.Value = clng(productid)
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@DeliverableID", 3,  &H0001)
				p.Value = clng(versionid)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@FieldID", 3,  &H0001)
				p.Value = 4
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@UserID", 3,  &H0001)
				p.Value = clng(CurrentUserID)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Username", 200,  &H0001,80)
				p.Value = left(CurrentDomain + "_" + Currentuser,80)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
				p.Value = clng(request("cboDEVStatus"))
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Notes", 200,  &H0001,200)
				p.Value = left(request("txtDEVNotes"),200)
				cm.Parameters.Append p

				cm.Execute rowschanged

				set cm=nothing

					
				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if
			end if 'End DEV
			
		end if
	next
	if strSuccess="1" then
		cn.CommitTrans
	else
		cn.RollbackTrans
	end if
			
	cn.Close
	set rs = nothing
	set cn = nothing
	
	
	
	
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=Request("app")%>">
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=request("txtTodayPageSection")%>" />
<input type="hidden" id="txtRowIDs" name="txtRowIDs" value="<%=RowIDs%>" />
<input type="hidden" id="txtIndex" name="txtIndex" value="<%=request("txtIndex")%>" />
</BODY>
</HTML>

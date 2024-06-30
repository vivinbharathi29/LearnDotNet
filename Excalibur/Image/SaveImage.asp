<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (txtSuccess.value == "1") {
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.closeExternalPopup();
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);                
           }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
                /*window.returnValue = 1;
                window.parent.close();*/
            }
        }
        else
            document.write("Unable to save Image Definition.");
		
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">


<%


	dim PriorityArray
	dim TagArray
	Dim ImageIDArray
	Dim ImageNameArray
	dim i
	dim cn
	dim cm
	dim p
	dim blnRegionsChanged
	dim blnDefinitionChanged
	dim rowschanged
	dim FoundErrors
	dim NewID
	dim strLog
	dim strChangeLog
	dim strRegionChangeLog
	dim strChangeMail
	dim strRegionChangeMail
	dim strSKU
	dim oMessage
	dim strBody
	dim strProductName
	dim strPreinstallTeam
	dim rs
	dim rs2
	dim strDelList
	dim strExecptions
	dim strImageID
	dim strProductEmail
	dim strImageDefinitionId
	
	strSKU = request("txtSku")
	if strSKU <> "" and request("txtSKUDigit") <> "" then
		strSKU = strSKU & "-xx" & request("txtSKUDigit")
	end if
	
	blnRegionsChanged = request("txtTag") <> request("cboPriority")

	blnDefinitionChanged =  request("tagImageType") <> request("cboImageType") or request("tagComments") <> request("txtComments")  or request("tagRTMDate") <> request("txtRTMDate") or request("tagOS") <>  request("cboOS") or request("cboType") <> request("tagType") or request("tagSW") <> request("cboSW") or request("tagSku") <> strSKU or request("cboBrand") <> request("tagBrand") or request("cboStatus") <> request("tagStatus") or request("cboDriveDefinition") <> request("tagDriveDefinition")

	NewID = trim(request("txtDisplayedID"))


	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	cn.BeginTrans

	FoundErrors = false	

	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")	
	
	cm.CommandText = "spGetProductVersionname"	

	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = request("txtProdID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	if not(rs.EOF and rs.BOF) then
		strProductName = rs("Name") & ""
		strPreinstallTeam = rs("preinstallteam") & ""
		strProductEmail = rs("Email") & ""
	end if
	rs.Close
	set rs=nothing

	strLog = ""
	strChangeLog = ""
	strChangeMail = ""	

	if  blnDefinitionChanged then
	
		if request("tagSku") <> strSku then
			strLog = strLog & "<TR><TD>SKU Number</TD>"
			strLog = strLog &  "<TD>" & request("tagSku") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & strSku & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "SKU Number Changed from " & request("tagSku") & " to " & strSku & "; "			
		end if
		
		if trim(request("txtBrandTag")) <>  trim(request("txtBrandText")) then
			strLog = strLog & "<TR><TD>Brand</TD>"
			strLog = strLog &  "<TD>" & request("txtBrandTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtBrandText") & "&nbsp;</TD></TR>"	
			strChangeLog = strChangeLog & "Brand Changed from " & request("txtBrandTag") & " to " & request("txtBrandText") & "; "			
		end if

		if request("tagOS") <>  request("cboOS") then
			strLog = strLog & "<TR><TD>Operating System</TD>"
			strLog = strLog &  "<TD>" & request("txtOSTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtOSText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "OS Changed from " & request("txtOSTag") & " to " & request("txtOSText") & "; "
		end if

		if request("tagSW") <>  request("cboSW") then
			strLog = strLog & "<TR><TD>Software</TD>"
			strLog = strLog &  "<TD>" & request("txtSWTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtSWText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "SW Changed from " & request("txtSWTag") & " to " & request("txtSWText") & "; "
		end if

		if request("tagType") <>  request("cboType") then
			strLog = strLog & "<TR><TD>Type</TD>"
			strLog = strLog &  "<TD>" & request("txtTypeTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtTypeText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Type changed from " & request("txtTypeTag") & " to " & request("txtTypeText") & "; "
		end if

		if request("tagStatus") <>  request("cboStatus") then
			strLog = strLog & "<TR><TD>Status</TD>"
			strLog = strLog &  "<TD>" & request("txtStatusTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtStatusText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Status Changed from " & request("txtStatusTag") & " to " & request("txtStatusText") & "; "
		end if

		if request("tagRTMDate") <>  request("txtRTMDate") then
			strLog = strLog & "<TR><TD>RTM Date</TD>"
			strLog = strLog &  "<TD>" & request("tagRTMDate") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtRTMDate") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "RTM Date Changed from " & request("txtRTMDate") & " to " & request("txtRTMDate") & "; "
		end if
		
		if request("tagDriveDefinition") <> request("cboDriveDefinition") then
			strLog = strLog & "<TR><TD>Master SKU Component</TD>"
			strLog = strLog &  "<TD>" & request("tagDriveDefinition") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("cboDriveDefinition") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Master Sku Component Changed from " & request("tagDriveDefinition") & " to " & request("cboDriveDefinition") & "; "
		end if

		if request("tagComments") <>  request("txtComments") then
			strLog = strLog & "<TR><TD>Comments</TD>"
			strLog = strLog &  "<TD>" & request("tagComments") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtComments") & "&nbsp;</TD></TR>"			
		end if

		strChangeMail = strProductName & ", " & strSku & ", Image Definition has changed: "
		strChangeMail = strChangeMail & strChangelog
		
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		if trim(request("txtDisplayedID")) = "" then
			Response.Write "Adding<BR>"
			cm.CommandText = "spAddImageDefinition"	

			Set p = cm.CreateParameter("@ProdID", 3,  &H0001)
			p.Value = request("txtProdID")
			cm.Parameters.Append p
		else
			Response.Write "Updating<BR>"
			cm.CommandText = "spUpdateImageDefinition"	
			
			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = request("txtDisplayedID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@ImageChange", 200, &H0001, 500)
			p.Value = left(strChangeLog, 500)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@RegionChange", 200, &H0001, 500)
			p.Value = ""
			cm.Parameters.Append p

		end if
		
		Set p = cm.CreateParameter("@SKU", 200, &H0001, 20)
		p.value = left(strSku,20)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@BrandID", 3, &H0001)
		p.Value = request("cboBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSID", 3, &H0001)
		p.Value = request("cboOS")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SWID", 3, &H0001)
		p.Value = request("cboSW")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ImageType", 200, &H0001, 10)
		p.value = left(request("cboType"),10)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ImageTypeID", 3, &H0001)
		p.value = clng(request("cboImageType"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@StatusID", 16, &H0001)
		p.Value = request("cboStatus")
		cm.Parameters.Append p

'		Set p = cm.CreateParameter("@Active", 3, &H0001)
'		if request("cboStatus") = 3 then
'			p.Value = 0
'		else
'			p.Value = 1
'		end if
'		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RTMDate", 135, &H0001)
		if trim(request("txtRTMDate")) = "" then
			p.Value = null
		else
			p.Value = request("txtRTMDate")
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Comments", 200, &H0001, 50)
		p.value = left(request("txtComments"),50)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@ImageDriveDefinitionId", 3,  &H0001)
		If Request("cboDriveDefinition") = "" Then
            p.Value = NULL
    	Else
		    p.Value = request("cboDriveDefinition")
		End If
		cm.Parameters.Append p

		if trim(request("txtDisplayedID")) = "" then
			Set p = cm.CreateParameter("@NewID", 3,  &H0002)
			cm.Parameters.Append p
		end if
		


		cm.Execute rowschanged

		if rowschanged <> 1 then
			FoundErrors = true
		else
			if trim(request("txtDisplayedID")) = "" then
				NewID = cm("@NewID")
			end if
		end if
		
		set cm = nothing
		
	end if
	Response.Write "New ID" & NewID & "<BR>"
    response.write "Tags:" & request("txtTag") & "<BR><BR>"
    response.write "ImageID:" & request("txtimageIDList") & "<BR><BR>"
	strRegionChangeLog = ""

	if blnRegionsChanged then
		PriorityArray = split(request("cboPriority"),",")
		TagArray = split(request("txtTag"),",")
		ImageIDArray = split(request("txtimageIDList"),",")
		ImageNameArray = split(request("txtimageNameList"),",")
	
		if not( ubound(PriorityArray) = ubound(TagArray) and  ubound(PriorityArray) = ubound(ImageIDArray)) then
			Response.Write "<BR><BR>Invalid Array Lengths<BR>"
			Response.Write "PriorityArray: " & ubound(PriorityArray) & "<BR>"
			Response.Write "TagArray: " & ubound(TagArray) & "<BR>"
			Response.Write "ImageIDArray: " & ubound(ImageIDArray) & "<BR>"
			FoundErrors = true
		else
			
			for i = lbound(PriorityArray) to ubound(PriorityArray)
			
				if trim(PriorityArray(i)) <> trim(TagArray(i)) then
					set cm = server.CreateObject("ADODB.Command")
					cm.ActiveConnection = cn
					cm.CommandType =  &H0004

					if trim(PriorityArray(i)) = "" then
						cm.CommandText = "spRemoveRegionFromImage"

						Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
						p.Value = NewID
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
						p.Value = ImageIDArray(i)
						cm.Parameters.Append p

						strRegionChangeLog = strRegionChangeLog & "Removed Region " & ImageNameArray(i) & "; "

					elseif trim(TagArray(i)) = "" then
						cm.CommandText = "spAddRegion2Image"

						Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
						p.Value = NewID
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
						p.Value = ImageIDArray(i)
						cm.Parameters.Append p
						
						Set p = cm.CreateParameter("@Priority", 200,  &H0001,6)
						p.Value = left(PriorityArray(i),6)
						cm.Parameters.Append p

						if len(trim(left(PriorityArray(i),6))) = 1 then
							strRegionChangeLog = strRegionChangeLog & "Added Region " & ImageNameArray(i) & " to Tier " & PriorityArray(i) & "; "
						else
							strRegionChangeLog = strRegionChangeLog & "Added Region " & ImageNameArray(i) & " to use dash " & PriorityArray(i) & "; "
						end if
												
					else
						cm.CommandText = "spUpdateRegion4Image"

						Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
						p.Value = NewID
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
						p.Value = ImageIDArray(i)
						cm.Parameters.Append p
						
						Set p = cm.CreateParameter("@Priority", 200,  &H0001,6)
						p.Value = left(PriorityArray(i),6)
						cm.Parameters.Append p

						if len(trim(left(PriorityArray(i),6))) = 1 then
							strRegionChangeLog = strRegionChangeLog & "Updated Region " & ImageNameArray(i) & " to Tier " & trim(left(PriorityArray(i),6)) & "; "
						else
							strRegionChangeLog = strRegionChangeLog & "Updated Region " & ImageNameArray(i) & " to use dash " & trim(left(PriorityArray(i),6)) & "; "
						end if

					end if
			
					cm.Execute rowschanged
	
					if rowschanged <> 1 then
						FoundErrors = true
						exit for
					end if
		
					set cm = nothing
			
				end if
			next			


			for i = lbound(PriorityArray) to ubound(PriorityArray)
				if trim(TagArray(i)) <> trim(PriorityArray(i)) then
						strLog = strLog &  "<TR>"
						strLog = strLog & "<TD>" & ImageNameArray(i) & "</TD>"
						strLog = strLog &  "<TD>" & TagArray(i) & "&nbsp;</TD>"
						strLog = strLog & "<TD>" & PriorityArray(i) & "&nbsp;</TD>"			
						strLog = strLog & "</TR>"
				end if
			next
		
		end if
	end if
	
	
	'Snapshot Deliverable List if releasing to factory	
	if trim(request("tagStatus")) = "1" and trim(request("cboStatus")) = "2" then
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")


		rs.Open "Select * from images with (NOLOCK) where ImagedefinitionID = " &  request("txtDisplayedID"),cn,adOpenStatic
		do while not rs.EOF
            if trim(rs("LockedDeliverableList") & "") = "" then
			    rs2.Open "spListDeliverablesInImage " & rs("ID") & ",0" ,cn,adOpenStatic 
			    strDelList = ""
			    strExecptions = ""
			    do while not rs2.EOF
				    strImageID = rs("ID")
				    if ( rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore") ) and rs2("InImage") and ( trim(rs2("Images") & "") = "" or instr(", " & rs2("Images") & ",", ", " & strImageID & ",")>0  or instr( rs2("Images") , "(" & strImageID & "=")>0 )  then
					    strDelList = strDelList & ", " & rs2("ID")
					    if instr(rs2("Images"),"(" & trim(rs("ID")) & "=") > 0 then
						    strExecptions = strExecptions & ";(" & rs2("ID") & "=" & GetExceptions(rs2("Images") & "", rs("ID") & "") & ")"
					    end if
				    end if
				    rs2.MoveNext
			    loop
			    rs2.Close
			    if strDelList <> "" then
				    strDelList = mid(strDelList,3) & ","
			    end if
			    if strExecptions <> "" then
				    strExecptions = mid(strExecptions,2)
				    strDelList = strDelList & ":" & strExecptions
			    end if
		
			    set cm = server.CreateObject("ADODB.Command")
			    cm.CommandType =  &H0004
			    cm.ActiveConnection = cn
		
			    cm.CommandText = "spUpdateImageLockedDeliverableList"	

			    Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			    p.Value = rs("ID")
			    cm.Parameters.Append p

			    Set p = cm.CreateParameter("@DeliverableList", 201, &H0001, 2147483647)
			    p.value = strDelList
			    cm.Parameters.Append p

			    cm.Execute rowschanged

			    set cm = nothing		
			
			
			    if rowschanged <> 1 then
				    FoundErrors = true
				    exit do
			    end if
            end if			
			rs.MoveNext
		loop
		rs.Close

		set rs2 = nothing
		set rs = nothing	
	end if
	
	
	
	strRegionChangeMail = strProductName & ", " & strSku & ", Region has changed: "
	strRegionChangeMail = strRegionChangeMail & strRegionChangelog
	

	Response.Write "Updating Log<BR>"

	if (blnRegionsChanged or blnDefinitionChanged) and (not FoundErrors)  then
		
		if trim(request("txtDisplayedID")) <> "" then
			cn.Execute "spImageModified " & NewID,rowschanged
			if rowschanged <> 1 then
				FoundErrors = true
			end if
		end if
				
		if (not founderrors) and trim(request("cboDCR")) <> "" then
			set cm = server.CreateObject("ADODB.Command")
			cm.ActiveConnection = cn
			cm.CommandType =  &H0004
			cm.CommandText = "spAddIMageLog"

			Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
			p.Value = clng(request("txtCurrentUserID"))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@DCRID", 3,  &H0001)
			if request("cboDCR") <> "" and isnumeric(request("cboDCR")) then
				p.Value = clng(request("cboDCR"))
			else
				p.Value = null
			end if
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = NewID
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@Details", 200,  &H0001,7500)
			if trim(request("txtDisplayedID")) = "" then
				p.Value = "Added"
			else
				p.Value = left(strLog,7500)
			end if
			cm.Parameters.Append p
		
			cm.Execute rowschanged
	
			if rowschanged <> 1 then
				FoundErrors = true
			end if
		
		end if
	end if


	if FoundErrors then
		Response.Write "Rolling Back<BR>"

		cn.RollbackTrans
		%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="0"><%
	else
		cn.CommitTrans
		Response.Write "Save Complete<BR>"
		%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="1"><%
		
		'update Region Changes
		if 	strRegionChangeLog <> "" and request("txtDisplayedID") <> "" then
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
	
			cm.CommandText = "spUpdateImageDefinitionRegionChange"	

			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = request("txtDisplayedID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@RegionChange", 200, &H0001, 500)
			p.Value = left(strRegionChangeLog, 500)
			cm.Parameters.Append p

			cm.Execute 
			Set cm=nothing
		end if

		If request("txtProdID") <> 100 AND strProductEmail = "1" Then
		    if strChangeLog <> "" or strRegionChangeLog <> "" then
			    'send Email to Preinstall DB team when there is a definition or region change
			    Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			    'Set oMessage.Configuration = Application("CDO_Config")		

			    if request("txtCurrentUserEmail") = "" then
				    oMessage.From = "max.yu@hp.com"
			    else
				    oMessage.From = request("txtCurrentUserEmail")
			    end if

			    if strPreinstallTeam = 1 then
				    oMessage.To = "houportpreindb@hp.com" 
			    elseif strPreinstallTeam = 2 then
				    oMessage.To = "TWN.PDC.NB-preinPM@hp.com"
			    else
				    oMessage.To = "max.yu@hp.com"
			    end if 
    		
			    oMessage.Subject = "Image Change Notification"

			    strBody = "<b><font size=3 face=verdana>" & "Image and Region Changes" & "</font></b>"
			    strBody = strBody & "<br><br>"
			    strBody = strBody & strChangeMail & "<br><br>"
			    strBody = strBody & strRegionChangeMail & "<br>"
    		
			    oMessage.HTMLBody = "<font size=2 face=verdana><Font color=red><b>NOTE: This email was generated by a change to the <u>definiton</u> of the images defined for the product.</b></font><BR><BR>" & strBody & "</font>"
    						
			    oMessage.Send 
			    Set oMessage = Nothing 		
		    end if
		End If

	end if
	
	cn.Close
	set cm = nothing
	set cn = nothing
	
	
	
Function GetExceptions(strList, strItem)
	dim strTemp
	strTemp = mid(strList, instr(strList,"(" & strItem & "=") + len(strItem) + 2) 
	strTemp = left(strTemp,instr(strTemp,")")-1)
	GetExceptions = strTemp
end function	
%>
</BODY>
</HTML>

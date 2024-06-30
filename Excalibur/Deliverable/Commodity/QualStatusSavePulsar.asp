<%@ Language=VBScript %>

<!-- #include file = "../../includes/noaccess.inc" -->
<!-- #include file="../../includes/EmailQueue.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--
    $(document).ready(function () {
        if ($("#txtSuccess")) {
            if ($("#txtSuccess").val() != "0" && $("#txtKeepItOpen").val() == "false")
            {
                if (parent.window.parent.loadDatatodiv != undefined) {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp('Deliverables');
                }
                else if (IsFromPulsarPlus()) {
                    window.parent.parent.parent.EditCommodityStatus2ReloadCallback($("#txtSuccess").val());
                    ClosePulsarPlusPopup();
                } else {
                    window.parent.closewindow($("#txtVersionID").val(), $("#txtProductDeliverableID").val(), $("#txtProdDelRelID").val(), $("#txtSuccess").val(), $("#txtTodayPageSection").val(), $("#txtProductID").val());
                }
            }
            else {
                window.parent.SetNewStatus($("#txtSuccess").val());
                document.location = txtRedirect.value;
                window.parent.repositionParentWindow();                
            }
        }
    });

//-->
</SCRIPT>

</HEAD>
<BODY>

<%
    response.flush
	dim cn
	dim cm
	dim rs
	dim strSuccess
	dim strSubsAdded
	dim strSubsRemoved
	dim i
	dim SubsLoadedArray
	dim SubsSelectedArray
	dim SubsAddedArray
	dim SubsRemovedArray
	dim strElement
	dim blnSupplyRestrictionChanged
	dim blnConfigurationRestrictionChanged
	dim strBody
	dim strQCompleteListHP
	dim strQCompleteListODM
	dim strFailedListODM
	dim strBridgeBody
	dim strExtraNote
    dim strReleaseName
    dim strRedirect
    dim blnKeepItOpen

    blnKeepItOpen = false
    strRedirect = ""
	strBridgeBody = ""

    blnKeepItOpen = request("txtKeepItOpen")
    strRedirect = request("txtRedirect")

	strFailedListODM = ""
	strQCompleteListHP = "TWNPDCNBCommodityTechnology@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;APJ-RCTO.SC@hp.com;claire.lin@hp.com"
	Select case trim(request("txtPartnerID"))
	case "2"
		strQCompleteListODM = ";IPC-ED1@inventec.com;IPCHP-Excalibur@inventec.com"
		strFailedListODM = ";IPC-ED1@inventec.com;IPCHP-Excalibur@inventec.com"
	case "7"
		strQCompleteListODM = ""
		strFailedListODM = ""
	case "3"
		strQCompleteListODM = ";A32KeyCommodity@compal.com"
		strFailedListODM = ";A32KeyCommodity@compal.com"
	case "4"
		strQCompleteListODM = ";TOPCommodityTeam@quantacn.com"
		strFailedListODM = ";TOPCommodityTeam@quantacn.com"
	case else
		strQCompleteListODM = ""
		strFailedListODM = ""
	end select

	blnSupplyRestrictionChanged = false
	blnConfigurationRestrictionChanged = false

'	dim strBodyComments

	if trim(request("txtComments")) = "" then
		strBodyComments = "No Comments Entered<BR>" 
	else
		strBodyComments = left(request("txtComments"),255) & "<BR>"
	end if
	
	SubsLoadedArray = split(request("txtSubsLoaded"),",")
	SubsSelectedArray = split(request("lstSub"),",")
	
	strSubsAdded = ""
	strSubsRemoved = ""
	
	'Look for Subs Added
	for each strElement in SubsSelectedArray
		if instr(", " & request("txtSubsLoaded") & ",",", " & trim(strElement) & ",") = 0 then
			strSubsAdded = strSubsAdded & "," & strElement
		end if
	next 

	'Look for Subs Removed
	for each strElement in SubsLoadedArray
		if instr(", " & request("lstSub") & ",",", " & trim(strElement) & ",") = 0 then
			strSubsRemoved = strSubsRemoved & "," & strElement
		end if
	next 
		
	if strSubsAdded <> "" then
		strSubsAdded = mid(strSubsAdded,2)
	end if	

	if strSubsRemoved <> "" then
		strSubsRemoved = mid(strSubsRemoved,2)
	end if	

	SubsAddedArray = split(strSubsAdded,",")
	SubsRemovedArray = split(strSubsRemoved,",")

		
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	cn.BeginTrans
	if trim(request("cboStatus")) = "0" then
		strSuccess = "&nbsp;"
	elseif request("chkSupplyRestricted") = "1" or request("chkConfigurationRestricted") = "1" then
		strSuccess = "RSTD:" & request("txtStatusText")
	elseif request("chkSupplyRestricted") = "" and request("chkConfigurationRestricted") = "" then
		strSuccess = "UNRS:" & request("txtStatusText")
	else
		strSuccess = request("txtStatusText")
	end if

	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
	
	if trim(request("cboStatus")) = "0" and request("chkDelete") = "on" then
		cm.CommandText = "spUnLinkVersionFromProductRelease"	
		
		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = clng(request("txtProdID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
		p.Value = clng(request("txtVersionID"))
		cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3,  &H0001)
		p.Value = clng(request("txtProdDelRelID"))
		cm.Parameters.Append p
	else
		cm.CommandText = "spUpdateCommodityStatusPulsar"	
		
		Set p = cm.CreateParameter("@ID", 3,  &H0001)
		p.Value = clng(request("txtID"))
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Status", 3,  &H0001)
		p.Value = clng(request("cboStatus"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RiskRelease", 16,  &H0001)
		if request("chkRiskRelease") = "on" and clng(request("cboStatus")) = 5 then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@UserID", 3,  &H0001)
		p.Value = clng(request("txtUserID"))
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@UserName", 200,  &H0001, 80)
		p.Value = left(request("txtUserName"),80)
		cm.Parameters.Append p
		
	
		Set p = cm.CreateParameter("@DCRID", 3,  &H0001)
		if clng(request("cboWhy")) < 3 then
			p.Value = clng(request("cboWhy"))
		else
			p.Value = clng(request("cboDCR"))
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SupplyChainRestriction", 16,  &H0001)
		if trim(request("chkSupplyRestricted")) = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ConfigurationRestriction", 16,  &H0001)
		if trim(request("chkConfigurationRestricted")) = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p		
		
		Set p = cm.CreateParameter("@Comments", 200,  &H0001,255)
		p.Value = left(request("txtComments"),255)
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@TestDate", 135,  &H0001)
		if isdate(request("txtTestDate")) then
			p.Value = cdate(request("txtTestDate"))
		else
			p.value = null
		end if
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@TestConfidence", 16,  &H0001)
		p.Value = cint(request("cboConfidence"))
		cm.Parameters.Append p

        Set p = cm.CreateParameter("@BatchMode",3, &H0001)
		p.Value = 0
		cm.Parameters.Append p
        
		if trim(request("tagSupplyRestriction")) <> trim(request("chkSupplyRestricted")) then
			blnSupplyRestrictionChanged = true
		end if

		if trim(request("tagConfigurationRestriction")) <> trim(request("chkConfigurationRestricted")) then
			blnConfigurationRestrictionChanged = true
		end if
	
	end if

	cm.Execute rowschanged
	
	set cm=nothing

	if rowschanged <> 1 then
		strSuccess = "0"
	end if	

	if strSuccess <> "0" and not (trim(request("cboStatus")) = "0" and request("chkDelete") = "on")then
		'Add Subassemblies	
		for each strElement in SubsAddedArray
			if isnumeric(strElement) and trim(strElement) <> "" then
				cn.Execute "spAddSubassemblyLinkRelease " & clng(request("txtID")) & "," & clng(strElement), rowschanged
				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if 
				'See if this is bridged and generate the notification row
				rs.open "spGetSubassemblyBridgeRelease " & clng(request("txtID")) & "," & clng(strElement),cn
				if not (rs.eof and rs.bof) then
					if rs("Bridged") then
						strBridgeBody = strBridgeBody & "<TR><TD>Bridge Added</TD><TD>" & rs("Product") & "</TD><TD>" & rs("deliverablename") & "</TD><TD>" & rs("partnumber") & "&nbsp;</TD><TD>" & rs("subassemblyname") & "</TD><TD>" & rs("subassembly") & "&nbsp;</TD></TR>"
					end if
				end if
				rs.close
			end if
		next

		'Remove Subassemblies
		for each strElement in SubsRemovedArray
			if isnumeric(strElement) and trim(strElement) <> "" then

				'See if this is bridged and generate the notification row
				rs.open "spGetSubassemblyBridgeRelease " & clng(request("txtID")) & "," & clng(strElement),cn
				if not (rs.eof and rs.bof) then
					if rs("Bridged") then
						strBridgeBody = strBridgeBody & "<TR><TD>Bridge Removed</TD><TD>" & rs("Product") & "</TD><TD>" & rs("deliverablename") & "</TD><TD>" & rs("partnumber") & "&nbsp;</TD><TD>" & rs("subassemblyname") & "</TD><TD>" & rs("subassembly") & "&nbsp;</TD></TR>"
					end if
				end if
				rs.close


				cn.Execute "spRemoveSubassemblyLinkRelease " & clng(request("txtID")) & "," & clng(strElement),rowschanged

				'Response.write clng(request("txtID")) & "," & clng(strElement) & "<BR>"

				if rowschanged <> 1 then
					strSuccess = "0"
					exit for
				end if 
			end if
		next
	
	end if

        
	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
        if request("txtTodayPageSection") = "" then 
            rs.open "usp_ProductDeliverable_GetQualStatus "	& clng(request("txtVersionID")) & "," & clng(request("txtProdID")),cn
	        if not (rs.EOF and rs.BOF) then
	            strSuccess = clng(request("txtVersionID")) & "|" & rs("Targeted") & "|" & rs("RlsRestrict") & "|" & rs("RlsQStatus")
	        end if
	        rs.Close
        end if
	end if

	
	if strBridgeBody <> "" and strSuccess <> "0" then
		rs.Open "spListEngineeringCoorginators",cn
		dim strTMPEmail
		strTMPEmail = ""
		do while not rs.EOF
			strTMPEmail = strTMPEmail & ";" & rs("Email")
			rs.MoveNext
		loop
		if strTMPEmail = "" then
			strTMPEmail = "max.yu@hp.com"
		else
			strTMPEmail= mid(strTMPEmail,2) & ";tdcesmail@hp.com"
		end if
		rs.Close

		Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		oMessage.To = strTMPEmail & ";" & request("txtUserEmail")
		oMessage.Subject = "Subassembly Bridge Updates"
		oMessage.HTMLBody = "<STYLE>TD{FONT-FAMILY: Verdana;FONT-SIZE:xx-small;}</STYLE><font size=2 face=verdana>The following subassembly actions have beed requested:</font><Table bgcolor=ivory cellpadding=2 cellspacing=0 border=1><TR bgcolor=beige><TD><b>Action</b></TD><TD><b>Product</b></TD><TD><b>Deliverable</b></TD><TD><b>Part&nbsp;Number</b></TD><TD><b>Subassembly&nbsp;Name</b></TD><TD><b>Subassembly&nbsp;Number</b></TD></TR>" & strBridgeBody & "</TABLE>"
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	
	end if

	if strSuccess <> "0" and ( (trim(request("txtStatusLoaded")) <> "6" and trim(request("cboStatus")) = "6") or (trim(request("txtStatusLoaded")) <> "18" and trim(request("cboStatus")) = "18") or (trim(request("txtStatusLoaded")) <> "7" and trim(request("cboStatus")) = "7") or (trim(request("txtStatusLoaded")) <> "10" and trim(request("cboStatus")) = "10") ) then
		Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")	
        dim strHDRecipients
		oMessage.From = request("txtUserEmail")
		if clng(request("txtProdID")) = 100 then
			oMessage.To = "max.yu@hp.com"
		else
			if request("txtDevCenter") = "2" then
				strHDRecipients = "TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;tdcesmail@hp.com;twnpdccnbcommoditypm@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com" & ";" & Request("txtPMsEmail") 
            elseif request("txtDevCenter") = "3"  then
                strHDRecipients = "TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com;tdcesmail@hp.com;twinkle.k.s@hp.com;sridevi.s@hp.com;rajendran.m@hp.com" & strFailedListODM & ";" & Request("txtPMsEmail") 
			else 
				strHDRecipients = "TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com" & strFailedListODM & ";" & Request("txtPMsEmail") 
			end if
		end if
		if trim(request("cboStatus")) = "6" then
			oMessage.Subject = "Hardware Drop Notification"
            strHDRecipients = strHDRecipients + ";tdcdtoedmfunction@hp.com"
		elseif trim(request("cboStatus")) = "7" then
			oMessage.Subject = "Hardware Hold Notification"
		elseif trim(request("cboStatus")) = "10" then
			oMessage.Subject = "Hardware Failure Notification"
		elseif trim(request("cboStatus")) = "18" then
			oMessage.Subject = "Hardware Status Changed to 'Service Only'"
		else
			oMessage.Subject = "Hardware Status Updated"
		end if
		oMessage.Importance = cdoHigh
		if trim(request("cboStatus")) = "18" then
		    strExtraNote = "<BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service."
		else
		    strExtraNote = ""
        end if
        oMessage.To = strHDRecipients
        oMessage.CC = "isc.sj@hp.com; "
		oMessage.HTMLBody = "<font size=2 face=verdana color=black>" & replace(replace(replace(request("txtFailBody"),"##ShowComments.##",strBodyComments),"##StatusText.##",request("txtStatusText")),"##ExtraNote.##",strExtraNote) & "</font>"
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	end if

	if  strSuccess <> "0" and (trim(request("txtStatusLoaded")) <> "5" and trim(request("cboStatus")) = "5") or (trim(request("cboStatus")) = "5" and request("tagRiskRelease") <> request("chkRiskRelease")  ) then
		Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		if clng(request("txtProdID")) = 100 then
			oMessage.To = "max.yu@hp.com"
		elseif request("txtDevCenter") = "2" then
			oMessage.To = Request("txtDevEmail") & ";" & strQCompleteListHP & strQCompleteListODM & ";tdcesmail@hp.com;twnpdccnbcommoditypm@hp.com" 
        elseif request("txtDevCenter") = "3" then
            oMessage.To = Request("txtDevEmail") & ";" & strQCompleteListHP & strQCompleteListODM & ";tdcesmail@hp.com;twinkle.k.s@hp.com;sridevi.s@hp.com;rajendran.m@hp.com"
		else
			oMessage.To = Request("txtDevEmail") & ";" & strQCompleteListHP & strQCompleteListODM
		end if
		if  request("chkRiskRelease") = "on" then
			oMessage.Subject = replace(request("txtQCompleteSubject"),"QComplete","Risk Release")
		else
			oMessage.Subject = request("txtQCompleteSubject")
		end if
		if  request("chkRiskRelease") = "on" then
			oMessage.HTMLBody = replace("<font size=2 face=verdana color=black>" & replace(request("txtQCompleteBody"),"##ShowComments.##",strBodyComments) & "</font>","QComplete","Risk Release")
		else
			oMessage.HTMLBody = "<font size=2 face=verdana color=black>" & replace(request("txtQCompleteBody"),"##ShowComments.##",strBodyComments) & "</font>"
		end if
        oMessage.CC = "isc.sj@hp.com; "
	    oMessage.DSNOptions = cdoDSNFailure
	    oMessage.Send 
	    Set oMessage = Nothing
	end if



	if blnSupplyRestrictionChanged and strSuccess <> "0" then		
		strBody = ""
		strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1 style=""WIDTH:100%""><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD>" & request("txtRestrictBody") & "</table>"
		strBody = strBody & "<BR><TABLE  style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>Product List</b></TD></TR><TR><TD>" & request("txtProduct") & "</TD></TR></TABLE>"
		strBody = strBody & "<BR><TABLE  style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>Comments</b></TD></TR><TR><TD>" & strBodyComments & "</TD></TR></TABLE>"
		strTo = "TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;kidwell.proceng@hp.com"
		strTo = strTo & ";" & request("txtUserEmail")
		strTo = strTo & ";" & Request("txtPMsEmail")
		
		
		Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		
		if lcase (request("txtUserEmail")) = "max.yu@hp.com"  then
			oMessage.To = "max.yu@hp.com"
		else
			oMessage.To = strTo 
		end if
		
		if request("chkSupplyRestricted") = "1" then
			oMessage.Subject = "Supply Chain Restriction Added"
			oMessage.HTMLBody = "<font size=2 face=verdana color=black><b>The following " & request("txtVendor") & " " & request("txtCategory") & " has been restricted (Supply) on the selected products.<BR><BR>" & strBody & "</font>"
		else
			oMessage.Subject = "Supply Chain Restriction Removed"
			oMessage.HTMLBody = "<font size=2 face=verdana color=black><b>The following " & request("txtVendor") & " " & request("txtCategory") & " is no longer restricted (Supply) on the selected products.<BR><BR>" & strBody &  "</font>"
		end if
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	end if

    if blnConfigurationRestrictionChanged and strSuccess <> "0" then
		
		strBody = ""
		strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1 style=""WIDTH:100%""><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD>" & request("txtRestrictBody") & "</table>"
		strBody = strBody & "<BR><TABLE  style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>Product List</b></TD></TR><TR><TD>" & request("txtProduct") & "</TD></TR></TABLE>"
		strBody = strBody & "<BR><TABLE  style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>Comments</b></TD></TR><TR><TD>" & strBodyComments & "</TD></TR></TABLE>"
        
        strTo = ""
        rs.open "spListSystemTeam " & clng(request("txtProdID")),cn
        do while not rs.EOF
            if trim(rs("Role") & "") = "Configuration Manager"  or trim(rs("Role") & "") = "Program Coordinator"  or trim(rs("Role") & "") = "Supply Chain" then
                strTo = strTo & ";" & rs("Email")
            end if
            rs.MoveNext
        loop
        rs.Close

		strTo = strTo & ";" & request("txtUserEmail")
        if strTo = "" or strTo = ";" then
            stTo = "max.yu@hp.com"
        else
            strTo = mid(strTo,2)
        end if		
		
		Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		
		if lcase (request("txtUserEmail")) = "max.yu@hp.com"  then
			oMessage.To = "max.yu@hp.com"
            strBody = strBody & "<BR><BR>Email Would have been sent to: " & strTo
		else
			oMessage.To = strTo
		end if
		
		if request("chkConfigurationRestricted") = "1" then
			oMessage.Subject = "Configuration Restriction Added"
			oMessage.HTMLBody = "<font size=2 face=verdana color=black><b>The following " & request("txtVendor") & " " & request("txtCategory") & " has been restricted (Configuration) on the selected products.<BR><BR>" & strBody & "</font>"
		else
			oMessage.Subject = "Configuration Restriction Removed"
			oMessage.HTMLBody = "<font size=2 face=verdana color=black><b>The following " & request("txtVendor") & " " & request("txtCategory") & " is no longer restricted (Configuration) on the selected products.<BR><BR>" & strBody &  "</font>"
		end if
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 	
	end if

    cn.Close
	set rs = nothing
	set cn = nothing	

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input type="hidden" id="txtRedirect" name="txtRedirect" value="<%=strRedirect%>" />
<input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="<%=blnKeepItOpen%>" />

<INPUT type="hidden" id="txtVersionID" name="txtVersionID" value="<%=request("txtVersionID")%>">
<INPUT type="hidden" id="txtProductDeliverableID" name="txtProductDeliverableID" value="<%=request("txtProductDeliverableID")%>">
<INPUT type="hidden" id="txtProdDelRelID" name="txtProdDelRelID" value="<%=request("txtProdDelRelID")%>">
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=Request("txtTodayPageSection")%>" />
<input type="hidden" id="txtProductID" name="txtProductID" value="<%=request("txtProdID")%>" />
</BODY>
</HTML>

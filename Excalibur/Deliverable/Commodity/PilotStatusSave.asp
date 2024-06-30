<%@ Language=VBScript %>

<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
    if (typeof (txtSuccess) != "undefined") {
        if ($("#txtSuccess").val() != "0") {
            //close window
            if (parent.window.parent.loadDatatodiv != undefined) {
                parent.window.parent.closeExternalPopup();
                parent.window.parent.reloadFromPopUp('Deliverables');
            }
            else
             window.parent.Close($("#txtVersionID").val(), $("#txDeliverableVersionID").val(), $("#txtSuccess").val(), $("#txtTodayPageSection").val());
        }
    }
}



//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">

<%

	dim cn
	dim cm
	dim strSuccess
	dim i
	dim strElement
		
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	cn.BeginTrans

	strSuccess = request("txtStatusText")

	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
	
	cm.CommandText = "spUpdatePilotStatus"	
		
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = clng(request("txtID"))
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
	p.Value = clng(request("cboStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UserID", 3,  &H0001)
	p.Value = clng(request("txtUserID"))
	cm.Parameters.Append p
		
	Set p = cm.CreateParameter("@UserName", 200,  &H0001, 80)
	p.Value = left(request("txtUserName"),80)
	cm.Parameters.Append p
		
	Set p = cm.CreateParameter("@Comments", 200,  &H0001,255)
	p.Value = left(request("txtComments"),255)
	cm.Parameters.Append p
	
	'This date will to set to "Now" in the stored procedure if no date is specified
	Set p = cm.CreateParameter("@PilotDate", 135,  &H0001)
	if isdate(request("txtPilotDate")) then
		p.Value = cdate(request("txtPilotDate"))
	else
		p.value = null
	end if
	cm.Parameters.Append p
	
	cm.Execute rowschanged
	
	set cm=nothing

	if rowschanged <> 1 then
		strSuccess = "0"
	end if	

	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	cn.Close
	set cn = nothing
	

	'Send Email if necessary
	if strSuccess <> "0" and trim(request("txtStatusLoaded")) <> trim(request("cboStatus")) and ( clng(request("cboStatus")) = 3 or clng(request("cboStatus")) = 5 or clng(request("cboStatus")) = 4  or clng(request("cboStatus")) = 6   or clng(request("cboStatus")) = 7 ) then
	
		dim strSubject
		dim strBody
		
		strSubject = request("txtVendor") & " " & request("txtDeliverable") & " [" & trim(request("txtVersion")) & "] set to " & request("txtStatusText") & " on " & request("txtProduct") 

		strBody = "<font size=2 color=black face=Verdana><b>" & strSubject & "</b></font><BR><BR>"
		strBody = strBody & "<font size=2 color=black face=Verdana>"
		strBody = strBody & "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD></tr>"
		strBody = strBody & "<TR>"
		if trim(request("txtVersionID")) <> "" then
			strBody = strBody & "<TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & request("txtVersionID") & """>" & request("txtVersionID") & "</a></TD>"
		else
			strBody = strBody & "<TD>&nbsp;</TD>"
		end if

		strBody = strBody & "<TD>" & request("txtProduct") & "</TD>"
		strBody = strBody &  "<TD>" & request("txtVendor") & "</TD>"
		strBody = strBody &  "<TD>" & request("txtDeliverable") & "</TD>"
		strBody = strBody &  "<TD>" & request("txtVersion") & "</TD>"
		strBody = strBody &  "<TD>" & request("txtModel") & "</TD>"
		strBody = strBody & "<TD>" & request("txtPartNumber") & "</TD>"
		strBody = strBody & "<TD>" & request("txtComments") & "</TD>"
		strBody = strBody & "</table></font>"

		if clng(request("cboStatus")) = 3 then
			strSubject = "Pilot Hold Notification"
		elseif clng(request("cboStatus")) = 4 then
			strSubject = "Pilot Cancellation Notification"
		elseif clng(request("cboStatus")) = 5 then
			strSubject = "Pilot Failure Notification"
		elseif clng(request("cboStatus")) = 6 then
			strSubject = "Pilot Complete Notification"
		elseif clng(request("cboStatus")) = 7 then
			strSubject = "Factory Hold Notification"
		else
			strSubject = "Pilot Status Updated"
		end if

		strTo = "TWNPDCNBCommodityTechnology@hp.com;kidwell.proceng@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;NotebookCommodityPlanningTeam@hp.com;twnpdccnbcommoditypm@hp.com"
		if request("txtPMEmail") <> "" then
			strTo = strTo & ";" & request("txtPMEmail")
		end if
		if request("txtDevEmail") <> "" then
			strTo = strTo & ";" & request("txtDevEmail")
		end if

		Select case trim(request("txtPartnerID"))
		case "2"
			strTo = strTo & ";IPC-ED1@inventec.com;IPCHP-Excalibur@inventec.com"
		case "7"
			strTo = strTo 
		case "3"
			strTo = strTo & ";A32KeyCommodity@compal.com"
		case "4"
			strTo = strTo & ";TOPCommodityTeam@quantacn.com"
		end select

		if clng(request("cboStatus")) = 7 then 'Factory Hold
    		strTo = strTo & ";notebook.npi.mm@hp.com"
        end if

		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		if trim(request("txtProduct")) = "Test Product 1.0" then
			oMessage.To = "max.yu@hp.com"
		else
			oMessage.To =  strTo 
		end if
		oMessage.Subject = strSubject
		'oMessage.Importance = cdoHigh
		
		'strBody = "TO: " & strTo & "<BR><BR>" & strBody
		
		
		oMessage.HTMLBody = strBody '"<font size=2 face=verdana color=black>" & replace(request("txtFailBody"),"##ShowComments.##",strBodyComments) & "</font>"
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 		
	end if	


%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input type="hidden" id="txtVersionID" name="txtVersionID" value="<%=request("txtVersionID")%>" />
<input type="hidden" id="txDeliverableVersionID" name="txDeliverableVersionID" value="<%=request("txtID")%>" />
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=Request("txtTodayPageSection")%>" />

</BODY>
</HTML>


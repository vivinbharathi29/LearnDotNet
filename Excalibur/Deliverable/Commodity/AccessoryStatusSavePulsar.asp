<%@ Language=VBScript %>

<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    $(document).ready(function () {
        if ($("#txtSuccess")) {            
            if ($("#txtSuccess").val() != "0" && $("#txtKeepItOpen").val() == "false") {
                if ('<%=Request("pulsarplusDivId")%>' != undefined && '<%=Request("pulsarplusDivId")%>' != "") {
                    parent.window.parent.closeExternalPopup();
                }
                else if (IsFromPulsarPlus()) {                   
                    ClosePulsarPlusPopup();
                }
                else {
                    window.parent.closewindow($("#txtVersionID").val(), $("#txtRowID").val(), $("#txtSuccess").val(), $("#txtTodayPageSection").val());
                }
            }
            else {
                window.parent.SetNewStatus($("#txtRowID").val(), $("#txtVersionID").val(), $("#txtSuccess").val(), $("#txtTodayPageSection").val());
                document.location = txtRedirect.value;
                window.parent.repositionParentWindow();
            }
        }
    });



//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript>

<%
	dim cn
	dim cm
	dim strSuccess
	dim i
	dim strElement
	dim blnKeepItOpen
    dim strRedirect 

    blnKeepItOpen = request("txtKeepItOpen")
    strRedirect = request("txtRedirect")

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

    strSuccess = request("txtStatusText")
	
	'cn.BeginTrans

	if trim(request("cboStatus")) = "0" and request("chkDelete") = "on" then
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn

		cm.CommandText = "spUnLinkVersionFromProductRelease"	
		
		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = clng(request("txtProdID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
		p.Value = clng(request("txtVersionID"))
		cm.Parameters.Append p
	
	
	else

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
	
		cm.CommandText = "spUpdateAccessoryStatusPulsar"	
			
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

		Set p = cm.CreateParameter("@KitNumber", 200,  &H0001, 20)
		p.Value = left(request("txtKitNumber"),20)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@KitDescription", 200,  &H0001, 120)
		p.Value = left(request("txtKitDescription"),120)
		cm.Parameters.Append p

		'This date will to set to "Now" in the stored procedure if no date is specified
		Set p = cm.CreateParameter("@AccessoryDate", 135,  &H0001)
		if isdate(request("txtAccessoryDate")) then
			p.Value = cdate(request("txtAccessoryDate"))
		else
			p.value = null
		end if
		cm.Parameters.Append p
	
	end if

	cm.Execute rowschanged
	
	set cm=nothing

	'if rowschanged <> 1 then
	'	strSuccess = "0"
	'end if	
       
	if strSuccess = "0" then
		'cn.RollbackTrans
	else
		'cn.CommitTrans
        if request("txtTodayPageSection") = "" then 
            set rs = server.CreateObject("ADODB.recordset")
            rs.open "usp_ProductDeliverable_GetAccessoryStatus " & clng(request("txtVersionID")) & "," & clng(request("txtProdID")),cn
	        if not (rs.EOF and rs.BOF) then
	            strSuccess = rs("RlsQStatus")
	        end if
	        rs.Close
        end if        
	end if

	'Lookup new status if setting link.
	if strSuccess <> "0" and trim(request("txtStatusLoaded")) <> trim(request("cboStatus")) and ( clng(request("cboStatus")) = -1) then
        set rs = server.CreateObject("ADODB.recordset")
		rs.Open "Select a.Name as Status from AccessoryStatus a with (NOLOCK) inner join Product_Deliverable_Release pdr with (NOLOCK) on pdr.AccessoryStatusID = a.id where pdr.id = " & clng(request("txtID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
		    strSuccess = "&nbsp;"
		else
		    strSuccess = rs("Status") & ""
		end if
		set rs = nothing
	end if
	
	cn.Close
	set cn = nothing
	

	'Send Email if necessary
	if strSuccess <> "0" and trim(request("txtStatusLoaded")) <> trim(request("cboStatus")) and ( clng(request("cboStatus")) = 3 or clng(request("cboStatus")) = 5 or clng(request("cboStatus")) = 4  or clng(request("cboStatus")) = 6 ) then
	
		dim strSubject
		dim strBody
		
		strSubject = "Accessory Status for " & request("txtVendor") & " " & request("txtDeliverable") & " [" & trim(request("txtVersion")) & "] set to " & request("txtStatusText") & " on " & request("txtProduct") 

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
			strSubject = "Accessory Hold Notification"
		elseif clng(request("cboStatus")) = 4 then
			strSubject = "Accessory Cancellation Notification"
		elseif clng(request("cboStatus")) = 5 then
			strSubject = "Accessory Failure Notification"
		elseif clng(request("cboStatus")) = 6 then
			strSubject = "Accessory Testing Complete"
		else
			strSubject = "Accessory Status Updated"
		end if

		strTo = "max.yu@hp.com" 
		if request("txtPMEmail") <> "" then
			strTo = strTo & ";" & request("txtPMEmail")
		end if
		if request("txtDevEmail") <> "" then
			strTo = strTo & ";" & request("txtDevEmail")
		end if


		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = request("txtUserEmail")
		if trim(request("txtProduct")) = "Test Product 1.0" then
			oMessage.To = "max.yu@hp.com"
		else
			oMessage.To = strTo 
		end if
		oMessage.Subject = strSubject
		'oMessage.Importance = cdoHigh
		
		'strBody = "TO: " & strTo & "<BR><BR>" & strBody
		
		
		oMessage.HTMLBody = strBody '"<font size=2 face=verdana color=black>" & replace(request("txtFailBody"),"##ShowComments.##",strBodyComments) & "</font>"
		oMessage.DSNOptions = cdoDSNFailure
		'oMessage.Send 
		Set oMessage = Nothing 		
	end if	


%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input type="hidden" id="txtRedirect" name="txtRedirect" value="<%=strRedirect%>" />
<input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="<%=blnKeepItOpen%>" />
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=Request("txtTodayPageSection")%>" />
<INPUT type="hidden" id="txtVersionID" name="txtVersionID" value="<%=request("txtVersionID")%>">
<INPUT type="hidden" id="txtProductDeliverableID" name="txtProductDeliverableID" value="<%=request("txtProductDeliverableID")%>">
<INPUT type="hidden" id="txtProdDelRelID" name="txtProdDelRelID" value="<%=request("txtProdDelRelID")%>">
<input type="hidden" id="txtRowID" name="txtRowID" value="<%=Request("txtRowID")%>" />
</BODY>
</HTML>


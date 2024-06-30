<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        
        if (txtSuccess.value != "-1") {
            if (IsFromPulsarPlus()) {
                window.parent.parent.parent.ShowCommodityPropertiesCallback(txtSuccess.value);
                ClosePulsarPlusPopup();
            }
            else {
                var iframeName = parent.window.name;
                if (iframeName != '') {
                    parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
                } else {
                    window.returnValue = txtSuccess.value;
                    window.close();
                }
            }
        }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

Saving Product Information.&nbsp; Please Wait...

<%

	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateProductVersionCommodity"
	cm.CommandType =  &H0004

	set p =  cm.CreateParameter("@ID", 3, &H0001)
	p.value = clng(request("txtID"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PDEID", 3, &H0001)
	p.value = clng(request("cboPDE"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ServicePM", 3, &H0001)
	p.value = clng(request("cboServicePM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ODMTestLeadID", 3, &H0001)
	p.value = clng(request("cboODMTestLead"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@WWANTestLeadID", 3, &H0001)
	p.value = clng(request("cboWWANTestLead"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OnCommodityMatrix",11, &H0001)
	if request("chkCommodities") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@WWANProduct",11, &H0001)
	if request("cboWWAN") = "" then
		p.Value = null
	else
		p.Value = clng(request("cboWWAN"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CommodityLock",11, &H0001)
	if request("chkCommodityLock") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SetStatus",16, &H0001)
	if request("txtServicePM") = "" or (request("txtServicePM") = "1" and  request("tagProductStatus") <> "4" and  request("tagProductStatus") <> "5"  ) then
		p.Value = 0 'No changes to status
		intNewStatus = 0
	elseif request("tagProductStatus")="4" and request("chkProductStatus") = "" then
		p.Value = 5 'Set to Product Status 5
		intNewStatus = 5
	elseif request("tagProductStatus")="5" and request("chkProductStatus") = "1" then
		p.Value = 4 'Set to Product Status 4
		intNewStatus = 4
	else
		p.value=0 'Default to no chnages
		intNewStatus = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ServiceLifeDate", 135, &H0001)
	if request("txtServiceEndDate") <> "" then
		p.Value = CDate(request("txtServiceEndDate"))
	else
		p.value = null
	end if
	cm.Parameters.Append p


	cm.Execute RowsEffected
	Set cm=nothing

	if cn.Errors.Count > 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""-1"">"
		Response.Write "<font size=2 face=verdana><b>Unable to update this product.</b></font>"
		cn.RollbackTrans
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""" & clng(request("cboPDE")) & """>"
		cn.CommitTrans
		Response.Write "<font size=2 face=verdana><b><BR><BR>Product Information Saved.</b></font>"
	end if

	
	if intNewStatus = 4 or intNewStatus = 5 then
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
		oMessage.From = request("txtCurrentUserEmail")
		oMessage.To= "max.yu@hp.com"
		if trim(request("txtID")) <> "100" then
			oMessage.CC = "tammy.schapiro@hp.com;rick.rostonski@hp.com;steve.bachmeier@hp.com" 
		end if
		if intNewStatus = 4 then
			oMessage.Subject = request("txtProductName") & " has been set to Post-Production in Excalibur" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductName")  & " has been transitioned to ""Post-Production"" status in Excalibur.</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
		else	
			oMessage.Subject = request("txtProductName") & " has been deactivated in Excalibur" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductName") & " has been transitioned to ""Inactive"" status in Excalibur.</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
		end if
		'oMessage.Importance = cdoHigh
		
		oMessage.Send 
		Set oMessage = Nothing 	
	end if
	dim strTO
	
	dim PartnerID
	dim ServerName
	
	'Send Commodity PM Reassignment email if needed
	if trim (request("tagCommodityPM")) <> trim(request("cboPDE"))and trim(request("cboPDE")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		rs.open "spGetEmployeeByID " & request("cboPDE"),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strTO = "max.yu@hp.com"
			PartnerID = "1"
		else
			strTO = rs("Email") & ""
			PartnerID = rs("PartnerID") & ""
		end if
		rs.Close

		if trim(Partnerid)="1" then
			ServerName = "http://" & Application("Excalibur_ServerName")
		else
			ServerName = "https://" & Application("Excalibur_ODM_ServerName")
		end if

		if lcase(strTO) <> lcase(request("txtCurrentUserEmail")) then
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
			oMessage.From = request("txtCurrentUserEmail")
			if clng(request("txtID")) = 100 then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= strTO
			end if
			oMessage.Subject = request("txtProductName") & " Commodity PM Assignment" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the Commodity PM for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Excalibur or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?Commodity=1&ID=" &  clng(request("txtID")) & """>Assign New Commodity PM</a></font></font>" 
			oMessage.Send 
			Set oMessage = Nothing 	
		end if
	end if	
	
'	'Send SETestLead Reassignment email
'	if trim (request("tagSETestLead")) <> trim(request("cboSETestLead"))and trim(request("cboSETestLead")) <> "0" then
		
'		set rs = server.CreateObject("ADODB.recordset")
		
'		rs.open "spGetEmployeeByID " & request("cboSETestLead"),cn,adOpenStatic
'		if rs.EOF and rs.BOF then
'			strTO = "max.yu@hp.com"
'			PartnerID = "1"
'		else
'			strTO = rs("Email") & ""
'			PartnerID = rs("PartnerID") & ""
'		end if
'		rs.Close

'		if trim(Partnerid)="1" then
'			ServerName = "http://" & Application("Excalibur_ServerName")
'		else
'			ServerName = "https://" & Application("Excalibur_ODM_ServerName")
'		end if
'
'		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
'		'Set oMessage.Configuration = Application("CDO_Config")
'		oMessage.From = request("txtCurrentUserEmail")
'		if clng(request("txtID")) = 100 then
'			oMessage.To= "max.yu@hp.com"
'		else
'			oMessage.To= strTO
'		end if
'		oMessage.CC = "max.yu@hp.com"
'		oMessage.Subject = request("txtProductName") & " SE Test Lead Assignment" 
'		oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the SE Test Lead for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Excalibur or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?Commodity=1&ID=" &  clng(request("txtID")) & """>Assign New SE Test Lead</a></font></font>" 
'		oMessage.Send 
'		Set oMessage = Nothing 	
'	end if		
	
	'Send ODMTestLead Reassignment email
	if trim (request("tagODMTestLead")) <> trim(request("cboODMTestLead"))and trim(request("cboODMTestLead")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		rs.open "spGetEmployeeByID " & request("cboODMTestLead"),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strTO = "max.yu@hp.com"
			PartnerID = "1"
		else
			strTO = rs("Email") & ""
			PartnerID = rs("PartnerID") & ""
		end if
		rs.Close

		if trim(Partnerid)="1" then
			ServerName = "http://" & Application("Excalibur_ServerName")
		else
			ServerName = "https://" & Application("Excalibur_ODM_ServerName")
		end if
		
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
		oMessage.From = request("txtCurrentUserEmail")
		if clng(request("txtID")) = 100 then
			oMessage.To= "max.yu@hp.com"
		else
			oMessage.To= strTO
		end if
		'oMessage.CC = "max.yu@hp.com"
		oMessage.Subject = request("txtProductName") & " ODM HW Test Lead Assignment" 
		oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the ODM HW Test Lead for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Excalibur or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?Commodity=1&ID=" &  clng(request("txtID")) & """>Assign New ODM HW Test Lead</a></font></font>" 
		oMessage.Send 
		Set oMessage = Nothing 	
	end if	
		
	'Send WWANTestLead Reassignment email
	if trim (request("tagWWANTestLead")) <> trim(request("cboWWANTestLead"))and trim(request("cboWWANTestLead")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		rs.open "spGetEmployeeByID " & request("cboWWANTestLead"),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strTO = "max.yu@hp.com"
			PartnerID = "1"
		else
			strTO = rs("Email") & ""
			PartnerID = rs("PartnerID") & ""
		end if
		rs.Close
		
		if trim(Partnerid)="1" then
			ServerName = "http://" & Application("Excalibur_ServerName")
		else
			ServerName = "https://" & Application("Excalibur_ODM_ServerName")
		end if

		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
		oMessage.From = request("txtCurrentUserEmail")
		if clng(request("txtID")) = 100 then
			oMessage.To= "max.yu@hp.com"
		else
			oMessage.To= strTO
		end if
		'oMessage.CC = "max.yu@hp.com"
		oMessage.Subject = request("txtProductName") & " WWAN Test Lead Assignment" 
		oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the WWAN Test Lead for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Excalibur or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?Commodity=1&ID=" &  clng(request("txtID")) & """>Assign New WWAN Test Lead</a></font></font>" 
		oMessage.Send 
		Set oMessage = Nothing 	
	end if	

	set rs = nothing
	cn.Close
	set cn=nothing

%>
     <INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=request("hdnApp")%>">
</BODY>
</HTML>

<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../../Scripts/PulsarPlus.js"></script>
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (txtSuccess.value != "-1") {
        var iframeName = parent.window.name;
        if (IsFromPulsarPlus()) {
                window.parent.parent.parent.ProductMissingWithHWPMCallback(txtSuccess.value);
                ClosePulsarPlusPopup();
        }
        else {
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
	dim FieldsSet
	FieldsSet = 0
	
	if clng(request("cboProcessorPM")) <> 0 then
		FieldsSet = FieldsSet+ 1
	end if
	if clng(request("cboCommHWPM")) <> 0 then
		FieldsSet = FieldsSet+ 1
	end if
	if clng(request("cboGraphicsControllerPM")) <> 0 then
		FieldsSet = FieldsSet+ 1
	end if
	if clng(request("cboVideoMemoryPM")) <> 0 then
		FieldsSet = FieldsSet+ 1
	end if

	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateProductVersionHWPM"
	cm.CommandType =  &H0004

	set p =  cm.CreateParameter("@ID", 3, &H0001)
	p.value = clng(request("txtID"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ProcessorPMID", 3, &H0001)
	p.value = clng(request("cboProcessorPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@CommHWPMID", 3, &H0001)
	p.value = clng(request("cboCommHWPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@GraphicsControllerPMID", 3, &H0001)
	p.value = clng(request("cboGraphicsControllerPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@VideoMemoryPMID", 3, &H0001)
	p.value = clng(request("cboVideoMemoryPM"))
	cm.Parameters.Append p

	cm.Execute RowsEffected
	Set cm=nothing

	if cn.Errors.Count > 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""-1"">"
		Response.Write "<font size=2 face=verdana><b>Unable to update this product.</b></font>"
		cn.RollbackTrans
	else
		if clng(request("txtFieldsRequested")) = clng(FieldsSet) then
			Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
		else
			Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		end if
		cn.CommitTrans
		Response.Write "<font size=2 face=verdana><b><BR><BR>Product Information Saved.</b></font>"
	end if

	dim strTO
	dim PartnerID
	dim ServerName
	
	
'	'Send Comm HW PM Reassignment email if needed
	if trim (request("tagCommHWPM")) <> trim(request("cboCommHWPM"))and trim(request("cboCommHWPM")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		'Lookup the New PM
		rs.open "spGetEmployeeByID " & request("cboCommHWPM"),cn,adOpenStatic
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
			Set oMessage = New EmailWrapper 
			oMessage.From = request("txtCurrentUserEmail")
			if clng(request("txtID")) = 100 then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= strTO
			end if
			'oMessage.cc ="max.yu@hp.com"
			oMessage.Subject = request("txtProductName") & " Comm HW PM Assignment" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the Comm HW PM for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Pulsar or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?HWPM=1&ID=" &  clng(request("txtID")) & """>Assign New Comm HW PM</a></font></font>" 
			oMessage.Send 
			Set oMessage = Nothing 	
		end if
	end if		
		
	
'	'Send Processor PM Reassignment email if needed
	if trim (request("tagProcessorPM")) <> trim(request("cboProcessorPM"))and trim(request("cboProcessorPM")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		'Lookup the New PM
		rs.open "spGetEmployeeByID " & request("cboProcessorPM"),cn,adOpenStatic
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
			Set oMessage = New EmailWrapper 
			oMessage.From = request("txtCurrentUserEmail")
			if clng(request("txtID")) = 100 then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= strTO
			end if
			'oMessage.cc ="max.yu@hp.com"
			oMessage.Subject = request("txtProductName") & " Processor PM Assignment" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the Processor PM for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Pulsar or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?HWPM=1&ID=" &  clng(request("txtID")) & """>Assign New Processor PM</a></font></font>" 
			oMessage.Send 
			Set oMessage = Nothing 	
		end if
	end if		
	
'	'Send Graphics Controller PM Reassignment email if needed
	if trim (request("tagGraphicsControllerPM")) <> trim(request("cboGraphicsControllerPM"))and trim(request("cboGraphicsControllerPM")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		'Lookup the New PM
		rs.open "spGetEmployeeByID " & request("cboGraphicsControllerPM"),cn,adOpenStatic
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
			Set oMessage = New EmailWrapper 
			oMessage.From = request("txtCurrentUserEmail")
			if clng(request("txtID")) = 100 then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= strTO
			end if
			'oMessage.cc ="max.yu@hp.com"
			oMessage.Subject = request("txtProductName") & " Graphics Controller PM Assignment" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the Graphics Controller PM for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Pulsar or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?HWPM=1&ID=" &  clng(request("txtID")) & """>Assign New Graphics Controller PM</a></font></font>" 
			oMessage.Send 
			Set oMessage = Nothing 	
		end if
	end if	

'	'Send Video Memory PM Reassignment email if needed
	if trim (request("tagVideoMemoryPM")) <> trim(request("cboVideoMemoryPM"))and trim(request("cboVideoMemoryPM")) <> "0" then
		
		set rs = server.CreateObject("ADODB.recordset")
		
		'Lookup the New PM
		rs.open "spGetEmployeeByID " & request("cboVideoMemoryPM"),cn,adOpenStatic
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
			Set oMessage = New EmailWrapper 
			oMessage.From = request("txtCurrentUserEmail")
			if clng(request("txtID")) = 100 then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= strTO
			end if
			'oMessage.cc ="max.yu@hp.com"
			oMessage.Subject = request("txtProductName") & " Video Memory PM Assignment" 
			oMessage.HTMLBody = "<font face=Arial size=2 color=black>You have been assigned as the Video Memory PM for " & request("txtProductName")  & ".  Please use the Edit Product link on the product page in Pulsar or the link provided below to reassign this product to someone else on your team if necessary.<BR><BR><a href=""" & ServerName & "/Excalibur.asp"">Open Pulsar</a>&nbsp;|&nbsp;<a href=""" & ServerName & "/mobilese/today/programs.asp?HWPM=1&ID=" &  clng(request("txtID")) & """>Assign New Video Memory PM</a></font></font>" 
			oMessage.Send 
			Set oMessage = Nothing 	
		end if
	end if	
	
	set rs = nothing
	cn.Close
	set cn=nothing

%>

</BODY>
</HTML>

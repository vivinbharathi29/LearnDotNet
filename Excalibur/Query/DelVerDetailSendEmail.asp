<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<TITLE>Excalibur</TITLE>

<%


	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")


	//Get User
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

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserSysAdmin = rs("SystemAdmin")
		CurrentUserPartner = trim(rs("PartnerID") & "")
		CurrentUserEmail = rs("Email")
		if rs("Domain") = "asiapacific" then
			strSite = 2
		else
			strSite = 1
		end if
	end if
	rs.Close

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	if CurrentUserSysAdmin then
		blnAdmin = true
	else
		blnAdmin = false
	end if
	

	strDisplayedID = request("ID")
	
	if request("ID") = "" then 'no deliverable version is specified
		Response.Write "No deliverable version is specified"		
	else
		'Check for valid Root ID
		blnIDFound = true
		if request("RootID") = "" and request("ID") <> "" then
			rs.Open "spGetRootID " & clng(request("ID")),cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				blnIDFound = false	
			else
				strRootID = rs("ID") & ""
			end if
			rs.Close
		else
			strRootID = request("RootID")
		end if
	
		if blnIDFound then
			'get version properties
			rs.Open " spGetVersionProperties4Web " & clng(Request("ID")),cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtID name=ID value=0>"
				Response.Write "<font size=2 face=verdana>Unable to find the requested deliverable version. (" & Request("ID") & ")</font>"
				blnIDFound = false
			else
				blnIDFound = true
				if instr(rs("Location") & "","Workflow Complete")>0 then
					strPreReleased = "1"
				end if
				strCertification = rs("Certification") & ""
				strCertificationStatus = rs("CertificationStatus") & ""
				strCertificationDate = rs("CertificationDate") & ""
				strCertificationID = rs("CertificationID") & ""
				strCertificationComments = rs("CertificationComments") & ""
				strCATStatus = rs("CATStatus") & ""
				strCategoryName = rs("category") & ""
				strModelNumber = rs("ModelNumber") & ""
				CurrentWorkflowLocation = rs("Location") & ""
				strTypeID = rs("TypeID") & ""
				strPart = request("PartNumber")
				strCDPart = rs("CDPartNumber") & ""
				strKitNumber = rs("CDKitNumber") & ""
				if trim(strKitNumber) = "N/A" then
					strChkKitNumber = ""
				else
					strChkKitNumber = "checked"
				end if			
				strLevelID = rs("LevelID") & ""
				strDependencies = rs("SWDependencies") & ""
				strPNPDevices = rs("PNPDevices") & ""
				if trim(rs("DeliverableName") & "") = "" then
					strDelName = rs("Name") & ""
				else
					strDelName = rs("DeliverableName") & ""
				end if
				strVersion = rs("Version") & ""
				strRevision = rs("Revision") & ""
				strPass = rs("Pass") & ""			
				strFilename = rs("Filename") & ""	
				strCodeName = rs("CodeName") & ""		
				strTransfer = rs("ImagePath") & ""			
				strDevID = rs("DeveloperID") & ""
				strDevManager = rs("DevManager") & ""
				strVendorVersion = rs("VendorVersion") & ""			
				strSoftpaqNumber = rs("SoftpaqNumber") & ""			
				strSoftpaqFixes = rs("SoftpaqFixes") & ""			
				strSoftpaqFileInfo = rs("SoftpaqFileInfo") & ""
				strSoftpaqInstructions = rs("SoftpaqInstructions") & ""
				strSysReboot = replace(replace(rs("Reboot") & "","1","checked"),"0","")
				strIconDesktop = replace(replace(rs("IconDesktop") & "","True","checked"),"False","")
				strIconMenu = replace(replace(rs("IconMenu") & "","True","checked"),"False","")
				strIcontray = replace(replace(rs("IconTray") & "","True","checked"),"False","")
				strIconPanel = replace(replace(rs("IconPanel") & "","True","checked"),"False","")
				strDeveloperNotification=replace(replace(rs("DeveloperNotification") & "","1","checked"),"0","")
				strPackageForWeb = replace(replace(rs("PackageForWeb") & "","True","checked"),"False","")
				strPropertyTabs = rs("PropertyTabs") & ""
				strSupersedes = rs("Supersedes") & ""
				strComments = rs("Comments") & ""			
				strChanges = rs("Changes") & ""
				strPNPDevices = rs("PNPDevices") & ""			
				strPart = rs("PartNumber") & ""
				strWorkflowID = rs("workflowid") & ""
				strIntroDate=trim(rs("IntroDate") & "")
				strEOLDate=trim(rs("EOLDate") & "")
				strSamplesDate=trim(rs("SampleDate") & "")
				strRohs = trim(rs("RoHSID") & "")
				strGreenSpec = trim(rs("GreenSpecID") & "")
				strHFCNChecked =  replace(replace(rs("HFCN") & "","True","checked"),"False","")
				strActive =  replace(replace(rs("ActiveVersion") & "","True",""),"False","checked")

				strIntroConfidence = rs("IntroConfidence") & ""
				strSamplesConfidence = rs("SamplesConfidence") & ""
			
				'if rs("VersionVendor") & "" <> "" and rs("VersionVendorID") & "" <> "" then
				'	strVendor = rs("VersionVendor") & ""
				'	strVendorID = rs("VersionVendorid") & ""
				'else
				'	strVendor = rs("Vendor") & ""
				'	strVendorID = rs("Vendorid") & ""
				'end if
				strVendor = rs("Vendor") & ""
				strVendorID = rs("Vendorid") & ""
				strVersionVendor = rs("VersionVendor") & ""
				strVersionVendorID = rs("VersionVendorid") & ""
			
			
				strNotes = rs("Notes") & ""
				strDescription = rs("Description") & ""			
				strRootFilename = rs("RootFilename") & ""
				strInstall = rs("Install") & ""
				strOSType = rs("OSType") & ""
				strReplicater = rs("Replicater") & ""
				strSilentInstall = rs("SilentInstall") & ""
				strARCDInstall = rs("ARCDInstall") & ""
				strPreinstall =  replace(replace(rs("Preinstall") & "","1","checked"),"0","")
				strRompaq = replace(replace(rs("Rompaq") & "","1","checked"),"0","")
				strCDImage = replace(replace(rs("CDImage") & "","1","checked"),"0","")
				strISOImage = replace(replace(rs("ISOImage") & "","1","checked"),"0","")
				strAR = replace(replace(rs("AR") & "","1","checked"),"0","")
				strCAB = replace(replace(rs("CAB") & "","1","checked"),"0","")
				strBinary = replace(replace(rs("Binary") & "","1","checked"),"0","")
				strFloppy = replace(replace(rs("FloppyDisk") & "","1","checked"),"0","")
				strPreinstallROM = replace(replace(rs("PreinstallROM") & "","1","checked"),"0","")
				strAdmin = replace(replace(rs("Admin") & "","1","checked"),"0","")
				strEndUserInst = rs("EndUserInstructions") & ""
				strSilentInst = rs("SilentInstructions") & ""
				strPackageType = rs("PackageType") & ""
				strReleaseType = rs("ReleaseType" & "")
				strMultiLanguage = rs("MultiLanguage") & ""
				strEffectiveDate = rs("EffectiveDate") & ""
				strCategoryID = trim(rs("CategoryID") & "")
			
				if strAR <> "" then
					strTransfer = ""
				end if

				if rs("InstallableUpdate") & "" = "0" then
					strUpdate = ""
				else
					strUpdate = "checked"
				end if
				if not rs("SSMCompliant")  then
					strSSM = ""
				else
					strSSM = "checked"
				end if

				if not rs("Desktops")  then
					strDesktops = ""
				else
					strDesktops = "checked"
				end if

				if not rs("Notebooks")  then
					strNotebooks = ""
				else
					strNotebooks = "checked"
				end if

				if not rs("Workstations")  then
					strWorkstations = ""
				else
					strWorkstations = "checked"
				end if

				if not rs("ThinClients")  then
					strThinClients = ""
				else
					strThinClients = "checked"
				end if

				if not rs("Monitors")  then
					strMonitors = ""
				else
					strMonitors = "checked"
				end if

				if not rs("Projectors")  then
					strProjectors = ""
				else
					strProjectors = "checked"
				end if

				if not rs("Handhelds")  then
					strHandhelds = ""
				else
					strHandhelds = "checked"
				end if

				if not rs("printers")  then
					strprinters = ""
				else
					strprinters = "checked"
				end if

				if not rs("PersonalAudio")  then
					strPersonalAudio = ""
				else
					strPersonalAudio = "checked"
				end if

				if not rs("Scriptpaq")  then
					strScriptpaq = ""
				else
					strScriptpaq = "checked"
				end if

		
				'Load Supported OTS
				rs.Close

				if strMultiLanguage = "1" then
					SoftpaqFrameHeight = 50'67
				else
					SoftpaqFrameHeight = 110 '176
				end if


				strOTSList = ""
				strSelectedOTS = ""
				on error resume next
				rs.open "spGetOTSByDelVersion "  & clng(strDisplayedID),cn,adOpenForwardOnly
				if cn.Errors.count = 0 then		
					do while not rs.EOF
						strOTSList = strOTSList &  rs("OTSNumber") & ","
						strSelectedOTS = strSelectedOTS & "<BR>" & rs("OTSNumber") & " - " &  rs("shortdescription") 
						rs.MoveNext
					loop
    				rs.Close
				end if

	
				'Load MilestoneSteps - save first and second
			

				rs.open "spGetFirstMilestone "  & clng(strDisplayedID),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFirstMilestoneID = 0
				else
					strFirstMilestoneID = rs("ID")
				end if

				'Load OS Supported By Version
				rs.Close

				rs.open "spGetSelectedOS "  & clng(strDisplayedID),cn,adOpenForwardOnly
				strSelectedOSIDs = ""
				strSelectedOS = ""
				do while not rs.EOF
					if rs("ID") & "" <> "16" then
						strSelectedOSIDs = strSelectedOSIDs & rs("ID") & ","
						strSelectedOS = strSelectedOS & rs("Name") & "; "
					end if
					rs.MoveNext
				loop
				if strSelectedOSIDs = "" then
					strSelectedOSIDs = "16,"
					strSelectedOS = ""
				end if

				'Load OS Supported By Root
				rs.Close
				rs.open "spGetOSForRoot "  & clng(strRootID),cn,adOpenForwardOnly
				strAvailableOS = ""
				do while not rs.EOF
					if rs("ID") & "" <> "16" then
						if instr("," & strSelectedOSIDs,"," & rs("ID") & ",") = 0 then
							strAvailableOS = strAvailableOS & rs("Name") & "; "
						end if
					end if
					rs.MoveNext
				loop
			
				dim strPartRow
				dim strKitRow
				'Load Langs Supported By Version
				rs.Close

				rs.open "spGetSelectedLanguages "  & clng(strDisplayedID),cn,adOpenForwardOnly
				strSelectedLangIDs = ""
				strSelectedLangs = ""
				strLangRelease = ""
				strPartRow=""
				strKitRow=""
				do while not rs.EOF
					if rs("ID") & "" <> "58" then
						if trim(strTypeID)="1" then
							strSelectedLangIDs = strSelectedLangIDs & rs("ID") & ","
							strSelectedLangs = strSelectedLangs & rs("Name") & "; "
						else
							strSelectedLangIDs = strSelectedLangIDs & rs("ID") & ","
							strSelectedLangs = strSelectedLangs & rs("Abbreviation") & " - " &  rs("Name") & "; "
							strLangRelease = strLangRelease & left(rs("Abbreviation") & "",2) & "Release name=chk" & left(rs("Abbreviation") & "",2) & "Release><label ID=lbl" & left(rs("Abbreviation") & "",2) & "Release>" & rs("Name") & "; "
							strSoftpaqLanguageRows = strSoftpaqLanguageRows & rs("Abbreviation") & rs("Abbreviation")
							if rs("PartNumber") <> "" then 
								strPartRow = strPartRow & rs("Abbreviation") & " - " & rs("Name") & ": " & rs("PartNumber") & "; "
							end if
							if rs("CDKitNumber") <> "" then
								strKitRow = strKitRow & rs("Abbreviation") & " - " & rs("Name") & ": " & rs("CDKitNumber") & "; "	
							end if
							if trim(rs("CDKitNumber")) = "N/A" then
								strChkKitNumber = ""
							else
								strChkKitNumber = "checked"
							end if						
						end if
					end if
					rs.MoveNext
				loop
				if strSelectedLangIDs = "" then
					strSelectedLangIDs = "58,"
					strSelectedLangs =""
				end if

				'Load Langs Supported By Root
				rs.Close
				rs.open "spGetLanguagesForRoot "  & clng(strRootID),cn,adOpenForwardOnly
				strAvailableLangs = ""
				do while not rs.EOF
					if rs("ID") & "" <> "58" then
						if instr("," & strSelectedLangIDs,"," & rs("ID") & ",") = 0 then
							strAvailableLangs = strAvailableLangs & "<Option value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
							strLangRelease = strLangRelease & "<INPUT checked type=""checkbox"" id=chk" & left(rs("name") & "",2) & "Release name=chk" & left(rs("name") & "",2) & "Release><label ID=lbl" & left(rs("name") & "",2) & "Release>" & rs("Name") & "<BR></label>"
							strSoftpaqLanguageRows = strSoftpaqLanguageRows & "<TR ID=SPRow" & left(rs("name") & "",2) & " style=""Display:none""><TD>" & left(rs("name") & "",2) & "</td><td><input id=""txtSP" & rs("ID") & """ name=""txtSP" & rs("ID")  & """ style=""WIDTH: 120px;"" maxlength=10></td><td><input id=""txtSPSup" & rs("ID") & """ name=""txtSPSup" & rs("ID") & """ style=""WIDTH: 100%"" maxlength=2048></TD></TR>"
							'strPartRow = strPartRow &  rs("Name")
							'strKitRow = strKitRow & rs("Name") 	
							strChkKitNumber = "checked"
						end if
					end if
					rs.MoveNext
				loop
			
			
				rs.Close
				
				'Load Products Supported by Version
				rs.open "spGetProductsForVersion " & clng(strDisplayedID),cn,adOpenForwardOnly
				strSelectedProductIDs = ""
				strSelectedProducts = ""
				do while not rs.EOF
					strSelectedProductIDs = strSelectedProductIDs  & rs("ID") & ","
					strSelectedProducts = strSelectedProducts & rs("Family") & " " & rs("Version") & "; "
					rs.MoveNext
				loop

				rs.Close

				strSQL = "spGetTargetedProductsForVersion " & clng(strDisplayedID)
				rs.Open strSQL,cn,adOpenStatic
  
				strTargetedProducts = ""
				strAllProducts = ""
				do while not rs.EOF
					strAllProducts = strAllProducts & ", " & rs("Family") & "&nbsp;" & rs("Version")
					if rs("Targeted") then
						strTargetedProducts = strTargetedProducts & ", " & rs("Family") & "&nbsp;" & rs("Version")
					end if
					rs.MoveNext
				loop
				rs.close			
			
				'Load Products Supported by Root
				rs.open "spGetProductsForRoot " & clng(strRootID),cn,adOpenForwardOnly
				strAvailableProducts = ""
				do while not rs.EOF
					if instr("," & strSelectedProductIDs,"," & rs("ID") & ",") = 0 then
						if rs("active") & "" <> "0" or rs("Sustaining") & "" <> "0" then
							strAvailableProducts = strAvailableProducts & "<Option value=" & rs("ID") & ">" & rs("Name") & " " & rs("Version") & "</OPTION>"
						end if
					end if
					rs.MoveNext
				loop
			end if
			rs.close
		end if
	end if
	
	
	
if instr(CurrentWorkflowLocation,"Workflow Complete")>0 and request("ID") <> "" then
	blnFound = false
	rs.open "spListPMsActive",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(rs("ID")) = trim(CurrentUSerID) then
			blnFound=true
			exit do
		end if
		rs.MoveNext
	loop
	rs.Close

	if blnFound and (trim(strTypeID) = "1") then
		blnRemoveProductAccess = true
	else
		blnRemoveProductAccess = false
	end if
	
end if

	dim ShowCertification
	if strCertification = "1" then
		ShowCertification = ""
	else
		ShowCertification = "none"
	end if
	if left(strFilename,5) = "HFCN_" then
		strHFCN="1"
	end if
	
	dim strTitleColor
	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0


	dim strDeliverable
	strDeliverable = strDelName & " [" & strVersion
	if strRevision <> "" then
		strDeliverable = strDeliverable & "," & strRevision
	end if
	if strPass <> "" then
		strDeliverable = strDeliverable & "," & strPass
	end if
	strDeliverable = strDeliverable & "]"
	Response.Write "<font face=verdana size=2><b>" & strDeliverable & " Details</b></font><BR><BR>"

    if strVendor="< Multiple Suppliers >" then
		strSQL = "spGetVendorList"
		rs.Open strSQL,cn,adOpenForwardOnly
		strVendor = ""
		do while not rs.EOF
			if rs("ID") <> 203 then
				if trim(strVersionVendorID) = trim(rs("ID") & "") then
					strVendor = strVendor & rs("Name")
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
	end if

	strSQL = "spGetEmployees"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		if strDevID = rs("ID") & "" or (request("ID") = "" and trim(CurrentUserId) = trim(rs("ID") & "") )then
			strDevname = rs("Name")	
			exit do
		end if
		rs.MoveNext
	loop
	rs.Close

%>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function PrintWindow(){
	//PrintLink.style.display = "none";
	chkNA.disabled = false;
	chkEMEA.disabled = false;
	chkCKK.disabled = false;
	chkAPD.disabled = false;
	chkGCD.disabled = false;
	chkLA.disabled = false;
	chkCustomer.disabled = false;
	chkAdd.disabled = false;
	chkModify.disabled = false;
	chkRemove.disabled = false;
	chkStatus.disabled = false;
	window.print();
	window.close();
}

function MailWindow(){
	if (frmSend.txtTo.value == "")
		{
			window.alert("Please enter email recipients first.");
			frmSend.txtTo.focus();
			return;
		}
	//PrintLink.style.display = "none";
	frmSend.txtEmailBody.value = ItemDetails.innerHTML;
	frmSend.submit();
}
function window_onload() {
	if (typeof(frmSend.txtTo) != "undefined")	
		frmSend.txtTo.focus();
}


function ChooseEmail(FieldID) {
	var strResult;
    if (FieldID == 1)      
        strResult = window.showModalDialog("../Email/AddressBook.asp?AddressList=" + frmSend.txtTo.value, "fromRoot", "dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");        
	else   
        strResult = window.showModalDialog("../Email/AddressBook.asp?AddressList=" + frmSend.txtCC.value, "fromRoot", "dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");         

	if (typeof(strResult) != "undefined")
		if (FieldID==1)
			frmSend.txtTo.value = strResult;
		else
			frmSend.txtCC.value = strResult;

}
function AddAddress(strEmail,strBox){
	var strText;
	
	if (strBox == 1)
		{
		strText = trim(frmSend.txtTo.value);
		if (strText == "")
			frmSend.txtTo.value = strEmail;
		else
			{
			if (strText.charAt(strText.length-1) != ";")
				frmSend.txtTo.value = strText + ";" + strEmail;
			else
				frmSend.txtTo.value = strText + strEmail;
			}
		}
	else
		{
		strText = trim(frmSend.txtCC.value);
		if (strText == "")
			frmSend.txtCC.value = strEmail;
		else
			{
			if (strText.charAt(strText.length-1) != ";")
				frmSend.txtCC.value = strText + ";" + strEmail;
			else
				frmSend.txtCC.value = strText + strEmail;
			}
		}
}


//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">

<font size=2 face=verdana>

<Span ID=SendEmailLink>
<font size=2 face=verdana><a href="javascript:MailWindow();">Send Email</a>
<BR></span>
<%
	dim rs 
	dim cn
	dim cm
	dim p
	dim strproducts
	dim strHFCNTeamEmail
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	
	strDescription = DescriptionTemplate
	strHFCNTeamEmail = "NBSCSWEngrs@hp.com"
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovals"
		


	
	if trim(CurrentUserEmail) = "" then
		Response.Write "<BR><BR>Only registered Excalibur users can send email using this function."
		%>
		<SCRIPT>
		//PrintLink.style.display = "none";
		</SCRIPT>
		<%
	else
%>
<BR>
	<form name=frmSend method=post action="DelVerDetailEmail.asp">

	<font size=1 face=verdana color=green>Enter SMTP Email addresses only (i.e., john.doe@hp.com)</font><BR>
	<Table bgcolor=ivory border=1 cellspacing=0 cellpadding=2 width=100%>
	<TR>
		<TD nowrap valign=top width=180><font face=verdana size=1><b>Send To:</b>
		<%if strHFCNTeamEmail <> "" then%>
			<BR>&nbsp;&nbsp;<a href="javascript:AddAddress('<%=strHFCNTeamEmail%>',1)">HFCN Team</a> |
		<%end if%>
		<a href="javascript: ChooseEmail(1);">Lookup Address</a>
		</font></TD>
		<TD valign=top>
			<INPUT style="Width=100%" type="text" id=txtTo name=txtTo>
		</TD>
	</TR>
	<TR>
		<TD nowrap valign=top width=180><font face=verdana size=1><b>CC:</b>
		<%if strHFCNTeamEmail <> "" then%>
			<BR>&nbsp;&nbsp;<a href="javascript:AddAddress('<%=strHFCNTeamEmail%>',2)">HFCN Team</a> |
		<%end if%>
		<a href="javascript: ChooseEmail(2);">Lookup Address</a>
			
		</font></TD>
		<TD valign=top>
			<INPUT style="Width=100%" type="text" id=txtCC name=txtCC>
		</TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Subject:</b></font></TD> 
		<TD valign=top><INPUT style="Width=100%" type="text" id=txtSubject name=txtSubject value="<%=request("ID") & " - " & strDeliverable & " Detail Information"%>"></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Notes:</b></font></TD>
		<TD valign=top><TEXTAREA rows=5 style="Width=100%" cols=20 id=txtNotes name=txtNotes></TEXTAREA></TD>
	</TR>
	</Table><INPUT type="hidden" id=txtFrom name=txtFrom value="<%=CurrentUserEmail%>">
	<TEXTAREA style="Display:none" rows=2 cols=20 id=txtEmailBody name=txtEmailBody></TEXTAREA>
		
	</form>
	
	<span ID=ItemDetails>	

	<font face=verdana size=2><BR><b> <%=strDeliverable%>  Detail Information:</b></font><br><br>

	<Table bgcolor=ivory cellpadding=2 cellspacing=0 border =1 width="100%">
	<TABLE bordercolor=black cellspacing=0 Border=1 width="100%">
	<TR>
	<% if trim(strTypeID) = "1"  then %>
		<TD><TABLE width="100%" >
		<% if strHFCNChecked = "true" then %>
		<TR><TD nowrap><font face=verdana size=1>This is an HFCN release</font></TD></TR>
		<% end if %>
		<TR><TD><font face=verdana size=1><b>Hardware Version:</b></font></TD><TD><font face=verdana size=1><%=strVersion%></font></TD></TR><TR><TD nowrap><font face=verdana size=1><b>Firmware Version:</b></font></TD><TD><font face=verdana size=1><%=strRevision%></font></TD></TR>
		<TR><TD nowrap><font face=verdana size=1><b>Rev:</b></font></TD><TD><font face=verdana size=1><%=strPass%></font></TD></TR>
		</table></TD>
	<% elseif trim(strTypeID) = "3" then %>
		<TD valign=top><TABLE width="100%">
		<TR><TD><font face=verdana size=1><b>ID:</b></font></TD><TD><font face=verdana size=1><a target="_blank" href="http://<%=Application("Excalibur_ServerName")%>/Excalibur/WizardFrames.asp?Type=1&ID=<%= strDisplayedID%>"><%=strDisplayedID%></a></font></TD></TR>
		<TR><TD><font face=verdana size=1><b>Version:</b></font></TD><TD><font face=verdana size=1><%=strVersion%></font></TD></TR>
		<%if trim(strCategoryID) <> "161" and strHFCNChecked = "true" then%>
			<TR><TD nowrap><font face=verdana size=1>This is an HFCN release</font></TD></TR>
		<% end if %>
		</table></TD>
	<% else %>
		<TD valign=top><TABLE width="100%" >
		<TR><TD><font face=verdana size=1><b>ID:</b></font></TD><TD><font face=verdana size=1><a target="_blank" href="http://<%=Application("Excalibur_ServerName")%>/Excalibur/WizardFrames.asp?Type=1&ID=<%= strDisplayedID%>"><%=strDisplayedID%></a></font></TD></TR>
		<TR><TD><font face=verdana size=1><b>Version:</b></font></TD><TD><font face=verdana size=1><%=strVersion%></font></TD></TR>
		<TR><TD nowrap><font face=verdana size=1><b>Revision:</b></font></TD><TD><font face=verdana size=1><%=strrevision%></font></TD></TR><TR><TD><font face=verdana size=1><b>Pass:</b></font></TD><TD><font face=verdana size=1><%=strPass%></font></TD></TR>
		</table></TD>
	<% end if %>
	<TD valign=top><TABLE width="100%" ><TR><TD nowrap><font face=verdana size=1><b>Vendor:</b></font></TD><TD><font face=verdana size=1><%=strVendor%></font></TD></TR>
		<TR><TD nowrap><font face=verdana size=1><b>Vendor Version:</b></font></TD><TD><font face=verdana size=1><%=strVersionVendor%></font></TD></TR>
		<% if trim(strTypeID) = "1"  then %>
			<TR><TD nowrap><font face=verdana size=1><b>HP Part Number:</b></font></TD><TD><font face=verdana size=1><%=strPart%></font></TD></TR>
			<TR><TD nowrap><font face=verdana size=1><b>Model Number:</b></font></TD><TD><font face=verdana size=1><%=strModelNumber%></font></TD></TR>
			</table></TD>
		<% else %>
			<TR><TD nowrap><font face=verdana size=1><b>FileName:</b></font></TD><TD><font face=verdana size=1><%=strFilename%></font></TD></TR>
			</table></TD>
		<% end if %>
	<TD valign=top><TABLE width="100%" >
		<% if trim(strTypeID) = "1"  then %>
			<TR><TD nowrap><font face=verdana size=1><b>Code Name:</b></font></TD><TD><font face=verdana size=1><%=strCodeName%></font></TD></TR>
		<% end if %>
		<TR><TD nowrap><font face=verdana size=1><b>Dev. Manager:</b></font></TD><TD><font face=verdana size=1><%=strDevManager%></font></TD></TR>
		<TR><TD nowrap><font face=verdana size=1><b>Developer:</b></font></TD><TD><font face=verdana size=1><%=strDevname%></font></TD></TR>
		<TR>
		<%if strTypeID = 1 then%>
			<td nowrap><font face=verdana size=1><b>Production&nbsp;Level:</b></font></td>    
		<%else%>
			<td nowrap><font face=verdana size=1><b>Build&nbsp;Level:</b></font></td>    
		<%end if%>
		<td>
			<%
			if strTypeID = 1 then
				rs.Open "spListDeliverableLevels 1" ,cn,adOpenForwardOnly
			else
				rs.Open "spListDeliverableLevels 2" ,cn,adOpenForwardOnly
			end if
			
			do while not rs.EOF
				if trim(strLevelID) = trim(rs("ID") & "") then
					strLevel = rs("name")
					exit do
				end if
				rs.MoveNext
			loop
			rs.Close
			
			%>
						
		</td><TD><font face=verdana size=1><%=strLevel%></font></TD></TR>
		</table></TD>
		</TR>

	<TR><TD colspan=3><TABLE width="100%">

	<%
	Count = 0
	if request("ID") <> "" then
	strSQL = "spGetDelMilestoneList " & clng(strRootID) & "," & clng(request("ID"))
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		strActualDate = rs("Actual") & ""
		if strActualDate = "" then
			strActualDate = "&nbsp;"
		elseif instr(strActualDate," ") > 0 then
			strActualDate = left(strActualDate,instr(strActualDate," ") - 1)
		end if
		
		strMilestone = rs("Milestone")
		strStatus = rs("Status") 
		strPlanned = rs("Planned")
		
%>
		<% if Count = 0 then %>
			<TR><TD align=center nowrap><font face=verdana size=1><u><b>Workflow Step</b></u><BR>
		<% else %>
			<TR><TD align=center nowrap><font face=verdana size=1>
		<% end if %>
		<% =strMilestone %></font></TD>
		
		<% if Count = 0 then %>
			<TD align=center nowrap><font face=verdana size=1><u><b>Status</b></u><BR>
		<% else %>
			<TD align=center nowrap><font face=verdana size=1>
		<% end if %>
		<% =strStatus %></font></TD>
		
		<% if Count = 0 then %>
			<TD align=center nowrap ><font face=verdana size=1><u><b>Planned Date</b></u><BR>
		<% else %>
			<TD align=center nowrap><font face=verdana size=1>
		<% end if %>
		<% =strPlanned %></font></TD>

		<% if Count = 0 then %>
			<TD align=center nowrap><font face=verdana size=1><u><b>Actual Date</b></u><BR>
		<% else %>
			<TD align=center nowrap><font face=verdana size=1>
		<% end if %>
		<% =strActualDate %></font></TD>
		
		<% Count = Count + 1 %>

<%
		rs.MoveNext
	loop
	rs.Close
	end if
	%>

	</TR></td>
	</TABLE></td></tr>


	    <%if trim(strTypeID) = "1" then%>
			<TR><TD colspan=3><TABLE width="100%" >
		<%else%>
			<TR style="Display:none"><TD colspan=3><TABLE width="100%" >
		<%end if%>
		<td nowrap><font face=verdana size=1><b>Samples&nbsp;Available&nbsp;Date:</b>&nbsp;&nbsp;&nbsp;
		<%if strSamplesDate <> "" then %>
			<%=strSamplesDate%>
		<% else %>
			Unknown&nbsp;
		<% end if %>
		&nbsp;&nbsp;&nbsp;<b>Confidence:</b>&nbsp;</font>
		<%if trim(strSamplesConfidence) = "1" then%>
			<font face=verdana size=1 color=green>High</font>
		<%elseif trim(strSamplesConfidence) = "2" then%>
			<font face=verdana size=1 color=black>Medium</font>
		<%elseif trim(strSamplesConfidence) = "3" then%>
			<font face=verdana size=1 color=red>Low</font>
		<%else%>
			<font face=verdana size=1>Unknown&nbsp;</FONT>
		<%end if %>
		</td>
	</TABLE></td></tr>
	
<%if strIntroDate <> "" then%>
	<tr><TD colspan=3><TABLE width="100%" >
	<%if trim(strTypeId) = "1" then%>
		<td nowrap><font face=verdana size=1><b>Mass&nbsp;Production:</b>&nbsp;&nbsp;
	<%else%>
		<td nowrap><font face=verdana size=1><b>Intro&nbsp;Date:</b>&nbsp;&nbsp;
	<%end if%>
	<% if strIntroDate <> "" then %>
		<% =strIntroDate %>
	<% else %>
		Unknown&nbsp;
	<% end if %> 
		&nbsp;&nbsp;<b>Confidence:</b>&nbsp;</font>
		<%if trim(strIntroConfidence) = "1" then%>
			<font face=verdana size=1 color=green>High</font>
		<%elseif trim(strIntroConfidence) = "2" then%>
			<font face=verdana size=1 color=black>Medium</font>
		<%elseif trim(strIntroConfidence) = "3" then%>
			<font face=verdana size=1 color=red>Low</font>
		<%else%>
			<font face=verdana size=1>Unknown&nbsp;</FONT>
		<%end if%>
		</td>
	</TABLE></TD></tr>
<%end if%>

	<% if strEOLDate <> "" or strActive="checked" then %>
	<TR><TD colspan=3><TABLE width="100%" ><tr>
		<td nowrap><font face=verdana size=1><b>Available&nbsp;Until:</b>&nbsp;&nbsp;&nbsp;&nbsp;
		<% if strEOLDate <> "" then %>
			<%=strEOLDate%>
		<% else%>
			Unknown&nbsp;
		<% end if %>
			</FONT>	
		<%if request("ID") = "" then %>
			<span style="Display:none">
		<%else%>
			<span>
		<%end if%>
		<% if strActive="checked" then %> 
			&nbsp;&nbsp;&nbsp;&nbsp;<font size=1 face=verdana>This version is End of Life.</font>
		<% end if %>
		</span></td>
	</TABLE></td></tr>
	<%end if%>    
    
    
    <% if strUpdate <> "" or strPackageForWeb <> "" then %> 
		<%if left(strFilename,5) <> "HFCN_" and  trim(strTypeID) <> "1" then%>
			<TR><TD colspan=3><TABLE width="100%" >
		<%else%>
			<TR  style="Display:none"><TD colspan=3><TABLE width="100%" >
		<%end if%>
		<td nowrap><font face=verdana size=1><b>Special Notes:</b>&nbsp;&nbsp;
		<% if strUpdate <> "" then %> 
			&nbsp;&nbsp;&nbsp;&nbsp;Installable&nbsp;Update,
		<% end if %>
		<% if strPackageForWeb <> "" then %> 
			&nbsp;&nbsp;Package&nbsp;For&nbsp;Web
		<% end if %>
		</font></td>
		</TABLE></td></tr>
	<%end if%>
	
	<%if strIconPanel <> "" or strIconDesktop <> "" or strIconMenu <> "" or strIconTray <> "" then %>

		<%if trim(strTypeID) = "1" then%>
			<TR style="Display:none"><TD colspan=3><TABLE width="100%">
		<%else%>
			<TR><TD colspan=3><TABLE width="100%" >
		<%end if%>
		<td nowrap><font face=verdana size=1><b>Icons Installed:</b>&nbsp;&nbsp;&nbsp;
		<%if strIconDesktop <> "" then %> 
			Desktop,&nbsp;&nbsp;&nbsp;
		<%end if %>
		<%if strIconMenu <> "" then %> 
			Start&nbsp;Menu,&nbsp;&nbsp;	
		<% end if %>
		<%if strIconTray <> "" then %> 
			System&nbsp;Tray,&nbsp;&nbsp;
		<% end if %>	
		<%if strIconPanel <> "" then %>
			Control&nbsp;Panel&nbsp;&nbsp;	
		<% end if %>
		</font></td>
		</TABLE></td></tr>
	
	<%end if%>

	<%if strPropertyTabs <> "" then%>
		<%if trim(strTypeID) = "1" then%>
			<TR style="Display:none"><TD colspan=3><TABLE width="100%">
		<%else%>
			<TR><TD colspan=3><TABLE width="100%" >
		<%end if%>
		<td nowrap><font face=verdana size=1><b>Property Tabs Added:</b>&nbsp;&nbsp;&nbsp;
		<%=replace(strPropertyTabs,"""","&quot;")%>
		</font></td>
		</TABLE></td></tr>
	<%end if%>

	<%
	if strTypeId = 1 then
		strShowROMComps = "none"
		strShowDistributions = "none"
	else
		if strTypeID = "3" and strPreinstall = "" and strFloppy = "" and strCDImage = "" and strScriptPaq = "" then
			strShowDistributions = "none"
		else
			strShowDistributions = ""
		end if

		if strTypeID <> "3" and strBinary = "" and strRompaq = "" and strPreinstallROM = "" and strCAB = "" then
			strShowROMComps = "none"
		else
			strShowROMComps = ""
		end if
	end if
	
	dim strDistributionsEnabled
	
	if strAR <> "" then
		strDistributionsEnabled = "DISABLED"
	else
		strDistributionsEnabled = ""
	end if	
	%>
 	
	<tr style="Display:<%=strShowDistributions%>"><TD colspan=3><TABLE width="100%">
	<td nowrap><font face=verdana size=1><b>Packaging:</b>&nbsp;&nbsp;
	<% if strPreinstall <> "" then %> Preinstall;&nbsp;&nbsp; <% end if %>
	<% if strFloppy <> "" then %> Diskette;&nbsp;&nbsp; <%end if%>	
	<% if strScriptPaq <> "" then %> Scriptpaq;&nbsp;&nbsp; <%end if%>
	<% if strCDImage <> "" then %> CD&nbsp;Files;&nbsp;&nbsp; <%end if%>
	<% if strISOImage <> "" then %> ISO&nbsp;Image;&nbsp;&nbsp; <%end if%>	
	<% if strAR <> "" then %> Replicater&nbsp;Only;&nbsp;&nbsp;	<%end if%>
	</font></td>
	</TABLE></td></tr>

	<tr style="Display:<%=strShowROMComps%>"><TD colspan=3><TABLE width="100%">
	<td nowrap><font face=verdana size=1><b>ROM Components:</b>
	<% if strBinary <> "" then %> Binary;&nbsp;&nbsp; <%end if %>
	<% if strRompaq <> "" then %> Rompaq&nbsp;&nbsp; <%end if%>		
	<% if strPreinstallROM <> "" then %> Preinstall;&nbsp;&nbsp; <%end if %>
	<% if strCAB <> "" then %> CAB&nbsp;&nbsp; <%end if %>
	</font></td>
	</TABLE></td></tr>

	<%if strISOImage = "checked" or strAR = "checked" then%>
    <tr><TD colspan=3><TABLE width="100%">
    <%else%>
    <tr style="Display:none"><TD colspan=3><TABLE width="100%">
    <%end if%>
	<td nowrap><font face=verdana size=1><b>Replicated&nbsp;By:</b>&nbsp;&nbsp;
    <%=strReplicater%>
	</font></td>
	</TABLE></td></tr>

    <%if left(strFilename,5) <> "HFCN_"  and (trim(lcase(strCDImage))="checked" or trim(lcase(strISOImage))="checked") then%>
	<tr><TD colspan=3><TABLE width="100%">
	<%else%>
	<tr style="Display:none"><TD colspan=3><TABLE width="100%">
	<%end if%>
	<td><font face=verdana size=1><b>CD/DVD Part Number:</b>
	<% if strMultiLanguage = "1" then%>
		<%=strCDPart%></td></font>
	<%else%>
		<%=strPartRow%>	</font></td></TABLE></td></tr>
		<tr><TD colspan=3><TABLE width="100%">
	<%end if%>
	
  
    <%if left(strFilename,5) <> "HFCN_" and trim(lcase(strChkKitNumber)) = "checked" and (trim(lcase(strCDImage))="checked" or trim(lcase(strISOImage))="checked" or trim(lcase(strAR))="checked") then %>
	<td>
	<%else%>
	<td style="Display:none">
	<%end if%>
		<font face=verdana size=1><b>Kit Number:</b>&nbsp;
	<% if strMultiLanguage = "1" then%>
		<%=strKitNumber%>
	<%else%>
		<%=strKitRow%>
	<%end if%>	
	</font></td>
	</TABLE></td></tr>
								
	<%if trim(strTypeID) = "1" then%>
		<tr><TD colspan=3><TABLE width="100%">
		<td nowrap><font face=verdana size=1><b>RoHS/Green&nbsp;Spec:</b>&nbsp;&nbsp;
        <%
        rs.open "spGetRoHSGreenDisplayName " & clng(strRohs) & "," & clng(strGreenSpec),cn
        if rs.eof and rs.bof then
            response.write "Unknown" 
        elseif trim(rs("Rohs") & "") <> "" and trim(rs("GreenSpec") & "") <> "" then
            response.write rs("Rohs") & "_" & rs("GreenSpec") 
        elseif trim(rs("Rohs") & "") <> ""  then
            response.write rs("Rohs") 
        elseif trim(rs("GreenSpec") & "") <> "" then
            response.write rs("GreenSpec") 
        else
            response.write "Unknown" 
        end if

        rs.close
        %>
	</font></td>
	</TABLE></td></tr>
	<%end if%>
	
	<%if request("ID") <> "" then%>
		<%if strAR <> "" then%>
			<TR style="display:none"><TD colspan=3><TABLE width="100%">
		<%else%>
			<TR><TD colspan=3><TABLE width="100%">
		<%end if%>
		<td nowrap><font face=verdana size=1><b>Location/Path:</b>&nbsp;&nbsp;&nbsp; 
			<% if left(strTransfer,2) = "\\" then%>
				<a target=_blank href="<%=strTransfer%>"><%=strTransfer%></a>
			<%else%>
				<%=strTransfer%>
			<%end if%>
	</font></td>
	</TABLE></td></tr>
	<%end if%>    

	<TR><TD colspan=3><TABLE width="100%">
	<td><font face=verdana size=1><b>Comments:</b>&nbsp;&nbsp; 
    	<%=strComments%>&nbsp;&nbsp;
	</font></td>
	</TABLE></td></tr>


	<TR><TD colspan=3><TABLE width="100%">
	<td><font face=verdana size=1><b>Observations fixed in this release:</b>&nbsp;&nbsp;&nbsp;
		<%=strSelectedOTS%>
	</font></td>
	</TABLE></td></tr>

	<TR><TD colspan=3><TABLE width="100%">
	<td><font face=verdana size=1><b>Modifications, Enhancements, or Reason for Release:</b>&nbsp;&nbsp; 
		<BR><%=replace(strChanges,vbcrlf,"<BR>")%>
	</font></td>
	</TABLE></td></tr>


	<TR><TD colspan=3><TABLE width="100%">
	<td nowrap><font face=verdana size=1><b>Supported Operating Systems:</b>&nbsp;&nbsp; 
	<% if strSelectedOS <> "" then %>
		<%=strSelectedOS%>
	<% else %>
		OS Independent&nbsp;&nbsp;
	<% end if %>
	</font></td>
	</TABLE></td></tr>

	<TR><TD colspan=3><TABLE width="100%">
	<% if trim(strTypeID)="1" then %>
		<td><font face=verdana size=1><b>Supported Countries:</b>&nbsp;&nbsp; 
	<% else %>
		<td><font face=verdana size=1><b>Supported Languages:</b>&nbsp;&nbsp; 
	<% end if %>
	<% if strSelectedLangs <> "" then %>
		<%=strSelectedLangs%>
	<% elseif trim(strTypeID)="1" then %>
		Country Independent
	<% else %>
		OS Independent&nbsp;&nbsp;
	<% end if %>
	</font></td>
	</TABLE></td></tr>


	<TR><TD colspan=3><TABLE width="100%">
	<td><font face=verdana size=1><b>PNP Devices dependencies:</b>&nbsp;&nbsp;&nbsp;
	<%
	do while instr(strPNPDevices,vbcrlf) > 0
		Response.Write left(strPNPDevices,instr(strPNPDevices,vbcrlf)-1) & "; "
		strPNPDevices = mid(strPNPDevices,instr(strPNPDevices,vbcrlf)+ 2)
	loop
	if strPNPDevices <> "" then
		Response.Write strPNPDevices
	end if
	%>	
	</font></td>
	</TABLE></td></tr>


	<%
	dim strLoadedDepends
	dim strDependRows
	strLoadedDepends=""
	strDependRows = ""
	if Request("ID") <> "" then
		rs.Open "spGetSelectedDepends " & clng(Request("ID")),cn,adOpenForwardOnly
		do while not rs.eof
			strDependRows= strDependRows &  "1" & chr(1) & rs("ID") & chr(1) & rs("name") & " ["  & rs("version") 
			if trim(rs("revision") & "") <> "" then
				strDependRows= strDependRows &  ","  & rs("revision") 
			end if
			if trim(rs("pass") & "") <> "" then
				strDependRows= strDependRows &  ","  & rs("pass") 
			end if
			strDependRows= strDependRows & "]" & chr(2)
			strLoadedDepends= strLoadedDepends & ","  & rs("ID")
			rs.MoveNext
		loop
		rs.Close
	end if

	if 	strLoadedDepends<> "" then
		strLoadedDepends = strLoadedDepends & ","
	end if
	rs.Open "spGetDepends4Version " & clng(strRootID),cn,adOpenForwardOnly
	do while not rs.eof
		'Response.Write rs("ID") & ":" & strLoadedDepends & "<BR>"
		if instr(strLoadedDepends,"," & rs("ID") & ",")=0 then
			strDependRows= strDependRows &  "0" & chr(1) & rs("ID") & chr(1) & rs("name") & " ["  & rs("version") 
			if trim(rs("revision") & "") <> "" then
				strDependRows= strDependRows &  ","  & rs("revision") 
			end if
			if trim(rs("pass") & "") <> "" then
				strDependRows= strDependRows &  ","  & rs("pass") 
			end if
			strDependRows= strDependRows & "]" & chr(2)
		end if
		rs.MoveNext
	loop
	rs.Close
	%>
	
	<% if strDependRows = "" then %>
		<TR style="Display:none"><TD colspan=3><TABLE width="100%">
		<td><font face=verdana size=1><b>Deliverables Dependencies:</b>&nbsp;&nbsp;
	<% else %>
		<TR><TD colspan=3><TABLE width="100%">
		<td><font face=verdana size=1><b>Deliverables Dependencies:</b>&nbsp;&nbsp;
	<% end if %>

	<%
		dim DependArray
		dim DependElements
		DependArray = split(strDependRows,chr(2))
	%>
	
	<%
	for i = lbound(DependArray) to ubound(DependArray)
		if trim(DependArray(i)) <> "" then
		DependElements = split(DependArray(i),chr(1))
		%>
		
			<% if trim(DependElements(0)) = "1" then%>
				<%=trim(DependElements(2))%>,&nbsp;&nbsp;
			<%end if%>
		<%		
			
		end if
	next
	%>
	
	</font></td>
	</TABLE></td></tr>

	<TR><TD colspan=3><TABLE width="100%">
	<td><font face=verdana size=1><b>Other Dependencies:</b>&nbsp;&nbsp;
	<%=strDependencies%>
	</font></td>
	</TABLE></td></tr>


	<tr><TD colspan=3><TABLE width="100%">
		<td valign="top" nowrap><b><font face=verdana size=1>Products:&nbsp;&nbsp;&nbsp;&nbsp;</b></td>
			<td width=100%>
		<table width="100%" bgcolor="white" border="1" cellpadding="1" cellspacing="1">
			<% if strcategoryID = 2 or strcategoryID = 3 or strcategoryID = 15 or strcategoryID = 33 or strcategoryID = 36 or strcategoryID = 40 or strcategoryID = 71 or strcategoryID = 131 then 'Commodity%>
				<tr bgcolor="gainsboro"><td><b>Product</b></td><td><b>Commodity&nbsp;PM</b></td><td><b>Dev.&nbsp;Approval</b></td><td><b>Test&nbsp;Status</b></td></tr>
			<%else%>
				<tr bgcolor="gainsboro"><td><b>Product</b></td><td><b>SE&nbsp;PM</b></td><td><b>Dev.&nbsp;Approval</b></td><td><b>SE&nbsp;PM&nbsp;Status</b></td><td><b>Preinstall</b></td></tr>
			<%end if%>			
				<%

					dim strProductStatusTable
					dim strProductStatusOutput
					dim strProductStatusPIStatus
					if request("ID") <> "" then
						if trim(strTypeID) = "1" then'strcategoryID = 2 or strcategoryID = 3 or strcategoryID = 15 or strcategoryID = 33 or strcategoryID = 36 or strcategoryID = 40 or strcategoryID = 71 or strcategoryID = 131 then	
							strSQL = "spGetproductStatus4Commodity " & clng(request("ID"))
						else
							strSQL = "spGetproductStatus4Deliverable " & clng(request("ID"))
						end if				

						rs.Open strSQL,cn,adOpenStatic
						strProductStatusTable = ""
						do while not rs.EOF
							if CurrentUserPartner = "1" or CurrentUserPartner = trim(rs("PartnerID")) then
								if rs("DeveloperNotificationStatus") = 1 then
									strDevStatus = "Approved"
								elseif rs("DeveloperNotificationStatus") = 2 then
									strDevStatus = "Rejected"
								elseif rs("DeveloperNotificationStatus") = 0 then
									strDevStatus = "Under Review" 
								else
									strDevStatus = ""
								end if
								'if strcategoryID = 2 or strcategoryID = 3 or strcategoryID = 15 or strcategoryID = 33 or strcategoryID = 36 or strcategoryID = 40 or strcategoryID = 71 or strcategoryID = 131 then 'Commodity
								if trim(strTypeID) = "1" then 'Hardware
									strProductStatusTable = strProductStatusTable & "<TR><TD bgcolor=White nowrap><font face=verdana size=1>" & rs("product")  & "</font></TD>"
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White nowrap><font face=verdana size=1>" & rs("Commoditypm") & "</font></TD>"	
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White nowrap><font face=verdana size=1>" & strDevStatus & "</font></td>"
									if (rs("TestStatus") & "") = "Date" then
										strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & rs("TestDate") & "</font></TD>"							
									else
										if trim(rs("TestStatus") & "") = "" then
											strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>Not&nbsp;Used</font></TD>"							
										else
											strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & rs("TestStatus") & "</font></TD>"							
										end if
									end if
									strProductStatusTable = strProductStatusTable & "</TR>"
								else						
									strProductStatusTable = strProductStatusTable & "<TR><TD bgcolor=White><font face=verdana size=1>" & rs("product")  & "</font></TD>"
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & rs("sepm") & "</font></TD>"							
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & strDevStatus & "</font></TD>"
									if rs("Targeted") then
										strProductStatusOutput = "Targeted"
									elseif not rs("Prereleased") then
										strProductStatusOutput = "Pending"
									elseif rs("PMAlert") then
										strProductStatusOutput = "In Progress"
									else
										strProductStatusOutput = "Available"
									end if
							
									if (not rs("Preinstall")) and rs("Prereleased") then
										strProductStatusPIStatus = "N/A"
									elseif not rs("Preinstall") then
										strProductStatusPIStatus = "TBD"
									elseif (not rs("InImage")) and rs("Targeted") then
										strProductStatusPIStatus = "In Progress"
									elseif rs("InImage") then
										strProductStatusPIStatus = "In Image"
									elseif strProductStatusOutput = "Available" then
										strProductStatusPIStatus = "Available"
									else
										strProductStatusPIStatus = "Pending"
									end if
							
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & strProductStatusOutput & "</font></TD>"
									strProductStatusTable = strProductStatusTable & "<TD bgcolor=White><font face=verdana size=1>" & strProductStatusPIStatus & "&nbsp;</font></TD>"							
									strProductStatusTable = strProductStatusTable & "</TR>"
								end if
							
							end if
							rs.MoveNext
						loop
						rs.Close
						if strProductStatusTable <> "" then
							Response.Write strProductStatusTable
						else
							response.write "<TR><TD colspan=4><font face=verdana size=1>No Product Status Available</font></td></tr>"
						end if
					end if
				%>
		</table>			
		</td>
	</tr>


</Table><BR>
<%
	set rs = nothing
	cn.Close
	set cn = nothing
%>


	
<BR>
<BR>
<BR>
<font size=1 face=verdana>Report Generated <%=Date()%></font>

</font>
</span>
<%end if%>

</BODY>
</HTML>

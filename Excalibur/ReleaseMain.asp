<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="_ScriptLibrary/fqdn.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cboMilestone_onchange() {
	var i;

	Milestone.SelectedMilestoneID.value = Release.cboMilestone.value
	Milestone.submit();
}


function instr(MyString,Find){
	return MyString.substr(0,MyString.indexOf(Find,0));
	
}

function instrAt(MyString,Find){
	return MyString.indexOf(Find,0);
	
}

function midstr(MyString,Find){
	return MyString.substr(MyString.indexOf(Find,0)+1);
	
}


function cmdRelease_onclick() {
	var strLangs = Release.LanguageList.value + ",";
	var strlang;
	while (instrAt(strLangs,",")>=0) 
		{
		strLang = instr(strLangs,",");
		document.all("cbo" + strLang + "Release").selectedIndex = 1;
		strLangs = midstr(strLangs,",");
		strLangs = midstr(strLangs,",");
		}

	cboReleaseLang_onchange();
}

function cmdFail_onclick() {
	var strLangs = Release.LanguageList.value + ",";
	var strlang;
	while (instrAt(strLangs,",")>=0) 
		{
		strLang = instr(strLangs,",");
		document.all("cbo" + strLang + "Release").selectedIndex = 2;
		strLangs = midstr(strLangs,",");
		strLangs = midstr(strLangs,",");
		}
	cboReleaseLang_onchange();
}


function cboReleasePriority_onchange() {
	if (Release.cboReleasePriority.selectedIndex == 0 || Release.cboReleasePriority.selectedIndex == 2)
		ReleaseJustificationSpan.style.display = "";
	else
		ReleaseJustificationSpan.style.display = "none";
}

function cboTransferServer_onchange() {
	if (Release.cboTransferServer.selectedIndex == 0)
		{
		PathRow.style.display = "none";
		Release.txtTransfer.value = "";
		}
	else
		{
		PathRow.style.display = ""
		if (Release.txtTransfer.value == "")
			Release.txtTransfer.value = "\\" + Release.txtFilename.value;
		}
}

/* RC: 2011-06-29 - Included in fqdn.js
function cmdAddServer_onclick() {
	var strNewName;
	var i;
	var blnFound;
	
	strNewName = window.prompt("Enter the fully qualified name of your Transfer Server.\rFORMAT: \\\\SERVERNAME\\SHARENAME", "");
	if (strNewName != null && strNewName != "")
		{
		
		if (instrAt(strNewName,":")>=0)
			{
				window.alert("You may not specify a driver letter as part of the server name.  Please try again.");
				return;
			}
		else if (strNewName.substr(0,2) != "\\\\")
			{
				window.alert("The server name must start with \\\\");
				return;
			}
		else if (strNewName.split(".").length-1 < 3)
			{
				window.alert("You must use the fully qualified name of your Transfer Server.");
				return;
			}		
		
			blnFound = false;
			for (i=0;i<Release.cboTransferServer.length;i++)
				{
					if (String(Release.cboTransferServer.options[i].text).toUpperCase() == String(strNewName).toUpperCase() || String(ConvertServerName(Release.cboTransferServer.options[i].text)).toUpperCase() == String(strNewName).toUpperCase())
						{
							Release.cboTransferServer.selectedIndex = i;
							cboTransferServer_onchange();
							blnFound = true;
							break;						
						}
				}		
		}
	else
		blnFound = true;
		
	if (!blnFound) 
		{
			Release.cboTransferServer.options[Release.cboTransferServer.length] = new Option(strNewName);
			Release.cboTransferServer.selectedIndex = Release.cboTransferServer.length-1;			
			cboTransferServer_onchange();
		}	
}
*/

/* RC: 2011-06-29 - Included in fqdn.js
function ConvertServerName(strName){
	if (strName.substring(0,2) != "\\\\")
		return strName;
	var strEnd="";
	var strServer="";
	var strBuffer = strName.substring(2);
	
	var re = new RegExp("[\\\\.]");
	var m = re.exec(strBuffer);
	if (m != null) 
		{
		strEnd = strBuffer.substring(m.index);
		if (strEnd.indexOf("\\") > -1)
			strEnd = strEnd.substring(strEnd.indexOf("\\"));
			
		strServer = strBuffer.substring(0,m.index);
		return "\\\\" + strServer + strEnd;	
		}
	else
		 return "\\\\" + strBuffer;

}
*/

function cboReleaseAll_onchange() {

	if (Release.cboReleaseAll.selectedIndex==0 && Release.tagISOFilesRequired.value=="1")
		{
		RowISORelease.style.display="";
		Release.txtISOFilesRequired.value="1"
		}
	else
		{
		RowISORelease.style.display="none";
		Release.txtISOFilesRequired.value="0"
		}
	
	if (Release.cboReleaseAll.selectedIndex==1)
		cmdFail_onclick();
	else
		cmdRelease_onclick();
		
		
	
}

function ISOLFSText_onclick() {
	if (Release.chkISOLFSFile.checked)
		Release.chkISOLFSFile.checked = false;
	else
		Release.chkISOLFSFile.checked = true;
}

function ISOLFSText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ISOLFSText_onmouseout() {
	window.event.srcElement.style.cursor = "default";
}

function ISOMD5Text_onclick() {
	if (Release.chkISOMD5File.checked)
		Release.chkISOMD5File.checked = false;
	else
		Release.chkISOMD5File.checked = true;
}

function ISOMD5Text_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ISOMD5Text_onmouseout() {
	window.event.srcElement.style.cursor = "default";
}

function cboReleaseLang_onchange() {
	var blnReleaseSome = false;
	var i;
	
	if (Release.LanguageCount.value == "1") 
		{
		if (Release.cboReleaseLang.selectedIndex==1)
			blnReleaseSome = true;
		}
	else
		{	
		for (i=0;i<Release.cboReleaseLang.length;i++)
			{
			if (Release.cboReleaseLang(i).selectedIndex==1)
				{
					blnReleaseSome = true;
				}
			}
		}
		
	if (blnReleaseSome && Release.tagISOFilesRequired.value=="1")
		{
		RowISORelease.style.display="";
		Release.txtISOFilesRequired.value="1"
		}
	else
		{
		RowISORelease.style.display="none";
		Release.txtISOFilesRequired.value="0"
		}



		
}
function ClearError() {
return true;
}

function TestPath(){
	window.onerror = ClearError;
	window.open ("file://" + Release.cboTransferServer.options[Release.cboTransferServer.selectedIndex].text + Release.txtTransfer.value);
	window.onerror = "";
}

function UploadIso(){
	UploadAddLinks.style.display = "none";
	UploadRemoveLinks.style.display = "none";
	Release.txtISOFilename.value = "";
	UploadISOLinks.style.display = "";
	Release.txtISOFilename.focus();
}

function UploadZip(strFileType) {
	var strID;
	var strPath;
	var strServer;
	strID = window.showModalDialog("PMR/SoftpaqFrame.asp?Title=Upload Deliverable&Page=../common/fileupload.aspx&filter=" + strFileType.toLowerCase(), "", "dialogWidth:600px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		UploadAddLinks.style.display = "none";
		UploadRemoveLinks.style.display = "";
		UploadISOLinks.style.display = "none";
		UploadPath.innerText = strID.substr(strID.lastIndexOf("\\")+1,strID.length);
		strPath = strID.substr(0,strID.lastIndexOf("\\"));
		strServer = strPath.substr(0,strPath.lastIndexOf("\\"));
		strPath = strPath.substr(strPath.lastIndexOf("\\")+1,strPath.length);
		Release.txtTransfer.value="\\" + strPath;
		Release.cboTransferServer.options[Release.cboTransferServer.length] = new Option(strServer,strServer);
		Release.cboTransferServer.selectedIndex = Release.cboTransferServer.length-1;
		}

}

function RemoveUpload(){
    UploadAddLinks.style.display = "";
	UploadRemoveLinks.style.display = "none";
	UploadISOLinks.style.display = "none";
	Release.txtTransfer.value="";
    Release.txtISOFilename.value = "";
    Release.cboTransferServer.selectedIndex=0;
}


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory>
<link href="style/wizard%20style.css" type="text/css" rel="stylesheet">



<form id="Release" method="post" action="saveRelease.asp">
<%

	dim strSQL
	dim rs
	dim cn
	dim MilestoneCount
	dim strMilestones
	dim LastMilestone
	dim LastMilestoneID	
	dim LastMilestoneNotify
	dim SelectedNilestoneID
	dim SelectedMilestoneNotify
	dim strFilename
	dim showFilename
	dim strTransfer
	dim ShowTransfer
	dim strComments
    dim strReplicator
	dim strDeliverable
	dim strVersion
	dim strLangugages
	dim FailName
	dim ReleaseOptions
	dim FailOptions
	dim strLanguageList
	dim ShowButtons
	dim strTransferWithNoServer
	dim ShowTransferPath
	dim strMultiLanguage
	dim strCategoryID
	dim blnRequiresTTS
	dim strISO
	dim strType
	dim blnFilesRequired
	dim blnHFCN
	dim CommercialProducts
	dim ConsumerProducts
	dim strPMEmailList
	dim strDevManagerEmail
	dim strDeveloperEmail
	dim strTesterEmail
	dim strAlsoNotify
	dim strTesterID
	dim strTTS
	dim strWWANFailureConfirmed
	dim strExecutionEngineerID
	dim strExecutionEngineerEmail
	dim strExecutionEngineerName
	dim strDeveloperName
	dim strMD5
	dim CurrentUserEmail
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserPartner
    dim CloneType
	
    CloneType=0
	strTransferWithNoServer = ""
	ShowTransferPath = "none"
	strPMEmailList = ""
	strDevManagerEmail = ""
	strDeveloperEmail = ""
	strDeveloperName = ""
	strTesterID = ""
	strAlsoNotify=""
	blnRequiresTTS = false
	strMD5 = ""
    strReplicator = ""
	
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
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

	CurrentUserID = 0
	CurrentUserEmail = ""
	CurrentUserPartner= 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserEmail = rs("Email")
		currentuserpartner = rs("PartnerID")
	end if
	rs.Close


	cn.execute "spUpdateReleaseStatuses " & clng(Request("ID")) 'Calculate and set the consumer/commercial product count for the deliverable

    rs.open "spGetExecutionEngineer "  & clng(Request("ID")),cn,adOpenForwardOnly
    if rs.eof and rs.bof then
	    strExecutionEngineerID = 0
	    strExecutionEngineerEmail = ""
	    strExecutionEngineerName= ""
    else
	    strExecutionEngineerID = trim(rs("ID") & "")
	    strExecutionEngineerEmail = rs("email") & ""
	    strExecutionEngineerName= rs("name") & ""
    end if
    rs.close
    	
	rs.Open "spGetVersionProperties4Release " & clng(Request("ID")),cn,adOpenForwardOnly
	strFileName = rs("Filename") & ""
	strTransfer = rs("Transfer") & ""
	strComments = rs("Comments") & ""
    strReplicator = rs("ar") & ""
	strDeliverable = rs("Name") & ""
	ConsumerProducts = rs("ConsumerReleaseStatus") & ""
	CommercialProducts = rs("CommercialReleaseStatus") & ""
	strCategoryID = trim(rs("categoryID") & "")
	strMD5 = rs("MD5") & ""
	blnRequiresTTS = rs("RequiresTTS")
	strMultiLanguage = trim(rs("MultiLanguage") & "")
	strWWANFailureConfirmed = trim(rs("WWANFailureConfirmed") & "")
	strTTS = trim(rs("TTS") & "")
	strWorkflowID = rs("WorkflowID") & ""
	strDevManagerEmail = rs("DevManagerEmail") & ""
	strDeveloperEmail = rs("DeveloperEmail") & ""
	strDeveloperName = rs("DeveloperName") & ""
	strAlsoNotify = rs("AlsoNotify") & ""
    CloneType = rs("CloneType") & ""
	strTesterID = rs("TesterID") & ""
	blnHFCN = rs("HFCN")
	if rs("AR") = 1 then
		blnFilesRequired=false
	else
		blnFilesRequired=true
	end if
	strISO = rs("ISOImage") & ""
	strType = trim(rs("TypeID") & "")
	strVersion = rs("Version")
	if trim(rs("Revision")) <> "" then
		strVersion = strVersion & "," & rs("Revision")
	end if
	if trim(rs("Pass")) <> "" then
		strVersion = strVersion & "," & rs("Pass")
	end if
		
	'if trim(strFilename) <> "" then
		ShowFilename = "none"
	'else
	'	ShowFilename = ""
	'end if
	
	rs.Close
	
	if trim(strTesterID) <> "" and trim(strTesterID) <> "0" then
		rs.Open "spGetEmployeeByID " & clng(strTesterID),cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			strTesterEmail = rs("Email") & ""
		end if
		rs.Close
	end if
	
		rs.open "spListServers",cn,adOpenForwardOnly
		strServers = "<OPTION Selected></Option>"
		do while not rs.EOF
			if ucase(left(strTransfer,len(rs("Name")))) = ucase(rs("Name")) then
				strServers = strServers & "<OPTION selected>" & rs("Name") & "</OPTION>"
				strTransferWithNoServer = mid(strTransfer,len(rs("Name"))+1)
				ShowTransferPath = ""
			else
				strServers = strServers & "<OPTION>" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close	
		
		if trim(currentuserpartner) <> "1" or trim(CloneType)="21" then
		    ShowTransferpath = "none"
		end if

		rs.Open "spListCommodityPMs4Version " &  clng(Request("ID")),cn,adOpenForwardOnly
		strPMEmailList = ""
        if rs.State = adStateOpen then		
            do while not rs.EOF
			    strPMEmailList = strPMEmailList & ";" & rs("Email")
			    rs.MoveNext
		    loop
		    rs.Close
		    if strPMEmailList <> "" then
			    strPMEmailList = mid(strPMEmailList,2)
		    end if
        end if
		rs.Open "spListTestLeads4Version " &  clng(Request("ID")),cn,adOpenForwardOnly
		strTestLeadList = ""
		do while not rs.EOF
			strTestLeadList = strTestLeadList & ";" & rs("Email")
			rs.MoveNext
		loop
		rs.Close
		if strTestLeadList <> "" then
			strTestLeadList = mid(strTestLeadList,2)
		end if		
		
		dim FunctionName
	
        if request("Action") = "2" then
			if trim(strType) = "1" then
			    Functionname = "Cancel"
			else
			    Functionname = "Fail"
	        end if
	    else
			Functionname = "Release"
	    end if

%>	
	<font size=3 face=verdana><b><%=Functionname%>&nbsp;<%=strDeliverable & " " & strVersion%></b></font><font size=1><BR><BR></font>
	
<%
	dim MaxMilestone
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		MaxMilestone = 2

	rs.Open "spGetWorkflowStepsInProgress " & clng(Request("ID")),cn,adOpenForwardOnly
		
	MilestoneCount  = 0
	strMilestones = "<Option Selected></option>"
	SelectedMilestoneNotify = ""
	strNotify = ""
	LastMilestone = ""
	LastMilestoneID = ""
	LastMilestoneNotify = ""
	if rs.EOF and rs.BOF then
		Response.Write "<font size=2 face=verdana>There is nothing to release.</font>"
	elseif   lcase(trim(strWWANFailureConfirmed)) = "false" and lcase(trim(strTTS)) = "failed" and blnRequiresTTS then 'and strCategoryID="131"
		Response.write "<font size=2 face=verdana>It can not be released to the next workflow step until it is reviewed by the WWAN Engineers because it failed TTS.</font>"
	else
		do while not rs.EOF
			if rs("ReportMilestone") <= MaxMilestone then
				if request("SelectedMilestoneID") = rs("ID") & "" then
					strMilestones = strMilestones & "<Option selected value=" & rs("ID") & ">" & rs("Milestone") & "</option>"
					SelectedMilestoneID = rs("ID") & ""
					strMilestoneName = rs("Milestone") & ""
					SelectedMilestoneNotify = rs("Notify") & ""
					if clng(ConsumerProducts) > 0  and trim(rs("NotifyConsumerSpecific") & "") <> "" then
						SelectedMilestoneNotify = SelectedMilestoneNotify & ";" & rs("NotifyConsumerSpecific")
					end if
					if clng(CommercialProducts) > 0 and trim(rs("NotifyCommercialSpecific") & "") <> "" then
						SelectedMilestoneNotify = SelectedMilestoneNotify & ";" & rs("NotifyCommercialSpecific")
					end if
				else
					strMilestones = strMilestones & "<Option value=" & rs("ID") & ">" & rs("Milestone") & "</option>"
					LastMilestoneNotify = rs("Notify") & ""
					LastMilestone = rs("Milestone")
					LastMilestoneID = rs("ID")
					if clng(ConsumerProducts) > 0  and trim(rs("NotifyConsumerSpecific") & "") <> "" then
						LastMilestoneNotify = LastMilestoneNotify & ";" & rs("NotifyConsumerSpecific")
					end if
					if clng(CommercialProducts) > 0 and trim(rs("NotifyCommercialSpecific") & "") <> "" then
						LastMilestoneNotify = LastMilestoneNotify & ";" & rs("NotifyCommercialSpecific")
					end if
				end if
				MilestoneCount = MilestoneCount + 1
			end if
			rs.Movenext
		loop
		rs.Close
		
		if MilestoneCount = 0 then
			Response.Write "<font size=2 face=verdana>This screen is only used for workflow transitions until a deliverable is released to the Release Team.</font>"
		else
		
%>		
	<table  WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<TD><b>Deliverable ID:</b></TD>
		<TD><font size=2 face=verdana><%=Request("ID")%></font></TD>
	</TR>
	<tr>
    	<%if request("Action") = "2" then%>
		    <td width="150" nowrap><b>Workflow Step:</b>
		<%else%>
		    <td width="150" nowrap><b>Release From:</b>
        <%end if%>
        
		<%if MilestoneCount > 1 then%>
		 <font color="red" size="1">*</font> 
		 <%end if%>
		 </td>
		<td colspan="10" width="100%">
<%
		if MilestoneCount = 1 then
			SelectedMilestoneID = LastMilestoneID
			strMilestoneName = lastMilestone
			SelectedMilestoneNotify = lastMilestoneNotify
			Response.Write strMilestoneName & "<SELECT style=""Display:none"" id=cboMilestone name=cboMilestone>" & strMilestones & "</SELECT>"
		else
			Response.Write "<SELECT id=cboMilestone name=cboMilestone LANGUAGE=""javascript"" onchange=""return cboMilestone_onchange()"">" & strMilestones & "</SELECT>"
		end if

%>
		</td>
	</tr>
	
<%if trim(SelectedMilestoneID) <> "" then  'Populate all of the release infor for the selected step%>

<%

		'if instr(lcase(strMilestoneName),"test") then
			FailName = "Fail"
		'else
		'	FailName = "Cancel"
		'end if

		if request("Action") = "2" then
			ReleaseOptions = "<OPTION value=0>No Release</OPTION><OPTION value=1>Release</OPTION><OPTION selected value=2>" & FailName & "</OPTION>"
		else
			ReleaseOptions = "<OPTION value=0>No Release</OPTION><OPTION selected value=1>Release</OPTION><OPTION value=2>" & FailName & "</OPTION>"
		end if
		'FailOptions = "<OPTION value=0>No Release</OPTION><OPTION value=1>Release</OPTION><OPTION selected value=2>" & FailName & "</OPTION>"

		rs.Open "spGetNextMilestone " & clng(Request("ID")) & "," & clng(trim(SelectedMilestoneID)),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strNextMilestoneID = 0		
			strNextMilestone = "Workflow Complete"
		else
			strNextMilestoneID = rs("ID")		
			strNextMilestone = rs("Milestone")
		end if
		rs.Close

%>
	<%if request("Action") = "2" then%>
		<TR style="Display:none">
	<%else%>
		<TR>
	<%end if%>

		<td width="150" nowrap><b>Release To:</b></td>
		<%if strNextMilestone = "Workflow Complete" then%>
			<td colspan="10" width="100%" ID=NextStepText><font color=red><b>Product Team (Deliverable Workflow Complete)</b></font></td>
		<%else%>
			<td colspan="10" width="100%" ID=NextStepText><font color=red><b><%=strNextMilestone%></b></font></td>
		<%end if%>
	</TR>
	<%
		'if strDeliverable = "Test Deliverable" then
		'	SelectedMilestoneNotify = ""
		'else
		if left(trim(SelectedMilestoneNotify),1) = ";" then
			SelectedMilestoneNotify = mid(SelectedMilestoneNotify,2)
		end if
		if blnHFCN and (ucase(strNextMilestone) = "WORKFLOW COMPLETE" or ucase(strNextMilestone) = "RELEASE TEAM"  or ucase(trim(strNextMilestone)) = "") then
			if trim(selectedMilestoneNotify) = "" then
				selectedMilestoneNotify = "NBSCSWEngrs@hp.com"
			else
				selectedMilestoneNotify = "NBSCSWEngrs@hp.com;" & selectedMilestoneNotify
			end if
		end if
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[PM]",strPMEmailList)
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[CommodityPM]",strPMEmailList)
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[DevManager]",strDevManagerEmail)
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[Developer]",strDeveloperEmail)
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[TestLeads]",strTestLeadList)
		selectedMilestoneNotify = replace(selectedMilestoneNotify,"[ComponentTestLead]",strTesterEmail)

        if left(selectedMilestoneNotify,1) = ";" then
            selectedMilestoneNotify = mid(selectedMilestoneNotify,2)
        end if

		'if request("Action") = "2" then
		'	SelectedMilestoneNotify = strDeveloperEmail
		'else
		if trim(strTesterEmail) <> "" then 'lcase(strMilestoneName) = "development" and
			if trim(selectedMilestoneNotify) = "" then
				selectedMilestoneNotify = strTesterEmail
			else
				selectedMilestoneNotify = strTesterEmail & ";" & selectedMilestoneNotify
			end if
		end if
		if trim(strExecutionEngineerEmail) <> "" then 
			if trim(selectedMilestoneNotify) = "" then
				selectedMilestoneNotify = strExecutionEngineerEmail
			else
				selectedMilestoneNotify = strExecutionEngineerEmail & ";" & selectedMilestoneNotify
			end if
		end if
		
		if instr(SelectedMilestoneNotify,strDeveloperEmail)=0 then 'Add the developer if they are not in the list already
			if trim(selectedMilestoneNotify) = "" then
				selectedMilestoneNotify = strDeveloperEmail
			else
				selectedMilestoneNotify = strDeveloperEmail & ";" & selectedMilestoneNotify
			end if
		end if

		if trim(strAlsoNotify) <> "" then
			if SelectedMilestoneNotify="" then
				SelectedMilestoneNotify = strAlsoNotify
			elseif right(SelectedMilestoneNotify,1)=";" then
				SelectedMilestoneNotify = SelectedMilestoneNotify & strAlsoNotify
			else
				SelectedMilestoneNotify = SelectedMilestoneNotify & ";" & strAlsoNotify
			end if
		end if
	
    if strMilestoneName = "Core Team" and trim(strType) = "1" and trim(strExecutionEngineerName) <> "" and trim(strExecutionEngineerID) <> "0" and trim(strExecutionEngineerID) <> "" and trim(strDeveloperEmail) <> trim(strExecutionEngineerEmail) then
    	response.write "<TR>"
		response.write "<td width=""150"" nowrap><b>Old Developer:</b></td>"
		response.write "<td colspan=10 width=""100%"">" & strDeveloperName & "</td>"
	    response.write "</TR>"
    	response.write "<TR>"
		response.write "<td width=""150"" nowrap><b>New Developer:</b></td>"
		response.write "<td colspan=10 width=""100%"">" & strExecutionEngineerName & "<INPUT style=""display:none;"" type=""text"" id=txtDeveloperID name=txtDeveloperID value=""" & trim(strExecutionEngineerID) & """><INPUT style=""display:none;"" type=""text"" id=txtDeveloperName name=txtDeveloperName value=""" & trim(strExecutionEngineerName) & """></td>"
	    response.write "</TR>"
        
    end if	
	
	 %>

	
	<TR>
		<td width="150" nowrap><b>Send Email To:</b></td>
		<td colspan="10" width="100%"><INPUT style="width:100%" type="text" id=txtNotify name=txtNotify value="<%=SelectedMilestoneNotify%>"></td>
	</TR>
	<%if trim(strType) = "1" then%>
		<%if request("Action") = "2" then%>
			<TR style="Display:none">
		<%else%>
			<TR>
		<%end if%>
		<td valign=top width="200" nowrap><b>Location:</b> 
			</td>
		<TD>
			<font size=1 color=green>Enter path to files or tell how to get hardware update.<BR></font>
			<INPUT type="text" id=txtHWLocation name=txtHWLocation style="WIDTH:100%" maxlength=255 value="<%=server.htmlencode(strTransfer)%>">
		</TD>
		</TR>
		<tr style="Display:none">
	<%else%>
		<INPUT type="text" id=txtHWLocation name=txtHWLocation style="WIDTH:100%;display:none">
		<%if blnFilesRequired and trim(currentuserpartner) = "1" and  trim(CloneType)<>"21" then%>
			<tr>
		<%else%>
			<tr style="Display:none">
		<%end if%>
	<%end if%>
		<td valign=top width="200" nowrap><b>Transfer Server:&nbsp;<font color="#ff0000" size="1">*</font></b> 
			<font color="#ff0000" size="1" ID=RequireServer style="Display:none">*</font>
			</td>
			<td>
				<SELECT style="width:400" id=cboTransferServer name=cboTransferServer LANGUAGE=javascript onchange="return cboTransferServer_onchange()">
					<%=strServers%>
				</SELECT>&nbsp;<INPUT type="button" value="Add" id=cmdAddServer name=cmdAddServer LANGUAGE=javascript onclick="return FQDNPrompt(false, Release.cboTransferServer, cboTransferServer_onchange)">
			</TD>
		</TR>
	<TR ID=PathRow style="Display:<%=ShowTransferPath%>">
		<td valign=top width="200" nowrap><b>Transfer&nbsp;Path:&nbsp;<font color="#ff0000" size="1">*</font></b> 
			<font color="#ff0000" size="1"  ID=RequirePath style="Display:none">*</font>
		</td>
		<TD>
			<input style="width:100%" type="text" id="txtTransfer" name="txtTransfer" value="<%=strTransferWithNoServer%>">
			<a href="javascript:TestPath();">Verify This Path is Correct</a>&nbsp;<font color=green face=verdana size=1><BR>Note: This may take a while if you are not logged into an NT domain.</font>		
		</TD>
	</TR>
	
    <%if (trim(currentuserpartner) <> "1" or trim(CloneType)="21") and trim(strType) <> "1" and trim(strTransfer) = "" then%>
        <tr>
            <td valign=top><b>
            <%if trim(CloneType) = "21" then %>
                Upload CVA File:
                            
            <%
                strUploadFileType = "CVA"
            else
                strUploadFileType = "Zip"
            %>
                Upload Deliverable:
            <%end if%>
            </b>&nbsp;<font color="#ff0000" size="1">*</font></td>
            <td>
                <div id=UploadAddLinks><a href="javascript: UploadZip('<%=strUploadFileType%>');">Upload <%=strUploadFileType%> File</a>&nbsp;
                <%if strUploadFileType <> "CVA" then %>
                |&nbsp;<a href="javascript: UploadIso();">Enter FTP Path</a>&nbsp;&nbsp;
                <%end if%>
                </div>
                <div id=UploadRemoveLinks style="display:none"><a href="javascript: UploadZip('<%=strUploadFileType%>');">Change</a> | <a href="javascript: RemoveUpload();">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath></label></div>
                <div id=UploadISOLinks style="display:none"><table width=100% cellpadding=0 cellspacing=0><tr><td valign=top><a href="javascript: RemoveUpload();">Cancel</a>&nbsp;&nbsp;</td><td><font size=1 face=verdana color=green>Example: ftp://www.myserver.com/myFolder</font></td></tr><tr><td width=10 valign=top nowrap><b>FTP&nbsp;Path:</b>&nbsp;</td><td width="100%"><input id="txtISOFilename" name="txtISOFilename" style="width:100%" type=text value=""></td></tr></table></div>
            </td>
        </tr>
    <%elseif trim(currentuserpartner) <> "1" and trim(strType) <> "1" then %>
        <tr>
            <td valign=top><b>Deliverable&nbsp;Path:</b>&nbsp;<font color="#ff0000" size="1">*</font></td>
            <td><%=server.HTMLEncode(strTransfer)%><input id="txtISOFilename" name="txtISOFilename" style="width:100%" type=hidden value="<%=server.htmlencode(trim(strTransfer))%>"></td>
            
        </tr>
    <%else%>
        <tr style=display:none>
            <td colspan=2><input id="txtISOFilename" name="txtISOFilename" style="width:100%" type=hidden value=""></td>
        </tr>
    <%end if%>	
    
<tr id="RowFilename" style="Display:<%=ShowFilename%>">
		<td nowrap><b>Filename:</b> <font color="#ff0000" size="1">*</font></td>
		<td><input id="txtFilename" name="txtFilename" style="WIDTH: 100px; HEIGHT: 22px" size="11" maxlength="20" value="<%=strFilename%>"><font color="blue" size="1"> (i.e., MARGI_A1.112, 6R_0324.00, etc)</font></td>
	</tr>

    <%if  trim(strType) <> "1" then%>
		<%if request("Action") = "2" or trim(CloneType)="21" then%>
			<TR style="Display:none">
		<%else%>
			<TR>
		<%end if%>
            <td nowrap><b>MD5:</b></td>
            <td colspan="10"><input id="txtMD5" maxlength="50" name="txtMD5" style="width: 250px;height: 22px" size="55" value="<%=trim(strMD5)%>"></td>
        </tr>
    <%end if%>


	<%if lcase(strNextMilestone) = "release team" then%>
		<tr id="RowReleasePriority" style="Display:">
	<%else%>
		<tr id="RowReleasePriority" style="Display:none">
	<%end if%>
		<td width="200" nowrap><b>Release Priority:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<SELECT style="width:150" id=cboReleasePriority name=cboReleasePriority LANGUAGE=javascript onchange="return cboReleasePriority_onchange()">
				<OPTION Value=1>High</OPTION>
				<OPTION Value=2 selected>Normal</OPTION>
				<OPTION Value=3>After-hours Support</OPTION>
			</SELECT>
			<span style="Display:none" ID=ReleaseJustificationSpan>Justification:&nbsp;<INPUT style="Width:190" type="text" maxlength=80 id=txtReleasePriorityJust name=txtReleasePriorityJust></span>
		</td>
	</tr>

  <%
	rs.Open "spGetLanguageStatuses4Milestone " & clng(Request("ID")) & "," & clng(SelectedMilestoneID),cn,adOpenForwardOnly
	strLanguages = ""
	strLanguageList = ""
	dim LanguageCount
	LanguageCount = 0
	do while not rs.EOF
		if strType="1" then
			if rs("Abbreviation") & "" = "XX" then
				strAbbreviation = "XX"
			else
				strAbbreviation = rs("Language") & ""
			end if
		else
			strAbbreviation = rs("Abbreviation") & ""
		end if
		
		if strAbbreviation = "XX" then
			if rs("Failed") then
				strLanguages = strLanguages & "<SELECT id=""cboReleaseLang"" LANGUAGE=javascript onchange=""return cboReleaseLang_onchange()"" style=""width:100"" name=""cbo" & strAbbreviation & "Release"">" & ReleaseOptions & "</SELECT>&nbsp;" & replace(replace(rs("Language"),">",""),"<","") & "&nbsp;(" & FailName & "ed)<BR>"
			else
				strLanguages = strLanguages & "<SELECT id=""cboReleaseLang"" LANGUAGE=javascript onchange=""return cboReleaseLang_onchange()"" style=""width:100"" name=""cbo" & strAbbreviation & "Release"">" & ReleaseOptions & "</SELECT>&nbsp;" & replace(replace(rs("Language"),">",""),"<","") & "<BR>"
			end if
		else
			if rs("Failed") then
				strLanguages = strLanguages & "<SELECT id=""cboReleaseLang"" LANGUAGE=javascript onchange=""return cboReleaseLang_onchange()"" style=""width:100""  name=""cbo" & strAbbreviation & "Release"">" & ReleaseOptions & "</SELECT>&nbsp;" & rs("Abbreviation") & " - " & replace(replace(rs("Language"),">",""),"<","") & "&nbsp;(" & FailName & "ed)<BR>"
			else
				strLanguages = strLanguages & "<SELECT id=""cboReleaseLang"" LANGUAGE=javascript onchange=""return cboReleaseLang_onchange()"" style=""width:100""  name=""cbo" & strAbbreviation & "Release"">" & ReleaseOptions & "</SELECT>&nbsp;" & rs("Abbreviation") & " - " & replace(replace(rs("Language"),">",""),"<","") & "<BR>"
			end if
		end if
		LanguageCount = LanguageCount + 1
		strLanguageList = strLanguageList & "," & strAbbreviation & "," & rs("ID")
		rs.MoveNext
	loop  
	rs.Close
	
	if strLanguageList <> "" then
		strLanguageList = mid(strLanguageList,2)
	end if
	if languageCount < 2 then
		ShowButtons = "none"
	else
		ShowButtons = ""
	end if

  %>
  
  <% if strMultiLanguage = "1" then%>
	<tr>
		<td valign=top nowrap><b>Status:</b></td>
		<TD width=100%>
		<%if request("Action") = "2" then%>
			<%if trim(strType) = "1" then%>
				Version Cancelled
			<%else%>
				<font color=red><b>Fail Version</b></font>
			<%end if%>
			<SELECT style="display:none;width:100" id=cboReleaseAll name=cboReleaseAll LANGUAGE=javascript onchange="return cboReleaseAll_onchange()">
				<OPTION selected>Fail</OPTION>
			</SELECT>
		<%elseif request("Action") = "1" then%>
			Release Version
			<SELECT style="display:none;width:100" id=cboReleaseAll name=cboReleaseAll LANGUAGE=javascript onchange="return cboReleaseAll_onchange()">
				<OPTION selected>Release</OPTION>
			</SELECT>
		<%else%>
			<SELECT style="width:100" id=cboReleaseAll name=cboReleaseAll LANGUAGE=javascript onchange="return cboReleaseAll_onchange()">
				<OPTION selected>Release</OPTION>
				<OPTION>Fail</OPTION>
			</SELECT>
		<%end if%>
			</td></tr>
	<tr id="RowLanguages" style="Display:none">
  <% else%>
	<tr id="RowLanguages">
  <%end if%>
		<td valign=top nowrap><b>Languages To Release:</b></td>
		<TD width=100%>
		<Table border=0><TR><TD nowrap valign=top >
			<%=strLanguages%>		
		</TD>
		<TD valign=top nowrap>&nbsp;&nbsp;
		<INPUT style="Display:<%=ShowButtons%>" type="button" value="Release All" id=cmdRelease name=cmdRelease LANGUAGE=javascript onclick="return cmdRelease_onclick()">
		<BR>&nbsp;&nbsp;&nbsp;<INPUT style="Display:<%=ShowButtons%>" type="button" value="     <%=FailName%> All     " id=cmdFail name=cmdFail LANGUAGE=javascript onclick="return cmdFail_onclick()">
		</TD></TR></TABLE>
		</td>
	</tr>
<%if strMilestoneName = "Development" and strISO = "1" then%>
	<INPUT type="hidden" id=txtISOFilesRequired name=txtISOFilesRequired value="1">
	<INPUT type="hidden" id=tagISOFilesRequired name=tagISOFilesRequired value="1">
	<tr id="RowISORelease">
		<td  valign=top width="200" nowrap valign="top"><b>ISO Files:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<font size=1 face=verdana color=red>The release team will not accept this release without these files.<BR></font>
			<INPUT type="checkbox" id=chkISOMD5File name=chkISOMD5File>&nbsp;<label ID=ISOMD5Text LANGUAGE=javascript onclick="return ISOMD5Text_onclick()" onmouseover="return ISOMD5Text_onmouseover()" onmouseout="return ISOMD5Text_onmouseout()">The MD5 file is supplied for each ISO image.</label><BR>
			<INPUT type="checkbox" id=chkISOLFSFile name=chkISOLFSFile>&nbsp;<label ID=ISOLFSText LANGUAGE=javascript onclick="return ISOLFSText_onclick()" onmouseover="return ISOLFSText_onmouseover()" onmouseout="return ISOLFSText_onmouseout()">The LFS file is supplied for each ISO image.</label>
		</td>
	</tr>
<%else%>
	<INPUT type="hidden" id=txtISOFilesRequired name=txtISOFilesRequired value="0">
	<INPUT type="hidden" id=tagISOFilesRequired name=tagISOFilesRequired value="0">
	<Label ID="RowISORelease"></label>
<%end if%>

<% if strType = 1 And cbool(blnHFCN) and lcase(strMilestoneName) = "development" then %>	
    <tr>
        <td valign=top nowrap><b>Test Recommendations:</b></td>
        <td><TEXTAREA style="WIDTH:100%;Height=75" wrap rows=2 cols=20 id="txtTestRecommendations" name="txtTestRecommendations"></TEXTAREA>
        </td>
    </tr>    
    <tr>
        <td valign=top nowrap><b>Sample Notes:</b></td>
        <td><TEXTAREA style="WIDTH:100%;Height=75" wrap rows=2 cols=20 id="txtSampleNotes" name="txtSampleNotes"></TEXTAREA>
        </td>
    </tr>    
<% end if %>

    <tr id="RowComments">
		<td valign=top nowrap><b>Comments:</b></td>
		<td><b><%=strMilestoneName & " Comments "%></b> <font size=1 color=green>(Appended to Previous Comments)<BR></font><TEXTAREA style="WIDTH:100%;Height=75" wrap rows=2 cols=20 id=txtComments name=txtComments></TEXTAREA><BR>
		<b><%="Previous Comments"%></b><BR><TEXTAREA readonly style="background-color:cornsilk;WIDTH:100%;Height=150" rows=2 cols=20 wrap=soft id=txtOldComments name=txtOldComments><%=strComments%></TEXTAREA>
		</td>
	</tr>
	
<%end if%>
	
	</table>		
				
<%		

			set rs = nothing
			cn.Close
			set cn = nothing
		end if 'Valid Milstones are left
	end if 'There are milestones to release for this version
%>
<INPUT type="hidden" id=SelectedMilestone name=SelectedMilestone value="<%=SelectedMilestoneID%>">
<INPUT type="hidden" id=SelectedMilestoneName name=SelectedMilestoneName value="<%=ucase(strMilestoneName)%>">
<INPUT type="hidden" id=NextMilestoneName name=NextMilestoneName value="<%=ucase(strNextMilestone)%>">
<INPUT type="hidden" id=NextMilestoneID name=NextMilestoneID value="<%=strNextMilestoneID%>">
<INPUT type="hidden" id=LanguageList name=LanguageList value="<%=strLanguageList%>">
<INPUT type="hidden" id=LanguageCount name=LanguageCount value="<%=LanguageCount%>">
<INPUT type="hidden" id=DeliverableID name=DeliverableID value="<%=Request("ID")%>">
<input type="hidden" id="CloneType" name="CloneType" value="<%=CloneType%>">
<INPUT type="hidden" id=DeliverableName name=DeliverableName value="<%=strDeliverable%>">
<INPUT type="hidden" id=txtWorkflowID name=txtWorkflowID value="<%=strWorkflowID%>">
<INPUT type="hidden" id=txtType name=txtType value="<%=strType%>">
<INPUT type="hidden" id=txtHFCN name=txtHFCN value="<%=lcase(blnHFCN)%>">
<INPUT type="hidden" id=txtFilesRequired name=txtFilesRequired value="<%=lcase(blnFilesRequired)%>">
<INPUT type="hidden" id=txtUserPartner name=txtUserPartner value="<%=trim(CurrentUserPartner)%>">
<INPUT type="hidden" id=txtReplicator name=txtReplicator value="<%=strReplicator%>">
</form>

<form id="Milestone" method="post" action="ReleaseMain.asp">
<INPUT type="hidden" id=SelectedMilestoneID name=SelectedMilestoneID value="<%=SelectedMilestoneID%>">
<INPUT type="hidden" id=ID name=ID value="<%=Request("ID")%>">
</form>
<%
    if request("Action") <> "2" and lcase(trim(strMilestoneName)) = "functional test" and lcase(trim(strNextMilestone)) = "release team" and ( lcase(trim(CurrentUserEmail)) = lcase(trim(strDevManagerEmail)) or lcase(trim(CurrentUserEmail)) = lcase(trim(strDeveloperEmail))) then
        response.Write "<INPUT type=""hidden"" id=txtWarnDeveloper name=txtWarnDeveloper value=1>"
    else
        response.Write "<INPUT type=""hidden"" id=txtWarnDeveloper name=txtWarnDeveloper value=0>"
    end if
%>

</BODY>
</HTML>

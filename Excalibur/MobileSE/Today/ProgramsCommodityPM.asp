
<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<head>

  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {

	ProgramInput.txtPDE.value= ProgramInput.tagSTLPath.value;
	ProgramInput.txtProgramMatrixPath.value= ProgramInput.tagProgramMatrixPath.value;
	LoadingRow.style.display="none";
	ButtonRow.style.display="";
	CurrentState =  "General";
	ProcessState();
	self.focus();

}

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		//for (i=0;i<event.srcElement.length;i++)
		for (i=event.srcElement.length-1;i>=0;i--)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}


function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }

function mouseover_Column(){
	event.srcElement.style.color="red";
	event.srcElement.style.cursor="hand";
	
}
function mouseout_Column(){
	event.srcElement.style.color="black";
}


function cmdDate_onclick(strField) {
	var strID;
	var i;
	
	strID = window.showModalDialog("../../Mobilese/today/calDraw1.asp",window.document.all(strField).value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			window.document.all(strField).value = strID;
		}
}

//function cmdSETestLeadAdd_onclick() {
//	ChooseEmployee (ProgramInput.cboSETestLead);
//}

function cmdWWANTestLeadAdd_onclick() {
	ChooseEmployee (ProgramInput.cboWWANTestLead);
}

function cmdODMTestLeadAdd_onclick() {
	ChooseEmployee (ProgramInput.cboODMTestLead);
}

function cmdPDEAdd_onclick() {
	ChooseEmployee (ProgramInput.cboPDE);
}

function ChooseEmployee(myControl){
	var ResultArray;

	if (txtProductPartnerID.value == "")
		ResultArray = window.showModalDialog("ChooseEmployee.asp","","dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	else
		ResultArray = window.showModalDialog("ChooseEmployee.asp?PartnerID=" + txtProductPartnerID.value,"","dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	
	if (typeof(ResultArray) != "undefined")
		{
		if (ResultArray[0] != 0)
			{
			myControl.options[myControl.length] = new Option(ResultArray[1],ResultArray[0]);
			myControl.selectedIndex = myControl.length-1;
			}
		}
}
//-->
</SCRIPT>
</head>
<link rel="stylesheet" type="text/css" href="../../style/wizard%20style.css">
<body>
<font face=verdana>

<%
	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
		
	end function	


	dim cn
	dim rs
	dim cm
	dim p
	dim CnString

	dim CheckCommodities

	dim strChecked
	dim strBrandsLoaded
	dim strReleasesLoaded
	dim strPDDPath
	dim strSCMPath
	dim strSTLPath
	dim strProgramMatrixPath
	dim strReferenceList
	dim strReferenceID
	dim strDevCenter
	dim strRegulatoryModel
	dim strHWStatus
	dim strHWStatusDisplay
	dim strProductName
	dim	strSETestLead 
	dim strSETestLeadName
	dim strODMTestLead 
	dim strWWANTestLead 
	dim strPDEList
	'dim strSETestLeadList
	dim strODMTestLeadList
	dim strWWANTestLeadList
	dim blnCommodityPM
	dim blnWWANTestLead
	dim blnODMTestLead
	dim strWWAN
	dim strServicePMID

	blnCommodityPM=false
	blnWWANTestLead=false
	blnODMTestLead=false
	strBrandsLoaded = ""
	strReleasesLoaded = ""
	strProductName = "Product"
	strSETestLead  = ""
	strSETestLeadName  = ""
	strODMTestLead  = ""
	strWWANTestLead  = ""
	'strSETestLeadList = ""
	strODMTestLeadList = ""
	strWWANTestLeadList = ""
	strWWAN = ""
	
	cnString =Session("PDPIMS_ConnectionString")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


' get current user info. Using user's phone number. TDC user's phone number should not start with +1 or 01., the user is not from TDC
' If current user is TDC, display both Program office PM and configurationPM
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPhone
	dim CurrentUserID
	dim blnServicePM
	dim CurrentUserEmail
	dim strCommodityPMName
	strCommodityPMName=""
	
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
	Set	rs = cm.Execute 
	
	set cm=nothing	
		
	if not (rs.EOF and rs.BOF) then
		CurrentUserEmail = rs("Email") & ""
		CurrentUserPhone = rs("Phone") & ""
		CurrentUserID = rs("ID") & ""
		blnServicePM = rs("ServicePM")
		blnCommodityPM = rs("CommodityPM")
	end if
	rs.Close

		CheckCommodities = "checked"
		'Response.Write "<font color=red size=1 face=verdana><b>Please contact Dave Whorton before adding products with PRS, PAV, ENT, TAB, WKS, or SMB in the Version field.</b></font>"

	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = clng(request("ID"))
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spGetProductVersion " & request("ID"),cn,adOpenForwardOnly

		strProductName= rs("DotsName") & ""

		if rs("CommodityLock") then
			strHWStatusDisplay = "none"
			strHWStatus = "checked"
		else
			strHWStatusDisplay = ""
			strHWStatus = ""
		end if

		if rs("OnCommodityMatrix") then	
			CheckHW = "checked"
		else
			CheckHW = ""
		end if

		if rs("WWANProduct") & "" = "" then	
			strWWAN = ""
		else
			strWWAN = abs(clng(rs("WWANProduct")))
		end if

		strPDE = rs("PDEID") & ""
		strServicePMID = rs("ServiceID") & ""
		strSETestLead = rs("SETestLead") & ""
		strODMTestLead = rs("ODMTestLeadID") & ""
		strWWANTestLead = rs("WWANTestLeadID") & ""
		
		strProductPartnerID = rs("Partnerid") & ""
		strServiceDate = rs("ServiceLifeDate") & ""
		strProductStatus = trim(rs("ProductStatusID"))
		if strProductStatus <> "5" then
			strProductActive = "checked"
		else
			strProductActive = ""
		end if 
		rs.Close

		Response.write "<H3>" & strProductName & " Properties</H3>" 


	strPDEList = ""
	
	dim strCommodityPMs
	strHWPMs = ""
	rs.Open "spListCommodityPMsAll 2",cn,adOpenForwardOnly
	do while not rs.EOF
		strHWPMs = strHWPMs & "," & trim(rs("ID"))
		rs.MoveNext
	loop
	rs.Close	

    dim strServicePMs
	strServicePMs = ""
	rs.Open "spListCommodityPMsAll 3",cn,adOpenForwardOnly
	do while not rs.EOF
		strServicePMs = strServicePMs & "," & trim(rs("ID"))
		rs.MoveNext
	loop
	rs.Close	

	rs.Open "spGetEmployees",cn,adOpenForwardOnly
	
	do while not rs.EOF
		if trim(strPDE) = trim(rs("ID")) then
			strPDEList = strPDEList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			strCommodityPMName = rs("Name") & ""
		elseif  instr(strHWPMs & ",","," & trim(rs("ID")) & "," ) > 0  then 'trim(rs("ID")) ="774" or trim(rs("ID")) ="446" or trim(rs("ID")) ="799" or trim(rs("ID")) ="1385" or trim(rs("ID")) ="807" or trim(rs("ID")) ="1637" or trim(rs("ID")) ="1161" or trim(rs("ID"))="1645" then
			if trim(rs("PartnerID") & "") = 1 or ( trim(rs("PartnerID") & "") =  trim(strProductPartnerID) )  then
				strPDEList = strPDEList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
        if trim(strServicePMID) = trim(rs("ID")) then
            strServicePMList = strServicePMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		elseif  instr(strServicePMs & ",","," & trim(rs("ID")) & "," ) > 0  then 
			if trim(rs("PartnerID") & "") = 1 or ( trim(rs("PartnerID") & "") =  trim(strProductPartnerID) )  then
				strServicePMList = strServicePMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
        end if
		rs.MoveNext
	loop
	rs.Close
	

	'strSETestLeadList = ""
	strODMTestLeadList = ""
	strWWANTestLeadList = ""
	if trim(strWWANTestLead) = "" or trim(strWWANTestLead) = "0" then
	    strWWANTestLead = 4511
	    strWWANTestLeadList = "<option selected value=""4511"">Cheng, Steven</option>"
	end if
	

	rs.Open "spGetTestLeadsAll",cn,adOpenStatic
	do while not rs.EOF
		if rs("Role") = "SE Test Lead" then
			if trim(strSETestLead) = trim(rs("ID")) then
				strSETestLeadName = rs("Name") & ""
			end if
		end if

		if rs("Role") = "ODM Test Lead" then
			if trim(strODMTestLead) = trim(rs("ID")) then
				strODMTestLeadList = strODMTestLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				strODMTestLeadName = rs("Name") & ""
			elseif trim(rs("PartnerID") & "") = 1 or ( trim(rs("PartnerID") & "") =  trim(strProductPartnerID) )  then
				strODMTestLeadList = strODMTestLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			if trim(currentuserid) = trim(rs("ID")) then
				blnODMTestLead=true
			end if
		end if

		if rs("Role") = "WWAN Test Lead" then
			if trim(strWWANTestLead) = trim(rs("ID")) then
				strWWANTestLeadList = strWWANTestLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				strWWANTestLeadName = rs("Name") & ""
			elseif trim(rs("PartnerID") & "") = 1 or ( trim(rs("PartnerID") & "") =  trim(strProductPartnerID) )  then
				strWWANTestLeadList = strWWANTestLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			if trim(currentuserid) = trim(rs("ID")) then
				blnWWANTestLead=true
			end if
		end if
		
		rs.MoveNext
	loop
	rs.Close
	
	

set cm=nothing
set rs=nothing
set cn=nothing

if trim(strSETestLeadName)="" then
	strSETestLeadName = "Not Assigned Yet"
else
	strSETestLeadName = longname(strSETestLeadName)
end if

if trim(strODMTestLeadName)="" then
	strODMTestLeadName = "Not Assigned Yet"
else
	strODMTestLeadName = longname(strODMTestLeadName)
end if

if trim(strWWANTestLeadName)="" then
	strWWANTestLeadName = "Not Assigned Yet"
else
	strWWANTestLeadName = longname(strWWANTestLeadName)
end if
	
if trim(strCommodityPMName)="" then
	strCommodityPMName = "Not Assigned Yet"
else
	strCommodityPMName = longname(strCommodityPMName)
end if
	


dim DisplayCommodityEdit
dim DisplayCommodityReadonly
dim DisplayWWANEdit
dim DisplayWWANReadonly
dim DisplayODMEdit
dim DisplayODMReadonly
dim DisplayCommodityPMRow
dim DisplayWWANRow
dim DisplayODMRow
dim DisplaySERow

DisplayCommodityPMRow = ""
DisplayWWANRow = ""
DisplayODMRow = ""
DisplaySERow = ""

if blnCommodityPM then
	DisplayCommodityEdit = ""
	DisplayCommodityReadonly = "none"
	DisplayWWANEdit = ""
	DisplayWWANReadonly = "none"
	DisplayODMEdit = ""
	DisplayODMReadonly = "none"
elseif blnServicePM then
    DisplayCommodityPMRow = "none"
    DisplayWWANRow = "none"
    DisplayODMRow = "none"
    DisplaySERow = "none"
else
	DisplayCommodityEdit = "none"
	DisplayCommodityReadonly = ""
	DisplayWWANEdit = "none"
	DisplayWWANReadonly = ""
	DisplayODMEdit = "none"
	DisplayODMReadonly = ""
end if


if blnWWANTestLead then
	DisplayWWANEdit = ""
	DisplayWWANReadonly = "none"
end if
if blnODMTestLead then
	DisplayODMEdit = ""
	DisplayODMReadonly = "none"
end if
	
	



%>

<FORM ACTION="ProgramSaveCommodityPM.asp" METHOD="post" id="ProgramInput">

<table ID=tabHW border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>

   <tr style=display:<%=DisplayCommodityPMRow%>>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Commodity PM:</font></strong></td>
    <td style=display:<%=DisplayCommodityEdit%>><SELECT id=cboPDE name=cboPDE style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strPDEList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdPDEAdd name=cmdPDEAdd LANGUAGE=javascript onclick="return cmdPDEAdd_onclick()">
	</td>
    <td style=display:<%=DisplayCommodityReadonly%>><%=strCommodityPMName%></TD>	
	</tr>
   <tr style=display:<%=DisplayODMRow%>>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>ODM&nbsp;HW&nbsp;Test&nbsp;Lead:&nbsp;</font></strong></td>
    <td style=display:<%=DisplayODMEdit%>>
		<SELECT id=cboODMTestLead name=cboODMTestLead style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strODMTestLeadList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdODMTestLeadAdd name=cmdODMTestLeadAdd LANGUAGE=javascript onclick="return cmdODMTestLeadAdd_onclick()">
	</td>
    <td style=display:<%=DisplayODMReadonly%>><%=strODMTestLeadName%></TD>	
	</tr>
   <tr style=display:<%=DisplayWWANRow%>>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>WWAN&nbsp;Test&nbsp;Lead:&nbsp;</font></strong></td>
    <td style=display:<%=DisplayWWANEdit%>><SELECT id=cboWWANTestLead name=cboWWANTestLead style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strWWANTestLeadList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdWWANTestLeadAdd name=cmdWWANTestLeadAdd LANGUAGE=javascript onclick="return cmdWWANTestLeadAdd_onclick()">
	</td>
	<td style=display:<%=DisplayWWANReadonly%>><%=strWWANTestLeadName%></TD>	
	</tr>
   <tr style=display:<%=DisplaySERow%>>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>SE&nbsp;Test&nbsp;Lead:&nbsp;</font></strong></td>
    <td><%=strSETestLeadName & "&nbsp;"%>			
	</td></tr>
	
	
	<%if not blnCommodityPM then%>
   <tr style="display:none">
   <%else%>
   <tr>
   <%end if%>
		<td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>WWAN&nbsp;Product:</font><font size=1 color=red face=verdana>&nbsp;*</font></strong></td>
		<td>
		<SELECT id=cboWWAN name=cboWWAN>
			<OPTION selected value=""></OPTION>
			<%if strWWAN = "1" then%>
				<OPTION selected value="1">Yes</OPTION>
			<%else%>
				<OPTION value="1">Yes</OPTION>
			<%end if%>
			<%if strWWAN = "0" then%>
				<OPTION selected value="0">No</OPTION>
			<%else%>
				<OPTION value="0">No</OPTION>
			<%end if%>
		</SELECT>
		</td>
	</tr>

	<%if not blnCommodityPM then%>
   <tr style="display:none">
   <%else%>
   <tr>
   <%end if%>
		<td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Matrix:</font></strong></td>
		<td>
		<INPUT type="checkbox" <%=CheckHW%> id=chkCommodities name=chkCommodities>&nbsp;<font face=verdana size=2>Include this product on the Hardware Matrix.</font>
		</td>
	</tr>

	<%if not blnCommodityPM then%>
   <tr style="display:none">
   <%else%>
   <tr>
   <%end if%>

		<td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Status:</font></strong></td>
		<td>
		<INPUT style="Display:<%=strHWStatusDisplay%>" type="checkbox" id=chkCommodityLock name=chkCommodityLock <%=strHWStatus%>> Locked
		</td>
	</tr>
	<%if blnServicePM then%>
   <tr>
   <%else%>
   <tr style="display:none">
   <%end if%>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Service&nbsp;PM&nbsp;</font></strong></td>
    <td><SELECT id=cboServicePM name=cboServicePM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strServicePMList%>			
		</SELECT>&nbsp;<INPUT type="button" value="Add" id=cmdServcePMAdd name=cmdServiceAdd LANGUAGE=javascript onclick="return cmdServicePMAdd_onclick()">
	</td>
	</tr>

	<%if blnServicePM then%>
   <tr>
   <%else%>
   <tr style="display:none">
   <%end if%>
		<td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>End&nbsp;of&nbsp;Service&nbsp;Life:</font></strong></td>
		<TD>
		<INPUT type="text" id=txtServiceEndDate name=txtServiceEndDate value="<%=strServiceDate%>">
			<a href="javascript: cmdDate_onclick('txtServiceEndDate')"><img ID="picTarget" SRC="../../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</TD>
	</tr>
	<%if blnServicePM and (strProductStatus = "4" or strProductStatus="5" ) then%>
   <tr>
		<td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Product&nbsp;Status:</font></strong></td>
		<TD><INPUT <%=strProductActive%> type="checkbox" id=chkProductStatus name=chkProductStatus value="1"> Active</TD>
	</tr>
   <%end if%>
</table>    
<%if blnServicePM then
	strServicePM="1"
else
	strServicePM=""
end if%>
<INPUT style="Display:none" type="text" id=txtServicePM name=txtServicePM value="<%=strServicePM%>">
<INPUT style="Display:none" type="text" id=tagProductStatus name=tagProductStatus value="<%=strProductStatus%>">
<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT style="Display:none" type="text" id=txtProductName name=txtProductName value="<%=strProductName%>">
<INPUT style="Display:none" type="text" id=txtCurrentUserEmail name=txtCurrentUserEmail value="<%=CurrentUserEmail%>">
<INPUT style="Display:none" type="text" id=tagODMTestLead name=tagODMTestLead value="<%=strODMTestLead%>">
<INPUT style="Display:none" type="text" id=tagWWANTestLead name=tagWWANTestLead value="<%=strWWANTestLead%>">
<INPUT style="Display:none" type="text" id=tagCommodityPM name=tagCommodityPM value="<%=strPDE%>">
 <INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=Request("app")%>">


</FORM>
<INPUT type="hidden" id=txtProductPartnerID name=txtProductPartnerID value="<%=trim(strProductPartnerID)%>">
</body>
</html>
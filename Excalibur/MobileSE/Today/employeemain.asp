<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=8" />  
<meta http-equiv="Expires" CONTENT="0">
<meta http-equiv="Cache-Control" CONTENT="no-cache">
<meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<script src="../../Scripts/verifyEmailAddress.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function DocAccess(strID){
    var strNTName = "";
    var strFirstName = "";
    var strLastName = "";
    var strEmail = "";
    var strPhone = "";
    var strDomain = "";
    var strGroup = "";
    
    if (AddEmployee.cboDomain.selectedIndex != 1 && AddEmployee.cboDomain.selectedIndex != 2 && AddEmployee.cboDomain.selectedIndex != 3 && AddEmployee.cboDomain.selectedIndex != 4)
        {
        alert("You must select an HP domain to continue.");
        AddEmployee.cboDomain.focus();
        }
    else if (AddEmployee.txtNTName.value == "") 
        {
        alert("You must enter an NT Name to continue.");
        AddEmployee.txtNTName.focus();
        }
    else if (strID==2 && AddEmployee.txtFirstName.value == "") 
        {
        alert("You must enter a First Name to continue.");
        AddEmployee.txtFirstName.focus();
        }
    else if (strID==2 && AddEmployee.txtLastName.value == "") 
        {
        alert("You must enter a Last Name to continue.");
        AddEmployee.txtLastName.focus();
        }
    else if (strID==2 && AddEmployee.txtPhone.value == "") 
        {
        alert("You must enter a Phone Number to continue.");
        AddEmployee.txtPhone.focus();
        }
    else if (strID==2 && AddEmployee.txtEmail.value == "") 
        {
        alert("You must enter an Email Address to continue.");
        AddEmployee.txtEmail.focus();
        }
    else if (strID==2 && AddEmployee.cboGroup.selectedIndex ==0)
        {
        alert("You must select a Group to continue.");
        AddEmployee.cboGroup.focus();
        }
    else
        {
        if (strID == 1)
            strNTName = AddEmployee.cboDomain.options[AddEmployee.cboDomain.selectedIndex].text + "\\" + AddEmployee.txtNTName.value;
        else
            {
            strNTName = AddEmployee.txtNTName.value;
            strFirstName = AddEmployee.txtFirstName.value;
            strLastName = AddEmployee.txtLastName.value;
            strEmail = AddEmployee.txtEmail.value;
            strPhone = AddEmployee.txtPhone.value;
            strDomain = AddEmployee.cboDomain.options[AddEmployee.cboDomain.selectedIndex].text;
            strGroup = AddEmployee.cboGroup.options[AddEmployee.cboGroup.selectedIndex].text;
            }

        if (strID==2)
            window.location.href = "mailto:PDC.CNB-PreInfra@hp.com?Subject=Request Access to Program Document Server&Body=Hello,%0a%0aReason for this Request: FILL IN YOUR REASON FOR MAKING THIS REQUEST HERE%0a%0aPlease grant access to the Program Documents on the tpopsgdev01 server to the following people:%0a%09Name: " + strFirstName + " " + strLastName + "%0a%09Domain: " + strDomain + "%0a%09Login Name: " + strNTName + "%0a%09Job Function: " + strGroup + "%0a%09Phone: " + strPhone + "%0a%09Email: " + strEmail + "%0a%0aThanks.";
        else
            window.location.href = "mailto:pulsar.support@hp.com?Subject=Request Access to Program Document Server&Body=Hello Kathey,%0a%0aPlease grant access to the Program Document Server to the following people:%0a%20%20%20%20%20" + strNTName + "%0a%0aThanks.";
        }
}

function cmdCancel_onclick() {
    CloseIframeDialog();
}

function CloseIframeDialog() {
	if (window.parent.document.getElementById('modal_dialog')) {
		window.parent.jQuery('#modal_dialog').dialog('close');
	}
	else {
		var iframeName = window.name;
		if (iframeName != '') {
			parent.window.parent.ClosePropertiesDialog();
		}
		else {
			window.close();
		}
	}
}

function VerifySave(){
	var i;
	var blnSuccess;
	var strNTName;
	var NTCount=0;
	blnSuccess = true;	

	if (AddEmployee.cboDomain.selectedIndex > 5)
		{
		AddEmployee.txtPartnerName.value=AddEmployee.cboDomain.options[AddEmployee.cboDomain.selectedIndex].text; 
		AddEmployee.txtPartnerID.value=AddEmployee.cboDomain.options[AddEmployee.cboDomain.selectedIndex].value; 

		var txtOdmEmail = AddEmployee.txtEmail.value;
		if (txtOdmEmail.length > 30)
		    AddEmployee.txtNTName.value = txtOdmEmail.substr(0, 30);
		else
		    AddEmployee.txtNTName.value = txtOdmEmail;
		}
	else
		{		
		AddEmployee.txtPartnerName.value=""; 
		AddEmployee.txtPartnerID.value=1;
    }    
	if (AddEmployee.txtEmail.value == "@hp.com" || AddEmployee.txtEmail.value == "")
	{
			window.alert("Email Address is required.");
			AddEmployee.txtEmail.focus();
			blnSuccess = false;
    }
    else if (AddEmployee.txtEmail.value.indexOf("@hp.com") < 0 && (AddEmployee.cboDomain.value=="auth" || AddEmployee.cboDomain.value=="americas" || AddEmployee.cboDomain.value=="asiapacific" || AddEmployee.cboDomain.value=="emea"))
    {        
	    window.alert("HP users must use @hp.com in the email address");
	    AddEmployee.txtEmail.focus();
	    blnSuccess = false;
	}    
	else if (!VerifyEmail(AddEmployee.txtEmail.value))
	{
	    window.alert("You must enter a valid Email Address.");
	    AddEmployee.txtEmail.focus();
	    blnSuccess = false;
	}
	else if (AddEmployee.txtNTName.value == "auth\\" || AddEmployee.txtNTName.value == "")
		{
			window.alert("NT User Name is required. You can find the NT Name for a person in the Alias field of the Outlook address book.");
			AddEmployee.txtNTName.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.txtNTName.value.indexOf('\\') >= 0)
		{
			window.alert("You can not enter a domain name into the NT User Name field.");
			AddEmployee.txtNTName.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.txtFirstName.value == "")
		{
			window.alert("First Name is required.");
			AddEmployee.txtFirstName.focus();
			blnSuccess = false;
		}

	else if (AddEmployee.txtLastName.value == "")
		{
			window.alert("Last Name is required.");
			AddEmployee.txtLastName.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.txtPhone.value == "")
		{
			window.alert("Phone Number is required.");
			AddEmployee.txtPhone.focus();
			blnSuccess = false;
		}

	else if (AddEmployee.cboGroup.selectedIndex == 0)
		{
			window.alert("Group is required.");
			AddEmployee.cboGroup.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.cboDomain.selectedIndex == 0 || AddEmployee.cboDomain.selectedIndex == 5)
		{
			window.alert("NT Domain or Partner name is required.");
			AddEmployee.cboDomain.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.cboDomain.selectedIndex > 5 && AddEmployee.cboVPN.selectedIndex == 0 && AddEmployee.txtID.value == "")
		{
			window.alert("You must specify whether or not the ODM user will be allowed to log into Pulsar.");
			AddEmployee.cboVPN.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.cboDomain.selectedIndex > 5 && AddEmployee.cboVPN.selectedIndex == 1 && AddEmployee.cboFTP.selectedIndex == 0 && AddEmployee.txtID.value == "")
		{
			window.alert("You must specify whether or not the ODM user will need access to download deliverables from Pulsar.");
			AddEmployee.cboFTP.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.cboDomain.selectedIndex > 5 && trim(AddEmployee.txtNotes.value) == "" && AddEmployee.txtID.value == "")
		{
			window.alert("You must specify roles or permissions required by this user.");
			AddEmployee.txtNotes.focus();
			blnSuccess = false;
		}
	else if (AddEmployee.cboDivision.selectedIndex == 0)
		{
			window.alert("Division is required.");
			AddEmployee.cboDivision.focus();
			blnSuccess = false;
		}
    else if (AddEmployee.txtEmail.value.indexOf("@hp.com") > 0 && !(AddEmployee.cboDomain.value=="auth" || AddEmployee.cboDomain.value=="americas" || AddEmployee.cboDomain.value=="asiapacific" || AddEmployee.cboDomain.value=="emea"))
    {        
	    window.alert("You must enter a valid Domain with Email Address.");
	    AddEmployee.txtEmail.focus();
	    blnSuccess = false;
	}

	//Look for dup NT Names or Email Addresses
	else
		{	
			if (AddEmployee.cboDomain.selectedIndex > 5 && AddEmployee.txtID.value == "") 
			    {
				for (i=0;i<AddEmployee.lstODMEmails.length;i++)
					{
					if (AddEmployee.lstODMEmails.options[i].text.toLowerCase() == AddEmployee.txtEmail.value.toLowerCase())
						{
						NTCount=NTCount +1 
						}	
					}
				if ( (NTCount > 0 && AddEmployee.txtID.value == "") || (NTCount > 1 && AddEmployee.txtID.value != "") )
					{
					window.alert("That email address is already used.  This person may already have an active account.");
					AddEmployee.txtEmail.focus();
					blnSuccess = false;
					}
					
			    }
			else if (AddEmployee.txtNTName.value.toLowerCase() != "[requested]") 
				{
				strNTName = AddEmployee.cboDomain.value + "\\" + AddEmployee.txtNTName.value;
				for (i=0;i<AddEmployee.lstNTNames.length;i++)
					{
					if (AddEmployee.lstNTNames.options[i].text.toLowerCase() == strNTName.toLowerCase())
						{
						NTCount=NTCount +1 
						}	
					}
				if ( (NTCount > 0 && AddEmployee.txtID.value == "") || (NTCount > 1 && AddEmployee.txtID.value != "") )
					{
					window.alert("That NTName is already used.");
					AddEmployee.txtNTName.focus();
					blnSuccess = false;
					}
					

				}
		}
		
	

	return blnSuccess;
}

function cmdOK_onclick() {
    if (VerifySave()) {
        document.getElementById("cmdCancel").disabled = true;
        document.getElementById("cmdOK").disabled = true;
        document.getElementById("txtNTName").disabled = false;
        window.AddEmployee.submit();
        window.parent.location.href = "/pulsarplus/today"
    }
}

function cboDomain_onclick() {
	if(AddEmployee.txtID.value=="" )
		{
		if (AddEmployee.cboDomain.selectedIndex > 4)
			{
			VPNRow.style.display="none";
			NTNameRow.style.display="none";
			NoteStar.style.display="none";
			//PDDAccessRow.style.display ="none"
			}
		else
			{
			VPNRow.style.display="none";
			FTPRow.style.display="none";
			NTNameRow.style.display="";
			NoteStar.style.display="none";
			//PDDAccessRow.style.display ="none"
			}
		}
}

function cboVPN_onclick() {
	if(AddEmployee.txtID.value=="" )
		{
		if (AddEmployee.cboDomain.selectedIndex > 4 && AddEmployee.cboVPN.selectedIndex==1)
			FTPRow.style.display="";
		else
			FTPRow.style.display="none";
		}
}


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

function window_onload() {
	AddEmployee.txtFirstName.focus();
}

function ApproveODMAccount(ID) {
	var strID;
	strID = window.showModalDialog("EmployeeODMAccountApproval.asp?ID=" + ID ,"","dialogWidth:200px;dialogHeight:200px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No"); 
	if (typeof(strID) != "undefined")
		{
			if (strID== "1")
				ApprovalRow.style.display="none";
		}
}


//-->
</SCRIPT>
</head>
<STYLE>
TEXTAREA{
	FONT-FAMILY: Tahoma;
	FONT-SIZE=10pt;
}
a:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
</STYLE>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">


<%



	dim strGroups
	dim strGroup
	dim cn
	dim rs
	dim cm
	dim p
	dim CurrentUser
	dim CurrentUserDomain
	dim cnString
	dim strNTName
	dim strName
	dim strODMEmails
	dim strPhone
	dim strEmail
	dim strWorkgroupID
	dim strDivision
	dim strPartner
	dim strPartnerName
	dim strDomain
	dim strFirstName
	dim strLastName
	dim CurrentUserSysAdmin
	dim strDomainList
  
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
  
	strName = ""
	strPhone = ""
	strEmail = ""
	strNTName = ""
	strWorkgroupID = ""
	strDivision = ""
	strPartner = ""
	strDomain = ""
	strODMEmails = ""
	strCurrentUser = Session("LoggedInUser")

	if instr(strCurrentUser,"\") then
		CurrentUserDomain = left(strCurrentUser,instr(strCurrentUser,"\")-1)
	else
		CurrentUserDomain = ""
	end if	
	

    'Get User
	dim CurrentDomain
    dim blnSupportAdmin

    CurrentUser = lcase(Session("LoggedInUser"))

	if instr(CurrentUser,"\") > 0 then
		CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
		CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = CurrentUser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing

	if not (rs.EOF and rs.BOF) then
        CurrentUserSysAdmin = rs("SystemAdmin")
	end if
	rs.Close
 
    blnSupportAdmin = CBool(CurrentUserSysAdmin)


	strDomainList = ""
	if request("ID") = "" and request("Source") <> "1" then 'Adding - Not from Today Page
		dim UserInfo
		dim x
		set UserInfo = new clsUser
			x = lookupuser(replace(lcase(Session("LoggedInUser")),"\",":"))
		strFirstName = userinfo.first
		strLastName = userinfo.last 
		strPhone = UserInfo.Phone
		strEmail = UserInfo.Email
		strNTName = lcase(Session("LoggedInUser"))
		strWorkgroupID = ""
		strDivision = UserInfo.Division
		if instr(strNTName,"\") then
			strDomain = left(strNTName,instr(strNTName,"\")-1)
			strNTName = mid(strNTName,instr(strNTName,"\")+1)
		else
			strDomain = ""
		end if
	end if
		
	if request("ID") <> "" then

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetEmployeeByID"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		if not (rs.EOF and rs.BOF) then
			strName = rs("Name") & ""
			strPhone = rs("Phone") & ""
			strEmail = rs("Email") & ""
			strNTName = rs("NTName") & ""
			strEmail = rs("Email") & ""
			strWorkgroupID = rs("WorkgroupID") & ""
			strDivision = rs("Division") & ""
			strDomain = rs("Domain") & ""
			strPartner = rs("PartnerID") & ""
		end if
		rs.Close
	
		if instr(strName,",") > 0 then
			strFirstName = mid(strName,instr(strName,",") + 2)
			strLastName = left(strName,instr(strName,",") -1)
		else
			strFirstName = strName
			strLastName = ""
		end if
	
	end if
  
	if false then 'request("ID") = "" and lcase(CurrentUserDomain) <> "americas" and lcase(CurrentUserDomain) <> "emea" and lcase(CurrentUserDomain) <> "asiapacific" then
		Response.Write "Access Denied.  You do not have access to add new employees to Pulsar. >" & strDomain & "<"
	else
		'Load Domain List
		strDomain = lcase(trim(strdomain))
		strDomainList = "<OPTION selected></OPTION>"
		strDomainList = strDomainList & "<optgroup label=""Domain"">"

		if strDomain = "americas" then
			strPartnerName = strDomain
			strDomainList = strDomainList & "<OPTION selected value=""americas"">americas</OPTION>"
		else
			strDomainList = strDomainList & "<OPTION value=""americas"">americas</OPTION>"
		end if
		if strDomain = "asiapacific" then
			strDomainList = strDomainList & "<OPTION selected value=""asiapacific"">asiapacific</OPTION>"
			strPartnerName = strDomain
		else
			strDomainList = strDomainList & "<OPTION value=""asiapacific"">asiapacific</OPTION>"
		end if
		if strDomain = "emea" then	
			strDomainList = strDomainList & "<OPTION selected value=""emea"">emea</OPTION>"
			strPartnerName = strDomain
		else
			strDomainList = strDomainList & "<OPTION value=""emea"">emea</OPTION>"
		end if  
		if strDomain = "auth" then	
			strDomainList = strDomainList & "<OPTION selected value=""auth"">auth</OPTION>"
			strPartnerName = strDomain
		else
			strDomainList = strDomainList & "<OPTION value=""auth"">auth</OPTION>"
		end if  
		if strDomain <> "" and strDomain <> "auth" and strDomain <> "emea" and strDomain <> "asiapacific" and strDomain <> "americas" and strDomain <> "excaliburweb" then	
			strDomainList = strDomainList & "<OPTION selected value=""" & strDomain & """>" & strDomain & "</OPTION>"
			strPartnerName = strDomain
		end if  
        strDomainList = strDomainList & "</optgroup>"

        Dim strPartnerType
		rs.Open "spListPartners 1, 0",cn,adOpenStatic
		do while not rs.EOF
			if rs("ID") <> 1 then
			    if trim(strPartnerType) <> trim(rs("PartnerType")) Then
                    strDomainList = strDomainList & "</optgroup><optgroup label=""" & trim(rs("PartnerType")) & """>"
                    strPartnerType = Trim(rs("PartnerType"))
			    End If
				if trim(strPartner) = trim(rs("ID")) then	
					strDomainList = strDomainList & "<OPTION selected value=""" & trim(rs("ID")) & """>" & rs("Name") & "</OPTION>"
					strPartnerName = rs("Name") 
				else
					strDomainList = strDomainList & "<OPTION value=""" & trim(rs("ID")) & """>" & rs("Name") & "</OPTION>"
				end if  
			end if
			rs.MoveNext
		loop
		rs.Close

        strDomainList = strDomainList & "</optgroup>"
		
	  
		rs.Open "spGetWorkgroups",cn,adOpenForwardOnly

		strGroups = "<option selected></option>"
		do while not rs.EOF
			if trim(strWorkgroupID) = trim(rs("ID")) then
				strGroups = strGroups & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				strGroup = rs("Name") & ""
			else
				strGroups = strGroups & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop

		rs.Close

        
		strODMEmails = ""
		rs.Open "Select Email from Employee with (NOLOCK) where partnerid > 1;",cn,adOpenForwardOnly 'removed Active=1. if ODM account has been disabled still counts as duplicated email
		do while not rs.EOF
			strODMEmails = strODMEmails & "<OPTION>" & lcase(trim(rs("email"))) & "</OPTION>"
			rs.MoveNext
		loop
        rs.close
        %>


<font face=verdana>
<FORM ACTION="employeesave.asp" METHOD="post" NAME="AddEmployee">
<font face=verdana size=3 color=black><b>
<%if request("ID") <> "" then%>
	Update Employee Registration
<%else%>
	Add New Employee
<%end if%>
</font></b><BR><BR>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">

<table border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr>
    <td valign=top width="160"><strong><font size=2>First&nbsp;Name:</font><font color=red size=1> *</font></strong></td>
    <td valign=top><INPUT id=txtFirstName name=txtFirstName style="WIDTH: 240px; HEIGHT: 22px" size=6 maxlength=30 value="<%=strFirstName%>"></td>
  </tr>
  <tr>
    <td width="160" valign=top><strong><font size=2>Last&nbsp;Name:</font><font color=red size=1> *</font></strong></td>
    <td valign=top><INPUT id=txtLastName name=txtLastName style="WIDTH: 240px; HEIGHT: 22px" size=6 maxlength=30 value="<%=strLastName%>"></td>
   </tr>
  <tr>
    <td width="160" valign=top nowrap><strong><font size=2>NT&nbsp;Domain/Partner:</font><font color=red size=1>&nbsp;*&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></strong></td>
    <%if request("ID") = "" or blnSupportAdmin = true then%>
		<td valign=top>
			<SELECT style="WIDTH: 240px;" id=cboDomain name=cboDomain LANGUAGE=javascript onchange="return cboDomain_onclick()">
				<%=strDomainList%>
			</SELECT>
		</td>
	<%else%>
		<td valign=top>
			<font size=2 face=verdana><%=strPartnerName%></font>
			<SELECT style="Display:none;WIDTH: 240px;" id=cboDomain name=cboDomain>
				<%=strDomainList%>
			</SELECT>
		</td>
	<%end if%>
   </tr>
  <tr ID=NTNameRow>
    <td width="160" valign=top><strong><font size=2>NT&nbsp;Login&nbsp;Name:</font><font color=red size=1> *</font></strong></td>
    <%if request("ID") = "" or blnSupportAdmin = true then%>
		<td valign=top><INPUT id=txtNTName name=txtNTName style="WIDTH: 240px; HEIGHT: 22px" size=6 value="<%=strNTName%>" maxlength=30></td>
	<%else%>
		<td valign=top><font size=2 face=verdana><%=strNTName%></font><INPUT id=txtNTName name=txtNTName style="display:none;WIDTH: 240px; HEIGHT: 22px" size=6 value="<%=strNTName%>" maxlength=30></td>
	<%end if%>
   </tr>
   
  <tr>
    <td width="160" valign=top><strong><font size=2>Phone:</font><font color=red size=1> *</font></strong></td>
    <td valign=top><INPUT id=txtPhone name=txtPhone style="WIDTH: 240px; HEIGHT: 22px" size=6 maxlength=30 value="<%=strPhone%>"></td>
   </tr>
   <%if blnSupportAdmin = true or request("ID") = "" then%>
	  <tr>  
	    <td width="160" valign=top><strong><font size=2>Group:</font><font color=red size=1> *</font></strong></td>
		<td valign=top>
			<SELECT style="WIDTH: 240px;" id=cboGroup name=cboGroup>
			<%=strGroups%>
			</SELECT></td>
	</tr>
  <%else%>
	<tr style=Display:none>
    <td width="160" valign=top><strong><font size=2>Group:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>
		<SELECT style="WIDTH: 240px;" id=cboGroup name=cboGroup>
		<%=strGroups%>
		</SELECT></td>
   </tr>
	<tr>
    <td width="160" valign=top><strong><font size=2>Group:</font><font color=red size=1> *</font></strong></td>
    <td valign=top><font size=2 face=verdana><%=strGroup%></font></td>
   </tr>
	
  <%end if%>
  <tr>
    <td width="160" valign=top><strong><font size=2>Division:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>
    <% if strPartner = "1" or request("ID") = "" or blnSupportAdmin = true then%>
		<SELECT style="WIDTH: 240px;" id=cboDivision name=cboDivision>
	<%else%>
		<font size=2 face=verdana>
		<%
		if strDivision=1 then
			Response.Write "Notebooks"
		elseif strDivision=2 then
			Response.Write "Desktops"
		elseif strDivision=3 then
			Response.Write "Other"
		end if
		%>
		</font>
		<SELECT style="DISPLAY:none;WIDTH: 240px;" id=cboDivision name=cboDivision>
	<%end if%>
		<Option selected></Option>
			<%if strDivision = "1" then%>
				<option selected value="1">Notebooks</option>
			<%else%>
				<option value="1">Notebooks</option>
			<%end if%>
			<%if strDivision = "2" then%>
				<option selected value="2">Desktops</option>
			<%else%>
				<option value="2">Desktops</option>
			<%end if%>
			<%if strDivision = "3" then%>
				<option selected value="3">Other</option>
			<%else%>
				<option value="3">Other</option>
			<%end if%>
		</SELECT></td>
   </tr>
   
  <tr>
    <td width="160" valign=top><strong><font size=2>Email&nbsp;Address:</font><font color=red size=1> *</font></strong></td>
    <% if strPartner = "1" or request("ID") = "" or blnSupportAdmin = true then%>
		<td valign=top><INPUT id=txtEmail name=txtEmail style="WIDTH: 240px; HEIGHT: 22px" size=10 value="<%=strEmail%>" maxlength=41><INPUT type="hidden" id=hidEmail name=hidEmail value="<%=strEmail%>"></td>
   <%else%>
		<td valign=top><INPUT type="hidden" id=txtEmail name=txtEmail style="WIDTH: 240px; HEIGHT: 22px" size=10 value="<%=strEmail%>" maxlength=41><font size=2 face=verdana><%=strEmail%></font></td>
   <%end if%>
   </tr>
   <TR ID=VPNRow style="display:none">
	<TD><b>ODM Access:<font color=red size=1> *</font></b></TD>
	<TD><SELECT id=cboVPN name=cboVPN style="WIDTH: 240px" LANGUAGE=javascript onchange="return cboVPN_onclick()">
			<OPTION value=0></OPTION>
			<OPTION selected value=1>Request a Login account (BPIA Access).</OPTION>
			<OPTION value=2>Do not allow Pulsar access.</OPTION>
		</SELECT>
	</TD>
   </TR>
   <TR ID=FTPRow style="display:none">
	<TD><b>Deliverable&nbsp;Downloads:<font color=red size=1>&nbsp;*</font></b></TD>
	<TD><SELECT id=cboFTP name=cboFTP style="WIDTH: 240px">
			<OPTION value=0></OPTION>
			<OPTION selected value=1>Request an SFTP download account.</OPTION>
			<OPTION value=2>Do not allow deliverable downloads.</OPTION>
		</SELECT>
	</TD>
   </TR>

   <%if request("ID") = "" then%>
  <tr>
  <%else%>
  <tr style="display:none">
  <%end if%>
    <td valign=top width="160"><strong><font size=2>Notes/Requests:</font><font style="display:none" color=red size=1 ID=NoteStar> *</font></strong><BR><font size=1 color=green face=verdana>List roles or permissions needed, etc.</font></td>
    <td valign=top>
    <TEXTAREA rows=2 style="width:100%" id=txtNotes name=txtNotes></TEXTAREA>
    </td>
  </tr>
<% if strNTName = "[requested]" and strPartner <> "1" and request("ID") <> "" and blnSupportAdmin = true then%>  
<TR ID=ApprovalRow>
    <td valign=top width="160"><strong><font size=2>Approval:</font></td>
    <td valign=top><a href="javascript: ApproveODMAccount(<%=clng(request("ID"))%>);">Approve Login Account Request</a>
	</TD>
</TR>
<%end if%>

</table>
<SELECT style="Display:none" size=2 id=lstNTNames name=lstNTNames>
    <% 
		rs.Open "Select NTName,ntDomain from Userinfo with (NOLOCK) where isactiveingal=1;",cn,adOpenForwardOnly
		do while not rs.EOF
			response.Write "<OPTION>" & rs("ntdomain") & "\" & rs("NTName") & "</OPTION>"
			rs.MoveNext
		loop
        rs.close
    %>
</SELECT>
<INPUT type="hidden" id=txtPartnerName name=txtPartnerName value="">
<INPUT type="hidden" id=txtPartnerID name=txtPartnerID>
<INPUT type="hidden" id=txtCurrentUser name=txtCurrentUser value=<%=strCurrentUser%>>

<table width="437" border=0>
  <tr><TD align=right>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
  </TD></tr>
</table>
    <select style="display:none" id="lstODMEmails"><%=strODMEmails%></select>

</FORM>

</font>
<%
	end if

Class clsUser
	Public First
	Public Last
	Public Email
	Public Phone
	Public Division
End Class


Private Function LookupUser(FindNTName)
  
   Dim ldap_server,ldap_base,ldap_port,ldap_ssl_port
   Dim objADOconn,strADOQueryString,objRS,adspath,numrecs,ADSConn 
   Dim user_filter,user_dn,myUser,cn,sn,businessunit
   Dim givenname,TelephoneNumber,employeeNumber,NTName,pos,Email,Manager,i
   
    on error resume next
   ldap_server = "ldap.hp.com"
   ldap_port = "389"
   ldap_ssl_port = "636"
   ldap_base = "o=hp.com"
        
   'Connect to LDAP directory
   Set objADOconn = CreateObject("ADODB.Connection")
   objADOconn.Provider = "ADsDSOObject"
   objADOconn.Properties("Encrypt Password") = False

   objADOconn.CommandTimeout = 300
   objADOconn.ConnectionTimeout = 300
   objADOconn.Open "Active Directory Provider"
    
    user_filter = "ntUserDomainId=" & UCase(Trim(FindNTName))
   strADOQueryString = "<LDAP://" & ldap_server & ":" & ldap_port & "/" & ldap_base & ">;(" & user_filter & ");adspath;subtree"
   Set objRS = objADOconn.Execute(strADOQueryString)
   
   numrecs = 0
   If Not objRS.EOF Then
       Do While Not objRS.EOF
         '  Debug.Print objRS.Fields(0).Value
           numrecs = numrecs + 1
           adspath = objRS.Fields(0).Value
           objRS.MoveNext
           If numrecs > 100 Then
            Exit Do
           End If
       Loop
    End If

    objRS.Close
    If numrecs > 0 Then
    
    Set ADSConn = GetObject(adspath)
    ADSConn.GetInfoEx Array("uid", "sn", "givenname", "telephonenumber", "employeenumber"), 0

    user_dn = ADSConn.adspath
    pos = InStr(10, user_dn, "/")
    pos = Len(user_dn) - pos
    user_dn = Right(user_dn, pos)
   
    sn = ""
    sn = ADSConn.Get("sn")
    If Err.Number Then
        Err.Clear
    End If
   
   businessunit = ""
   businessunit = ADSConn.Get("hpBusinessUnit") 'hpOrganizationChart
   If Err.Number Then
        Err.Clear
   End If
	if instr(lcase(businessunit),"notebook") > 0 then
		businessunit = "1"
	elseif instr(lcase(businessunit),"desktop") > 0 then
		businessunit = "2"
	elseif instr(lcase(businessunit),"workstation") > 0 then
		businessunit = "2"
	elseif instr(lcase(businessunit),"business pc") > 0 then
		businessunit="2"
	elseif instr(lcase(businessunit),"consumer pc") > 0 then
		businessunit="2"
	elseif instr(lcase(businessunit),"handhelds") > 0 then
		businessunit="1"
	else
		businessunit="3"
	end if
   
  
   givenname = ""
   givenname = ADSConn.Get("givenname")
   If Err.Number Then
        Err.Clear
   End If
   
   TelephoneNumber = ""
   TelephoneNumber = ADSConn.Get("telephonenumber")
   If Err.Number Then
        Err.Clear
   End If
     
    Email = ""
    Email = ADSConn.Get("uid")
   If Err.Number Then
        Err.Clear
   End If
     
	userinfo.first = givenname
	UserInfo.Last = sn
	UserInfo.Email = email
	UserInfo.Division = businessunit
	UserInfo.Phone = telephonenumber
    End If

End Function


		set rs = nothing
		cn.Close
		set cn = nothing


%>

</body>
</html>
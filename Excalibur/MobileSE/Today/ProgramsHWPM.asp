
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



function ChooseEmployee(myControl){
	var ResultArray;

	ResultArray = window.showModalDialog("ChooseEmployee.asp","","dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 

	if (typeof(ResultArray) != "undefined")
		{
		if (ResultArray[0] != 0)
			{
			myControl.options[myControl.length] = new Option(ResultArray[1],ResultArray[0]);
			myControl.selectedIndex = myControl.length-1;
			}
		}
}

function UpdateOwners(ID){
    window.open ("OTSHWComponents.asp?ID=" + ID,"_blank","width=700,height=350,location=1,menubar=1,resizable=1,scrollbars=1,status=1, titlebar=1,toolbar=1");
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
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserEmail
	
	dim	strCommHWPM 
	dim strVideoMemoryPM 
	dim strGraphicsControllerPM 
	dim strProcessorPM
	dim strProcessorPMList
	dim strVideoMemoryPMList
	dim strGraphicsControllerPMList
	dim strCommHWPMList
	dim strSelectedApproverIDs
	dim strSelectedApprovers
	dim blnPM
	
	blnPM = false
	
	strCommHWPM = ""
	strVideoMemoryPM = "" 
	strGraphicsControllerPM = ""
	strProcessorPM = ""
	strProcessorPMList = ""
	strVideoMemoryPMList = ""
	strGraphicsControllerPMList = ""
	strCommHWPMList	 = ""
	strSelectedApproverIDs = ""
	strSelectedApprovers = ""
		
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn

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
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("Email") & ""
	end if
	rs.Close

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

	if rs.EOF and rs.BOF then
		strProductName= ""
	else
		strProductName= rs("DotsName") & ""
		strCommHWPM = rs("CommHWPMID") & ""
		strVideoMemoryPM = rs("VideoMemoryPMID") & ""
		strGraphicsControllerPM = rs("GraphicsControllerPMID") & ""
		strProcessorPM = rs("ProcessorPMID") & ""
	end if
	rs.Close
	
        blnPM = false
        rs.open "spListSystemTeamDropdowns " & clng(request("ID")),cn
        do while not rs.eof
            if trim(currentuserid) = trim(rs("ID")) and lcase(rs("role") & "") = "platform development" then
                blnPM = true
                exit do
            end if
            rs.movenext
        loop
        rs.close
   	
	if trim(strProductName) = "" or trim(CurrentUserID) = "" then
		Response.write "Unable to find the requested product."
	else
		Response.write "<h3>" & strProductName & " Properties</h3>" 

		dim blnProcessorPM
		dim blnGraphicsControllerPM
		dim blnVideoMemoryPM
		dim blnCommPM
	  
        if blnPM then
		    blnProcessorPM = true
		    blnGraphicsControllerPM = true
		    blnVideoMemoryPM = true
		    blnCommPM = true
        else
		    blnProcessorPM = false
		    blnGraphicsControllerPM = false
		    blnVideoMemoryPM = false
		    blnCommPM = false
	    end if

		rs.Open "spGetHardwareTeamAccessList " & clng(CurrentUserID),cn,adOpenForwardOnly
		do while not rs.EOF
			if rs("HWTeam") = "Comm" and rs("Products") > 0 then
				blnCommPM = true
			elseif rs("HWTeam") = "VideoMemory" and rs("Products") > 0 then
				blnVideoMemoryPM = true
			elseif rs("HWTeam") = "Processor" and rs("Products") > 0 then
				blnProcessorPM = true
			elseif rs("HWTeam") = "GraphicsController" and rs("Products") > 0 then
				blnGraphicsControllerPM = true
			end if
			rs.MoveNext
		loop
		rs.Close	

		
		rs.Open "spListSystemTeamDropdowns",cn,adOpenStatic 
		do while not rs.EOF

			if rs("Role") = "Chipset/Processor PM" then
				if trim(strProcessorPM) = trim(rs("ID")) then
					strProcessorPMList = strProcessorPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				else
					strProcessorPMList = strProcessorPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				end if
			end if

			if rs("Role") = "Video Memory PM" then
				if trim(strVideoMemoryPM) = trim(rs("ID")) then
					strVideoMemoryPMList = strVideoMemoryPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				else
					strVideoMemoryPMList = strVideoMemoryPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				end if
			end if

			if rs("Role") = "Graphics Controller PM" then
				if trim(strGraphicsControllerPM) = trim(rs("ID")) then
					strGraphicsControllerPMList = strGraphicsControllerPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				else
					strGraphicsControllerPMList = strGraphicsControllerPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				end if
			end if

			if rs("Role") = "Comm HW PM" then
				if trim(strCommHWPM) = trim(rs("ID")) then
					strCommHWPMList = strCommHWPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				else
					strCommHWPMList = strCommHWPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
	

		set cm=nothing
		set rs=nothing
		set cn=nothing


		dim DisplayProcessorRow
		dim DisplayGraphicsControllerRow
		dim DisplayVideoMemoryRow
		dim DisplayCommHWRow
		dim FieldsRequested
		
		FieldsRequested = 0

		if 	blnProcessorPM then
			DisplayProcessorRow = ""
			FieldsRequested = FieldsRequested + 1
		else
			DisplayProcessorRow = "none"
		end if

		if 	blnGraphicsControllerPM then
			DisplayGraphicsControllerRow = ""
			FieldsRequested = FieldsRequested + 1
		else
			DisplayGraphicsControllerRow = "none"
		end if

		if 	blnVideoMemoryPM then 
			DisplayVideoMemoryRow = ""
			FieldsRequested = FieldsRequested + 1
		else
			DisplayVideoMemoryRow = "none"
		end if

		if 	blnCommPM then
			DisplayCommHWRow = ""
			FieldsRequested = FieldsRequested + 1
 		else
			DisplayCommHWRow = "none"
		end if
	
	
%>
<FORM ACTION="ProgramSaveHWPM.asp" METHOD="post" id="ProgramInput">
<table ID=tabHW border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>

   <tr style="display:<%=DisplayProcessorRow%>">
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Chipset/Processor&nbsp;PM:&nbsp;</font></strong></td>
    <td><SELECT id=cboProcessorPM name=cboProcessorPM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strProcessorPMList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdProcessorPMAdd name=cmdProcessorPMAdd LANGUAGE=javascript onclick="ChooseEmployee (cboProcessorPM)">
	</td>
	</tr>
   <tr style="display:<%=DisplayCommHWRow%>">
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Comm&nbsp;HW&nbsp;PM:&nbsp;</font></strong></td>
    <td><SELECT id=cboCommHWPM name=cboCommHWPM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strCommHWPMList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdCommHWPMAdd name=cmdCommHWPMAdd LANGUAGE=javascript onclick="ChooseEmployee (cboCommHWPM)">
	</td>
	</tr>
   <tr style="display:<%=DisplayGraphicsControllerRow%>">
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Graphics&nbsp;Controller&nbsp;PM:&nbsp;</font></strong></td>
    <td><SELECT id=cboGraphicsControllerPM name=cboGraphicsControllerPM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strGraphicsControllerPMList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdGraphicsControllerPMAdd name=cmdGraphicsControllerAdd LANGUAGE=javascript onclick="ChooseEmployee (cboGraphicsControllerPM)">
	</td>
	</tr>
   <tr style="display:<%=DisplayVideoMemoryRow%>">
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Video&nbsp;Memory&nbsp;PM:&nbsp;</font></strong></td>
    <td><SELECT id=cboVideoMemoryPM name=cboVideoMemoryPM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strVideoMemoryPMList%>			
		</SELECT>&nbsp;<INPUT style="Display:" type="button" value="Add" id=cmdVideoMemoryPMAdd name=cmdVideoMemoryPMAdd LANGUAGE=javascript onclick="ChooseEmployee (cboVideoMemoryPM)">
	</td>
	</tr>
	<%if blnPM then %>
   <tr>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>OTS&nbsp;HW&nbsp;Components&nbsp;</font></strong></td>
    <td><a href="javascript: UpdateOwners(<%=clng(request("ID"))%>);">Update Owners</a></td>
	</tr>
	<%end if %>
</table>    
<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT style="Display:none" type="text" id=txtProductName name=txtProductName value="<%=strProductName%>">
<INPUT style="Display:none" type="text" id=txtCurrentUserEmail name=txtCurrentUserEmail value="<%=CurrentUserEmail%>">
<INPUT style="Display:none" type="text" id=txtFieldsRequested name=txtFieldsRequested value="<%=FieldsRequested%>">

<INPUT style="Display:none" type="text" id=tagGraphicsControllerPM name=tagGraphicsControllerPM value="<%=strGraphicsControllerPM%>">
<INPUT style="Display:none" type="text" id=tagVideoMemoryPM name=tagVideoMemoryPM value="<%=strVideoMemoryPM%>">
<INPUT style="Display:none" type="text" id=tagCommHWPM name=tagCommHWPM value="<%=strCommHWPM%>">
<INPUT style="Display:none" type="text" id=tagProcessorPM name=tagProcessorPM value="<%=strProcessorPM%>">

</FORM>
<% end if%>
</body>
</html>
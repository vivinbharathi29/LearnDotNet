<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cboTodayLink_onchange(){
        if(cboTodayLink.selectedIndex==0 || cboTodayLink.selectedIndex==1 || cboTodayLink.selectedIndex==4)
            spnReportFormat.style.display="none";
        else
            spnReportFormat.style.display="";
    }


    function cmdOK_click() {
        if (txtName.value == "" && txtReportType.value=="1")
            alert("You must enter a name for the custom report.");
        else if (txtName.value == "")
            alert("You must enter a name for the profile.");
        else 
            {    
	        var OutArray = new Array();
		    OutArray[0]= txtName.value;
		    OutArray[1]= cboTodayLink.options[cboTodayLink.selectedIndex].value;
		    OutArray[2]= cboReportFormat.options[cboReportFormat.selectedIndex].value;

            if(navigator.appName != "Microsoft Internet Explorer" && navigator.appName != "Internet Explorer" && navigator.appName != "IE")
               if (typeof( window.parent.opener) != "undefined")
                    {
                    window.parent.opener.txtReturnValue.value=OutArray[0];
                    window.parent.opener.txtReturnValue2.value=OutArray[1];
                    window.parent.opener.txtReturnValue3.value=OutArray[2];
                    }
		    window.returnValue = OutArray;
	       // window.parent.opener='X';
	        //window.parent.open('','_parent','')
	        window.parent.close();				
            }
}

function window_onload(){
    txtName.focus();
}

//-->
</SCRIPT>
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">
<style>
body
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
}
td
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
}    
h1
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
    font-weight:bold;    
}
</style>
</HEAD>

<BODY bgcolor=ivory onload="javascript: window_onload();">

<%
    if trim(request("ID")) = "" then 
        if request("ReportType") = "1" then
            response.write "<h1>Add Custom Report</h1>"
        else
            response.write "<h1>Add Report Profile</h1>"
        end if
    else
        if request("ReportType") = "1" then
            response.write "<h1>Update Custom Report</h1>"
        else
            response.write "<h1>Update Report Profile</h1>"
        end if
    end if
	
    dim cn, rs
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	dim CurrentDomain, CurrentUser, CurrentUserID, CurrentUserDivision, CurrentUserPartner
    
    'Get User
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
		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = rs("PartnerID") & ""
	end if
	rs.Close

    dim strName
    dim strTodayPageLink
    dim strDefaultReportFormat
    dim OptionArray
    dim strDisplayReportSpan
    dim strDisplayFormatSpan
    dim ValueArray
    dim strNameTitle
    dim ShowLink

    if request("ReportType") = "1" then
        strNameTitle= "Report&nbsp;Name"
        ShowLink = "none"
    else
        strNameTitle= "Profile&nbsp;Name"
        ShowLink = ""
    end if

    strName = ""
    strTodayPageLink = "0"
    strDefaultReport = "0"
    
    if trim(request("ID")) <> "" then
        rs.open "spGetReportProfile " & clng(request("ID")),cn
        if not (rs.EOF and rs.bof) then
            strName = rs("ProfileName") & ""
            strTodayPageLink = trim(rs("TodayPageLink") & "")
            strDefaultReportFormat = trim(rs("DefaultReport") & "")
        end if
        rs.Close
    end if

    if trim(strTodayPageLink) = "0" or trim(strTodayPageLink) = "1" or trim(strTodayPageLink) = "4" then
        strDisplayReportFormatSpan = "none"
    else
        strDisplayReportFormatSpan = ""
    end if
%>
<table width="100%" style=" background-color:cornsilk;" border=1 bordercolor =tan cellspacing="0" cellpadding="2">
        <tr>
            <td nowrap width="120"><b><%=strNameTitle%>:</b>&nbsp; <font color="#ff0000" size="1">*</font></td>
            <td nowrap><input id="txtName" name="txtName" style="width:100%" maxlength=120 type="text" value="<%=strName%>"></td>
        </tr>

       <tr style="display:<%=ShowLink%>">
            <td nowrap width="120"><b>Today&nbsp;Page&nbsp;Link:&nbsp;&nbsp;&nbsp;</b></td>
            <td nowrap style="width:100%"> 
                <select id="cboTodayLink" style="width:320" onchange="javascript: cboTodayLink_onchange()">
                <%

                    OptionArray = split("0|No Link Displayed,1|Advanced Search screen,2|Summary Report,3|Detailed Report,4|Email OTS Owners...,5|Deliverable Status,6|Product Status",",")
                    for i = 0 to ubound(OptionArray)
                        ValueArray = split(OptionArray(i),"|")
                       ' if clng(ValueArray(0)) = 2 then
                       '     response.write "<optgroup style="""" label=""-- Standard Reports -------------------------""></optgroup>"
                       ' elseif clng(ValueArray(0)) > 5 then
                       '     response.write "<optgroup style="""" label=""-- Custom Reports -------------------------""></optgroup>"
                        'end if
                        if clng(strTodayPageLink) = clng(ValueArray(0)) then
                            response.write "<option selected value=" & ValueArray(0) & ">" & ValueArray(1) & "</option>"
                        else
                            response.write "<option value=" & ValueArray(0) & ">" & ValueArray(1) & "</option>"
                        end if
                    next 

                	rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	                strProfileOptions = ""
	                do while not rs.EOF
                        if clng(strTodayPageLink) = clng(rs("ID")) then
                            response.write "<option selected value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        else
                            response.write "<option value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        end if
                        rs.MoveNext
	                loop
	                rs.Close
		
	                rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	                do while not rs.EOF
                        if clng(strTodayPageLink) = clng(rs("ID")) then
                            response.write "<option selected value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        else
                            response.write "<option value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        end if
                    	rs.MoveNext
	                loop
	                rs.Close

	                rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	                do while not rs.EOF
                        if clng(strTodayPageLink) = clng(rs("ID")) then
                            response.write "<option selected value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        else
                            response.write "<option value=" & clng(rs("ID")) & ">" & rs("ProfileName") & "</option>"
                        end if
                	rs.MoveNext
	                loop
	                rs.Close


                 %>
                </select>            
                <span id=spnReportFormat style="white-space: nowrap;display:<%=strDisplayReportFormatSpan%>">
                &nbsp;&nbsp;Format:
                <select id="cboReportFormat" style="width:80">
                <%
                    OptionArray = split("HTML,Excel,Word",",")
                    for i = 0 to ubound(OptionArray)
                        if clng(strDefaultReportFormat) = clng(i) then
                            response.write "<option selected value=" & i & ">" & OptionArray(i) & "</option>"
                        else
                            response.write "<option value=" & i & ">" & OptionArray(i) & "</option>"
                        end if
                    next 
                 %>
                </select>
                </span>
           </td>
       </tr>
    </table>

<%
	set rs = nothing
	cn.Close
	set cn = nothing

%>
<INPUT type="hidden" id=txtProfileID name=txtProfileID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtReportType name=txtReportType value="<%=request("ReportType")%>">


</BODY>
</HTML>

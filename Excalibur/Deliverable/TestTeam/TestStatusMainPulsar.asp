<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var origStatus = "";
var origReceived = "";
var origNotes = "";

$(document).ready(function () {
    origStatus = $("#cboStatus").val();
    origReceived = $("#txtReceived").val();
    origNotes = $("#txtNotes").val();

    window.parent.frames["LowerWindow"].enableButton();
    $("#txtRedirect").val("TestStatusMainPulsar.asp?VersionID=" + $("#txtVersionID").val() + "&ProductID=" + $("#txtProductID").val() + "&FieldID=" + $("#txtFieldID").val() + "&ProductDeliverableReleaseID=" + $("#txtProdDelRelID").val() + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&RowID=" + $("#txtRowID").val());
});

function CheckTextSize(field, maxLength) {
	if (field.value.length > maxLength + 1)
		{
		field.value = field.value.substring(0, maxLength);
		alert("The maximum size of this field in 200 characters. You input has been truncated.");
		}
	else if (field.value.length >= maxLength)
		{
		window.event.keyCode=0;
		field.value = field.value.substring(0, maxLength);
		}
} 

function cboStatus_onclick() {
	if (frmMain.cboStatus.selectedIndex==2 || frmMain.cboStatus.selectedIndex==3)
		RequireNotes.style.display="";
	else
		RequireNotes.style.display="none";
}

function window_onload() {
	frmMain.cboStatus.focus();
}

function SwitchRelease(ProdID, VersionID, FieldID, ProductDeliverableReleaseID) {
    var isModified = false;
    $("#txtKeepItOpen").val(true);

    if (origStatus != $("#cboStatus").val()) {
        isModified = true;
    }
    else if (origReceived != $("#txtReceived").val()) {
        isModified = true;
    }
    else if (origNotes != $("#txtNotes").val()) {
        isModified = true;
    }
    
    if (isModified) {
        if (confirm("Do you want to save your changes for this Release?")) {
            $("#txtRedirect").val("TestStatusMainPulsar.asp?VersionID=" + VersionID + "&ProductID=" + ProdID + "&FieldID=" + FieldID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&RowID=" + $("#txtRowID").val());
            window.parent.frames["LowerWindow"].cmdOK_onclick();
        }
        else {
            document.location = "TestStatusMainPulsar.asp?VersionID=" + VersionID + "&ProductID=" + ProdID + "&FieldID=" + FieldID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&RowID=" + $("#txtRowID").val();
            window.parent.repositionParentWindow();
        }
    }
    else {
        document.location = "TestStatusMainPulsar.asp?VersionID=" + VersionID + "&ProductID=" + ProdID + "&FieldID=" + FieldID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&RowID=" + $("#txtRowID").val();
        window.parent.repositionParentWindow();
    }
}
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<LINK href="../../style/wizard style.css" type=text/css rel=stylesheet >
<font size=3 face=verdana><b>
<%
	if request("FieldID") = "2" then
		response.write "Update ODM HW Test Status</b><BR></font>"
	elseif request("FieldID") = "3" then
		response.write "Update COMM Test Status</b><BR></font>"
    elseif request("FieldID") = "4" then
		response.write "Update DEV Test Status</b><BR><BR></font>"
	else
		response.write "Update SE Test Status</b><BR></font>"
	end if

	dim cn 
	dim rs
	dim strName
	dim strVersion
	dim strRevision
	dim strPass
	dim strTypeID
	dim strModelNumber
	dim strPartNumber
	dim strEOLDate
	dim strVendor
	dim strPMEmail
	dim strStatus
	dim strUnitsReceived
	dim strTestNotes
	dim DisplayNotesRequired
	dim OTSRootText
	dim OTSVersionText
	dim OTSRootCount
	dim OTSVersionCount
	dim strRootID
	dim strOTSNumbers
  	dim CurrentUserPartner
  	dim strProductPartner

	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	
	
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("email") & ""
		CurrentUserPartner = rs("PartnerID") & ""
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
		CurrentUserPartner = 0
	end if
	rs.Close

	rs.Open "spGetProductVersion " & clng(request("ProductID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
  		strProductPartner = "0"
	else
  		strProductPartner = rs("PartnerID") & ""
	end if
	rs.Close

	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strProductPartner) <> trim(CurrentUserPartner) or trim(CurrentUserPartner) = "0" then
			set rs = nothing
			set cn=nothing
				
			'Response.Redirect "../../NoAccess.asp?Level=1"
		end if
	end if
	
    dim ProductDeliverableReleaseID, ProductDeliverableReleaseName, strReleaseLink, ProductDeliverableID
    ProductDeliverableReleaseID = 0
    ProductDeliverableReleaseName = ""
    strReleaseLink = ""

    dim intDefaultReleaseID
    if Request("ReleaseID") then
        intDefaultReleaseID = clng(Request("ReleaseID"))	    
    end if

	if (intDefaultReleaseID = 0) and (Request("TodayPageSection") = "") then
	    strSql = "select top 1 pr.ID, pr.Name from ProductVersion_Release pvr join ProductVersionRelease pr on pr.ID = pvr.ReleaseID where pvr.ProductVersionID= " & clng(Request("ProdID")) & " order by pr.ReleaseYear desc, pr.ReleaseMonth desc;"
		rs.open strSql,cn
		if not (rs.EOF and rs.BOF) then
            intDefaultReleaseID = rs("ID")
		end if	
		rs.close
    end if
    
    if Request("ProductDeliverableReleaseID") then
        ProductDeliverableReleaseID = trim(Request("ProductDeliverableReleaseID"))            
    end if    

    if Request("TodayPageSection") = "" then
        strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID, pdr.ProductDeliverableID " &_
                 "from Product_Deliverable pd " &_
                 "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID and pdr.targeted = pd.targeted " &_
                 "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                "where pd.ProductVersionID= " & request("ProductID") & " and pd.DeliverableVersionID= " & Request("VersionID") & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
    else 
       strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID, pdr.ProductDeliverableID " &_
                "from Product_Deliverable pd " &_
                "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                "where pd.ProductVersionID= " & request("ProductID") & " and pd.DeliverableVersionID= " & Request("VersionID")
        
       if ProductDeliverableReleaseID > 0 then
            strSql = strSql & " and pdr.id = " & ProductDeliverableReleaseID & " order by pvr.id desc"
       else
            if intDefaultReleaseID > 0 then
                strSql = strSql & " and pdr.id = " & intDefaultReleaseID & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
            else 
                strSql = strSql & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
            end if
       end if
    end if

	rs.open strSql, cn
        	    
    strReleaseLink = "Switch Releases:&nbsp;"
    
    Do until rs.EOF            
        
        if strReleaseLink <> "Switch Releases:&nbsp;" then
            strReleaseLink = strReleaseLink & " | " 
        end if

        if ProductDeliverableReleaseID > 0 and ProductDeliverableReleaseID = trim(rs("ID")) then
            strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
            ProductDeliverableReleaseName = rs("Name")
            ProductDeliverableID = rs("ProductDeliverableID")
        else
            if rs("ReleaseID") = intDefaultReleaseID and ProductDeliverableReleaseID = 0 then
                strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                ProductDeliverableReleaseID = rs("ID")
                ProductDeliverableReleaseName = rs("Name")
                ProductDeliverableID = rs("ProductDeliverableID")
            else
                strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProductID") & "," & Request("VersionID") & "," & Request("FieldID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
            end if
       end if
              
       rs.MoveNext
    Loop

    if ProductDeliverableReleaseID = 0 then
        strReleaseLink = "Switch Releases:&nbsp;"
        dim count
        count = 0
        rs.MoveFirst
        Do until rs.EOF            
            if strReleaseLink <> "Switch Releases:&nbsp;" then
                strReleaseLink = strReleaseLink & " | " 
            end if

            if  count = 0 then
                strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                ProductDeliverableReleaseID = rs("ID")
                ProductDeliverableReleaseName = rs("Name")
                ProductDeliverableID = rs("ProductDeliverableID")
            else
                strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProductID") & "," & Request("VersionID") & "," & Request("FieldID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
            end if
               
            count = count + 1
            rs.MoveNext
        Loop
    end if
    rs.Close
    	
	response.Write("<span style='font-family: Verdana; font-size: 9pt;'>" & strReleaseLink & "</span>")
	
	rs.Open "spGetVersionProperties4Web " & clng(request("VersionID")),cn,adOpenForwardOnly
	if not(rs.EOF and rs.BOF) then
		strName = trim(rs("DeliverableName") & "")
		strDelID = "<a target=_blank href=""../../Query/DeliverableVersionDetails.asp?Type=1&RootID=" & rs("RootID") & "&ID=" & rs("VersionID") & """>" & rs("versionID") & "</a>"
		strVersion = trim(rs("Version") & "")
		strRootID = rs("RootID") & ""
		strRevision = trim(rs("Revision") & "")
		strPass = trim(rs("Pass") & "")
		strTypeID = trim(rs("TypeID") & "")
		strModelNumber = trim(rs("ModelNumber") & "")
		strPartNumber = trim(rs("PartNumber") & "")
		strEOLDate = rs("EOLDate") & ""
		if trim(rs("VersionVendor") & "") <> "" then
			strVendor = rs("VersionVendor") & ""
		else
			strVendor = rs("Vendor") & ""
		end if
	end if
	rs.Close
	
	rs.Open "spGetCategoryPM " & clng(request("VersionID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strPMEmail=""
	else
		strPMEmail = rs("Email") & ""
	end if
	rs.Close


	rs.Open "spGetTestLeadStatusPulsar " & clng(ProductDeliverableReleaseID) & "," & clng(request("FieldID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strProductName =""
		strStatus = ""
		strUnitsReceived = ""
		strTestNotes = ""
	else
		strProductName =rs("Product") & ""
		strStatus = rs("TestStatus") & ""
		strUnitsReceived = rs("UnitsReceived") & ""
		strTestNotes = rs("TestNotes") & ""
	end if
	rs.Close
	
	if strStatus = "2" or strStatus = "3" then
		DisplayNotesRequired = ""
	else
		DisplayNotesRequired = "none"
	end if 
	
	set rs = nothing
	cn.Close
	set cn = nothing

	if strProductName = "" or strName = "" then
		Response.write "<font size=2 face=verdana>Deliverable not found. (" & request("ID") & ")</font>"	
	else

%>
<form ID=frmMain action="TestStatusSavePulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Deliverable:</b></TD>
	<TD width="100%" colspan=3><%=strName%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Product:</b></TD>
	<TD><%=strProductName%>&nbsp;</TD>
	<TD valign=top><b>Deliverable&nbsp;ID:</b></TD>
	<TD width="100%" colspan=3><%=strDelID%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>HW Version:&nbsp;&nbsp;&nbsp;&nbsp;</b></TD>
	<TD width=40%><%=strVersion%>&nbsp;</TD>
	<TD nowrap valign=top><b>Vendor:</b></TD>
	<TD width=60%><%=strVendor%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>FW Version:</b></TD>
	<TD><%=strRevision%>&nbsp;</TD>
	<TD nowrap valign=top><b>Model&nbsp;Number:</b></TD>
	<TD><%=strModelNumber%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Revision:</b></TD>
	<TD><%=strPass%>&nbsp;</TD>
	<TD nowrap valign=top><b>Part&nbsp;Number:</b></TD>
	<TD><%=strPartNumber%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk style=display:none>
	<TD nowrap valign=top><b>OTS - Root:</b></TD>
	<TD><%=OTSRootText%></TD>
	<TD nowrap valign=top><b>OTS - Version:</b></TD>
	<TD>4 Open Observations</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Test&nbsp;Status:</b></TD>
	<TD>
		<SELECT style="width:100%" id=cboStatus name=cboStatus LANGUAGE=javascript onclick="return cboStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<%if trim(strStatus) = "1" then%>
				<OPTION value=1 selected >Passed</OPTION>
			<%else%>
				<OPTION value=1>Passed</OPTION>
			<%end if%>
			<%if trim(strStatus) = "2" then%>
				<OPTION value=2 selected>Failed</OPTION>
			<%else%>
				<OPTION value=2>Failed</OPTION>
			<%end if%>
			<%if trim(strStatus) = "3" then%>
				<OPTION value=3 selected>Blocked</OPTION>
			<%else%>
				<OPTION value=3>Blocked</OPTION>
			<%end if%>
			<%if trim(strStatus) = "4" then%>
				<OPTION value=4 selected>Watch</OPTION>
			<%elseif clng(request("FieldID")) = 3 then%>
				<OPTION value=4>Watch</OPTION>
			<%end if%>
			<%if trim(strStatus) = "5" then%>
				<OPTION value=5 selected>N/A</OPTION>
			<%elseif clng(request("FieldID")) = 3 then%>
				<OPTION value=5>N/A</OPTION>
			<%end if%>
		</SELECT>
	</TD>
	<TD nowrap valign=top><b>Samples&nbsp;Available:&nbsp;&nbsp;</b></TD>
	<TD ><INPUT style="width:60" maxlength=3 type="text" id=txtReceived name=txtReceived value="<%=strUnitsReceived%>"> <font size=1 color=green face=verdana>Total for your group.</font>
	</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Test&nbsp;Notes:</b>&nbsp;<span style="Display:<%=DisplayNotesRequired%>" ID=RequireNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD colspan=3>
	<TEXTAREA rows=4 style="width:100%" id=txtNotes name=txtNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"><%=strTestNotes%></TEXTAREA>
	</TD>
</TR>
</table>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtFieldID name=txtFieldID value="<%=request("FieldID")%>">
<INPUT type="hidden" id=txtPMEmail name=txtPMEmail value="<%=strPMEmail%>">
<input type="hidden" id="txtRedirect" name="txtRedirect" value="" />
<input type="hidden" id="txtProdDelRelID" name="txtProdDelRelID" value="<%=ProductDeliverableReleaseID%>" />
<input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="false" />
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=request("TodayPageSection")%>" />
<input type="hidden" id="txtRowID" name="txtRowID" value="<%=request("RowID")%>" />
</form>

	<%end if%>
</BODY>
</HTML>

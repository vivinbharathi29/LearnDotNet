<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
    <script src="_ScriptLibrary/jsrsClient.js"></script>
<!--<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>-->

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Update Status</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
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

function cboDevStatus_onclick() {
	if (frmUpdateDevStatus.cboDevStatus.value==2)
		RequireNotes.style.display="";
	else
		RequireNotes.style.display="none";
		
}

function window_onload() {
	if (frmUpdateDevStatus.txtStatusID.value=="")
		frmUpdateDevStatus.cboDevStatus.focus();
	else
		frmUpdateDevStatus.txtComments.focus();
}

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
        <link href="style/wizard%20style.css" rel="stylesheet" />
<!--<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">-->
<%
    
	if trim(request("ID")) = "" then
		Response.Write "Unable to find the requested record."
	else
'	    response.Write trim(request("ID"))
		dim cn
		dim rs
		dim DisplayNotesRequired
		dim strProduct
		dim strDeliverable
		dim strVersion
		dim strPart
		dim strModel
		dim strVendor
		dim strType
		dim strTestNotes
		dim strDefaultStatus
		dim strStatusRejectText
		dim strStatusAcceptText
		dim strStatusReviewText
        dim IDArray
        dim myIDarr

		strStatusRejectText = ""
		strStatusAcceptText = ""
		strStatusReviewText = ""
        
        IDArray = split(request("ID"),",")
		
		DisplayNotesRequired = "none"
		strDefaultStatus=""
		if request("StatusID") <> "" then
			if isnumeric(request("StatusID")) then
				strDefaultStatus = request("StatusID")
			end if
		end if
		if trim(strDefaultStatus)="2" then
			DisplayNotesRequired = ""
		end if
	
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")

        myIDarr = split(IDArray(0), "_")

		rs.open "spGetProductDeliverableSummaryByID " & clng(myIDarr(0)) & "," & clng(myIDarr(1)), cn, adOpenForwardOnly
		if (rs.EOF and rs.BOF) then
			strProduct = ""
			strDeliverable = ""
			strVersion = ""
			strPart = ""
			strModel = ""
			strVendor = ""
			strType=""
			strTestNotes = ""
		else
			strType=rs("TypeID") & ""
			strProduct = rs("product") & ""
			strDeliverable = rs("deliverable") & ""
			strVersion = rs("version") & ""
			if trim(rs("revision") & "") <> "" then
				strVersion = strVersion & "," & rs("revision") & ""
			end if
			if trim(rs("pass") & "") <> "" then
				strVersion = strVersion & "," & rs("pass") & ""
			end if
			strPart = rs("PartNumber") & ""
			strModel = rs("ModelNumber") & ""
			strVendor = rs("Vendor") & ""
			strTestNotes = rs("DeveloperTestNotes") & ""
		end if
		rs.Close
	if trim(strProduct) ="" then
		Response.Write "Unable to find the requested record."
	else
		if strDefaultStatus <> "" then
			strTestNotes = ""
		end if 

%>

<form id="frmUpdateDevStatus" method="post" action="UpdateDevStatusSave.asp?">
<font size=3 face=verdana><b>
    <%if trim(request("TypeID")) = "3" then%>
	    Request Removal From Product
	<%elseif trim(strDefaultStatus) = "1" and trim(request("TypeID")) = "2" then%>
		Approve Release to Production for <%=strproduct%>
	<%elseif trim(strDefaultStatus) = "2" and trim(request("TypeID")) = "2" then%>
		Reject Release to Production for <%=strproduct%>
	<%elseif trim(request("TypeID")) = "2" then%>
		Update Release to Production status for <%=strproduct%>
	<%else%>
		Update Developer Status for <%=strproduct%>
	<%end if%>
	</b></font><BR><BR>
<table ID=TableDevStatus Name=TableDevStatus  border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
	<TR bgcolor=cornsilk>
		<TD width=150><b>Deliverable:&nbsp;&nbsp;&nbsp;&nbsp;</b></TD>
		<TD colspan=3 width=100%><%=strDeliverable%></TD>
	</TR>
	<TR bgcolor=cornsilk>
		<%if trim(strType)="1" then%>
			<TD width=150><b>HW,FW,Rev:</b></TD>
		<%else%>
			<TD width=150><b>Version:</b></TD>
		<%end if%>
		<TD width=40%><%=strVersion%></TD>
		<TD width=150><b>Vendor:</b></TD>
		<TD width=40%><%=strVendor%></TD>
	</TR>
	<%if trim(strType)="1" then%>
	<TR bgcolor=cornsilk>
		<TD width=150><b>Part:</b></TD>
		<TD width=40%><%=strPart%></TD>
		<TD width=150><b>Model:</b></TD>
		<TD width=40%><%=strModel%></TD>
	</TR>
	<%end if%>
	<TR bgcolor=cornsilk>

<%

	strSQL = "spGetDeveloperNotificationStatus " & clng(myIDarr(0)) & "," & clng(myIDarr(1))
	rs.Open strSQL,cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then

		Response.Write "<TD width=150><b>Status:</b></TD>"
		
		if trim(strType) <> "1" then 'Not Hardware
			strStatusAcceptText = "Approved"
			strStatusRejectText = "Rejected"
			strStatusReviewText = "Under&nbsp;Review"
		else	'Hardware
			if trim(request("TypeID")) = "2" then
				strStatusAcceptText = "Approved for Production"
				strStatusRejectText = "Not Approved for Production"
				strStatusReviewText = "Under&nbsp;Review"
			else
				strStatusAcceptText = "Approved for Test"
				strStatusRejectText = "Not Approved for Test"
				strStatusReviewText = "Under&nbsp;Review"
			end if
		end if
        
        if trim(request("TypeID")) = "3" then
			Response.Write "<TD width=""100%"" colspan=3 >Requesting Removal</TD>"
			Response.Write "<Select style=""display:none"" id=cboDevStatus name=cboDevStatus><Option selected value=2>Removal Requested</option></Select>"
		elseif trim(strDefaultStatus)= "1" and  trim(request("TypeID")) = "2" then
			Response.Write "<TD width=""100%"" colspan=3 >Approved for Production</TD>"
			Response.Write "<Select style=""display:none"" id=cboDevStatus name=cboDevStatus><Option selected value=1>Approved for Production</option></Select>"
			'Response.Write "<INPUT type=""hidden"" id=cboDevStatus name=cboDevStatus value=""1"">"
		elseif trim(strDefaultStatus)= "2" and  trim(request("TypeID")) = "2" then
			Response.Write "<TD width=100% colspan=3 >Not Approved for Production</TD>"
			Response.Write "<Select style=""display:none"" id=cboDevStatus name=cboDevStatus><Option selected value=2>Not Approved for Production</option></Select>"
'			Response.Write "<INPUT type=""hidden"" id=cboDevStatus name=cboDevStatus value=""2"">"
			DisplayNotesRequired = ""
		else
			Response.Write "<TD width=""100%"" colspan=3 >"
	
			Response.Write "<SELECT style=""width=100%"" id=cboDevStatus name=cboDevStatus LANGUAGE=javascript onclick=""return cboDevStatus_onclick()"">"
			if (trim(request("TypeID")) = "2" and trim(rs("DeveloperTestStatus")&"") = "1") or (trim(request("TypeID")) <> "2" and trim(rs("DeveloperNotificationStatus")&"") = "1") then
				Response.Write "<OPTION selected value=1>" & strStatusAcceptText & "</OPTION>"
			else
				Response.Write "<OPTION value=1>" & strStatusAcceptText & "</OPTION>"
			end if
			if (trim(request("TypeID")) = "2" and trim(rs("DeveloperTestStatus")&"") = "2") or (trim(request("TypeID")) <> "2" and trim(rs("DeveloperNotificationStatus")&"") = "2") then
				Response.Write "<OPTION selected value=2>" & strStatusRejectText & "</OPTION>"
				DisplayNotesRequired = ""
			else
				Response.Write "<OPTION value=2>" & strStatusRejectText & "</OPTION>"
			end if
			if (trim(request("TypeID")) = "2" and (trim(rs("DeveloperTestStatus")&"") = "0" or trim(rs("DeveloperTestStatus")&"") = "") ) or (trim(request("TypeID")) <> "2" and (trim(rs("DeveloperNotificationStatus")&"") = "0" or trim(rs("DeveloperNotificationStatus")&"") = "")) then
				Response.Write "<OPTION selected value=0>" & strStatusReviewText & "</OPTION>"
			else
				Response.Write "<OPTION value=0>" & strStatusReviewText & "</OPTION>"
			end if

			Response.Write "</SELECT>"
		end if
		Response.Write "</TD>"
		
	end if
	
	rs.Close

%>
	</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Comments:</b>
	&nbsp;<span style="Display:<%=DisplayNotesRequired%>" ID=RequireNotes><font color="#ff0000" size="1">*</font>
	<BR>
	<font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font>
	</td>
	<TD colspan=3 ><TEXTAREA style="WIDTH:100%; HEIGHT:80px" id=txtComments name=txtComments onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"><%=strTestNotes%></TEXTAREA></td>
</TR>

<%
    if trim(request("TypeID")) = "3" then
        response.Write "<TR bgcolor=cornsilk><TD nowrap valign=top><b>Products:</b></td><td colspan=3>"
        dim strID
        dim strIDarr
        for each strID in IDArray
            strIDarr = split(strID, "_")
    		rs.open "spGetProductDeliverableSummaryByID " & clng(trim(strIDarr(0))) & "," & clng(trim(strIDarr(1))),cn,adOpenForwardOnly
            if trim(rs("ID")) = trim(strIDarr(0)) then
                response.write rs("Product") 
            else
                response.write ", " & rs("Product") 
            end if
            rs.close
        next
        response.Write "</td></tr>"
    end if
%>

</Table>

<INPUT type="hidden" id=txtPDID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtStatueID name=txtStatusID value="<%=request("StatusID")%>">
<INPUT type="hidden" id=txtTypeID name=txtTypeID value="<%=request("TypeID")%>">
<INPUT type="hidden" id=txtStatusName name=txtStatusName value="">
</form>
	<%
		end if
	end if

	set rs=nothing
	cn.Close
	set cn=nothing

	%>
	
</BODY>
</HTML>

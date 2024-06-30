<%@ Language=VBScript %>


<HTML>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
    <!--  
    var origStatus = "";
    var origComment = "";

    $(document).ready(function () {

        origStatus = $("#cboStatus").val();
        origComment = $("#txtComments").val();
        origDate = $("#txtPilotDate").val();

        window.parent.frames["LowerWindow"].enableButton();
        $("#txtRedirect").val("PilotStatusMainPulsar.asp?ProdID=" + $("#txtProdID").val() + "&VersionID=" + $("#txtVersionID").val() + "&ProductDeliverableReleaseID=" + $("#txtProdDelRelID").val() + "&TodayPageSection=" + $("#txtTodayPageSection").val());
    });
   
    function cmdDate_onclick(FieldID) {
        var strID;
		
        strID = window.showModalDialog("../../mobilese/today/caldraw1.asp",frmStatus.txtPilotDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
        if (typeof(strID) != "undefined")
            frmStatus.txtPilotDate.value = strID;
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

    function window_onload() {
        //frmStatus.cboStatus.focus();
    }

    function cboStatus_onchange() {
        var strRequired = $("#txtDateRequired").val().indexOf("," + $("#cboStatus").val() + ",");
        var strShow = $("#txtDateShow").val().indexOf("," + $("#cboStatus").val() + ",");
        var strStatus = $("#cboStatus option:selected").text();
        var strCommentsRequired = $("#txtCommentsRequired").val().indexOf("," + $("#cboStatus").val() + ",");

        if (strShow != -1) {
            $("#DateRow").show();
            if (strRequired != -1)
                $("#DateStar").show();
            else
                $("#DateStar").hide();
        }
        else {
            $("#DateRow").hide();
            $("#txtPilotDate").val("");
            $("#DateStar").hide();
        }

        if (strCommentsRequired == -1)
            $("#CommentStar").hide();
        else
            $("#CommentStar").show();
    }

    function SwitchRelease(ProdID, VersionID, ProductDeliverableReleaseID) {
        var isModified = false;
        $("#txtKeepItOpen").val(true);

        if (origStatus != $("#cboStatus").val()) {
            isModified = true;
        }
        else if (origComment != $("#txtComments").val()) {
            isModified = true;
        }
        else if (origDate != $("#txtPilotDate").val()) {
            isModified = true;
        }
        
        if (isModified) {
            if (confirm("Do you want to save your changes for this Release?")) {
                $("#txtRedirect").val("PilotStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val());
                window.parent.frames["LowerWindow"].cmdOK_onclick();
            }
            else {
                document.location = "PilotStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val();
                window.parent.repositionParentWindow();
            }
        }
        else {
            document.location = "PilotStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val();
            window.parent.repositionParentWindow();
        }
    }
    //-->
</script>
</head>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

.DelTable TBODY TD{
	BORDER-TOP: gray thin solid;
}


</STYLE>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<%
	dim cn
	dim rs
	dim cm
	dim p
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim strPilotStatus
	dim strDate
	dim strDeliverable
	dim strID
	dim strPartNumber
	dim strProduct
	dim strVendor
	dim blnAdmin
	dim blnFound
	dim strStatusList
	dim strStatusText
	dim strQualStatus
	dim strComments
	dim strHW
	dim strFW
	dim strRev
	dim strModel
	dim strStatusSelected
	dim CurrentUserEmail
	dim strQCompleteSubject
	dim strQCompleteBody
	dim strFailSubject
	dim strFailBody
	dim strPMEmail
	dim strDevEmail
	dim strDevCenter
	dim strRows
	dim strPartnerID
	dim ProductDeliverableReleaseID
    dim ProductDeliverableReleaseName
	dim strReleaseLink
    	
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
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
	end if
	rs.Close
	
	if currentuserid=31 or currentuserid=8 then
		blnadmin=true
	else
		blnadmin=false
	end if
	
	strID=""
	strPilotStatus = ""
	strDate = ""
	strVendor=""
	strPartNumber=""
	strDeliverable = ""
	strProduct= ""
	strStatusList = ""
	strDevStatus = ""
	strComments = ""
	strHW = "&nbsp;"
	strFW = "&nbsp;"
	strRev = "&nbsp;"
	strModel = "&nbsp;"
	blnFound = false
	strStatusSelected = ""
	strFailSubject = ""
	strFailBody = ""
	strPMEmail = ""
	strDevEmail = ""
	strDevCenter = ""
	strRows = ""
	strPartnerID = ""
	strReleaseLink = ""
	ProductDeliverableReleaseID = 0
    ProductDeliverableReleaseName = ""
	
	if request("ProdID") = "" or request("VersionID") = "" then
		Response.Write "Not enough information supplied to process your request."
	else
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
            strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID and pdr.targeted = pd.targeted " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & request("ProdID") & " and pd.DeliverableVersionID= " & Request("VersionID") & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
        else 
            strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & request("ProdID") & " and pd.DeliverableVersionID= " & Request("VersionID")

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
            else
                if rs("ReleaseID") = intDefaultReleaseID and ProductDeliverableReleaseID = 0 then
                    strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                    ProductDeliverableReleaseID = rs("ID")
                    ProductDeliverableReleaseName = rs("Name")
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProdID") & "," & Request("VersionID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
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
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProdID") & "," & Request("VersionID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
                end if
                
                count = count + 1
                rs.MoveNext
            Loop
        end if
        rs.Close

		rs.Open "spGetPilotStatusRelease " & ProductDeliverableReleaseID, cn, adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strNumber = ""
		else
			blnFound = true
			
			strID = rs("ID") & ""
			strProduct = rs("Product") & ""
			strDevCenter = trim(rs("DevCenter") & "")
			strDeliverable = rs("DeliverableName") & ""
			strPilotStatus = rs("PilotStatusID") & ""
			strPartNumber = rs("PartNumber") & ""
			strModel = rs("ModelNumber") & "&nbsp;"
			strDate = rs("PilotDate") & ""
			if isdate(strDate) then
				strDate = formatdatetime(strDate,vbshortdate)
			end if
			strVendor=rs("Vendor") & ""
			strComments = rs("PilotNotes") & ""
			strHW = rs("Version") & "&nbsp;"
			strFW = rs("Revision") & "&nbsp;"
			strRev = rs("Pass") & "&nbsp;"
			strQualStatus = trim(rs("TestStatus") & "")
			strPartnerID = rs("PartnerID") & ""
			if strQualStatus = "" then
				strQualStatus = "Not Used"
			elseif strQualStatus = "Date" then
				strQualStatus = rs("TestDate") & "&nbsp;"
			end if
		end if
		rs.Close
		
		strStatusSelected = ""
		strDateShow = ""
		strDateRequired = ""
		strCommentsRequired = ""
		rs.Open "spListPilotStatus",cn,adOpenForwardOnly
		do while not rs.EOF
			strStatusText = rs("Name") & ""
			if trim(rs("ID")) = trim(strPilotStatus) then
				strStatusList = strStatusList & "<option selected value=""" & rs("ID") & """>" & strStatusText & "</option>"
				strStatusSelected = strStatusText
			else	
				strStatusList = strStatusList & "<option value=""" & rs("ID") & """>" & strStatusText & "</option>"
			end if
			if trim(rs("DateField")) & "" = "2" then
				strDateShow = strDateShow & "," & trim(rs("ID"))
			elseif	trim(rs("DateField")) & "" = "1" then
				strDateShow = strDateShow & "," & trim(rs("ID"))
				strDateRequired = strDateRequired & "," & trim(rs("ID"))
			end if
			if rs("CommentsRequired") then
				strCommentsRequired = strCommentsRequired & "," & trim(rs("ID"))
			end if
			
			rs.movenext
		loop
		strDateShow = strDateShow & ","
		strDateRequired = strDateRequired & ","
		strCommentsRequired = strCommentsRequired & ","
		rs.Close



		strDevEmail = ""
		rs.open "spGetDeliverableDeveloper " & clng(request("VersionID")),cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			if trim(rs("DeveloperEmail") & "") <> "" then
				strDevEmail = strDevEmail & ";" & rs("DeveloperEmail")
			end if
			if trim(rs("DevManagerEmail") & "") <> "" then
				strDevEmail = strDevEmail & ";" & rs("DevManagerEmail")
			end if
		end if
		rs.Close
		if trim(strDevEmail) = "" then
			strDevEmail = "max.yu@hp.com"
		else
			strDevEmail = mid(strDevEmail,2)
		end if


		strPMEmail = ""
		rs.Open "spListCommodityPMs4Version " & clng(request("VersionID")),cn,adOpenStatic
		do while not rs.EOF
			strPMEmail = strPMEmail & ";" & rs("Email")
			rs.MoveNext
		loop
		rs.Close
		if trim(strPMEmail) = "" then
			strPMEmail = "max.yu@hp.com"
		else
			strPMEmail = mid(strPMEmail,2)
		end if
		
		
	end if
	
	if 	blnFound then
%>



<font face=verdana size=2><b>
<label ID="lblTitle"><%=strVendor%>&nbsp;<%=strDeliverable%> on <%=strProduct%></label></b></font>
<br />
<span style="font-family: Verdana; font-size: 9pt;"><%=strReleaseLink %></span>
<% 

	dim strVersion
	
	strVersion = ""
	if trim(strHW) <> "&nbsp;" then
		strVersion =  strHW
	end if
	if trim(strFW) <> "&nbsp;" then
		strVersion = strVersion & "," & strFW
	end if
	if trim(strRev) <> "&nbsp;" then
		strVersion = strVersion & "," & strRev
	end if

%>

<form id="frmStatus" method="post" action="PilotStatusSavePulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td valign=top width=120 nowrap><b>Qual.&nbsp;Status:</b>&nbsp;</td>
		<td>
			<%=strQualStatus%>
		</td>
		<td valign=top width=60 nowrap><b>HW&nbsp;Ver:</b>&nbsp;</td>
		<td><%=strHW%></td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Model:</b>&nbsp;</td>
		<td>
			<%=strModel%>
		</td>
		<td valign=top width=60 nowrap><b>FW&nbsp;Ver:</b>&nbsp;</td>
		<td><%=strFW%></td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Part&nbsp;Number:</b>&nbsp;</td>
		<td>
			<%=strPartNumber%>&nbsp;
		</td>
		<td valign=top width=60 nowrap><b>Rev:</b>&nbsp;</td>
		<td><%=strRev%></td>
	</tr>



	<tr>
		<td valign=top width=120 nowrap><b>Pilot&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan=3>
			<SELECT style="width:160" id=cboStatus name=cboStatus language=javascript onchange="cboStatus_onchange();">
			<%=strStatusList%>
			</SELECT>
		</td>
	</tr>
	
	<%if instr(strDateShow,"," & trim(strPilotStatus) & ",") > 0 then %>
		<tr ID=DateRow>
	<%else%>
		<tr ID=DateRow style="display:none">
	<%end if%>
		<td valign=top width=120 nowrap><b>Date:</b>
		<%if instr(strDateRequired,"," & trim(strPilotStatus) & ",") > 0 then %>
			&nbsp;<font color="red" size="1" ID=DateStar>*</font>&nbsp;
		<%else%>
			&nbsp;<font style="Display:none" color="red" size="1" ID=DateStar>*</font>&nbsp;
		<%end if%>
		</td>
		<td colspan=3>
			<INPUT type="text" id=txtPilotDate name=txtPilotDate value="<%=strDate%>">&nbsp;<a href="javascript: cmdDate_onclick()"><img ID="picTarget" SRC="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Comments:</b>
		<%if instr(strCommentsRequired,"," & trim(strPilotStatus) & ",") > 0 then %>
			<font color="red" size="1" ID=CommentStar>*</font>&nbsp;
		<%else%>
			<font style="Display:none" color="red" size="1" ID=CommentStar>*</font>&nbsp;
		<%end if%>
		
		</td>
		<td colspan=7>	<TEXTAREA rows=3 style="WIDTH:100%" id=txtComments name=txtComments><%=strComments%></TEXTAREA>
			
			
		</td>
	</tr>
</table>


<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=strID%>">
<INPUT style="Display:none" type="text" id=txtUserID name=txtUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=CurrentDomain & "_" & CurrentUser%>">
<INPUT type="hidden" id=txtUserEmail name=txtUserEmail value="<%=CurrentUserEmail%>">
<INPUT type="hidden" id=txtQualStatus name=txtQualStatus value="<%=strQualStatus%>">
<INPUT type="hidden" id=txtStatusText name=txtStatusText value="">
<INPUT type="hidden" id=txtVendor name=txtVendor value="<%=strVendor%>">
<INPUT type="hidden" id=txtDeliverable name=txtDeliverable value="<%=strDeliverable%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtVersion name=txtVersion value="<%=strVersion%>">
<INPUT type="hidden" id=txtModel name=txtModel value="<%=strModel%>">
<INPUT type="hidden" id=txtPartNumber name=txtPartNumber value="<%=strPartNumber%>">
<INPUT type="hidden" id=txtStatusLoaded name=txtStatusLoaded value="<%=trim(strPilotStatus)%>">
<INPUT type="hidden" id=txtProduct name=txtProduct value="<%=strProduct%>">
<INPUT type="hidden" id=txtPMEmail name=txtPMEmail value="<%=strPMEmail%>">
<INPUT type="hidden" id=txtDevEmail name=txtDevEmail value="<%=strDevEmail%>">
<INPUT type="hidden" id=txtPartnerID name=txtPartnerID value="<%=strPartnerID%>">
<input type="hidden" id="txtRedirect" name="txtRedirect" value="" />
<input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="false" />
<input type="hidden" id="txtProductDeliverableID" name="txtProductDeliverableID" value="<%=Request("ProductDeliverableID")%>" />
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=Request("TodayPageSection")%>" />
<input type="hidden" id="txtProdDelRelID" name="txtProdDelRelID" value="<%=ProductDeliverableReleaseID%>" />
<input type="hidden" id="txtReleaseName" name="txtReleaseName" value="<%=ProductDeliverableReleaseName%>" />
<input type="hidden" id="txtProdID" name="txtProdID" value="<%=request("ProdID")%>" />
<INPUT type="hidden" id=txtDateShow name=txtDateShow value="<%=strDateShow%>">
<INPUT type="hidden" id=txtDateRequired name=txtDateRequired value="<%=strDateRequired%>">
<INPUT type="hidden" id=txtCommentsRequired name=txtCommentsRequired value="<%=strCommentsRequired%>">
<input type="hidden" id="txtShowOnlyTargetedRelease" name="txtShowOnlyTargetedRelease" value="<%=Request("ShowOnlyTargetedRelease")%>" />
</form>
<%end if

	cn.Close
	set cn = nothing
	set rs = nothing


%>

</BODY>
</HTML>

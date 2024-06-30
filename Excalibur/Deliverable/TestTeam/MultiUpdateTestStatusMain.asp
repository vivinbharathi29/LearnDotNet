<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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

function cboIntegrationStatus_onclick() {
	if (frmMain.cboIntegrationStatus.selectedIndex==2 || frmMain.cboIntegrationStatus.selectedIndex==3)
		RequireIntegrationNotes.style.display="";
	else
		RequireIntegrationNotes.style.display="none";
}

function cboODMStatus_onclick() {
	if (frmMain.cboODMStatus.selectedIndex==2 || frmMain.cboODMStatus.selectedIndex==3)
		RequireODMNotes.style.display="";
	else
		RequireODMNotes.style.display="none";
}

function cboWWANStatus_onclick() {
	if (frmMain.cboWWANStatus.selectedIndex==2 || frmMain.cboWWANStatus.selectedIndex==3)
		RequireWWANNotes.style.display="";
	else
		RequireWWANNotes.style.display="none";
}

function cboDEVStatus_onclick() {
	if (frmMain.cboDEVStatus.selectedIndex==2 || frmMain.cboDEVStatus.selectedIndex==3)
		RequireDEVNotes.style.display="";
	else
		RequireDEVNotes.style.display="none";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
.VersionTable TD
{
	FONT-SIZE: xx-small;
    COLOR: black;
    FONT-FAMILY: Verdana
}
</STYLE>
<BODY bgcolor=ivory>
<LINK href="../../style/wizard style.css" type=text/css rel=stylesheet >
	
<%   
	dim cn 
	dim rs
	dim strSQL
	dim strProductID
  	dim strProductPartner
  	dim strProductName
  	dim CurrentUserPartner
  	dim blnAdmin
  	dim TestStatusArray
  	dim strCommodityPM
  	
  	TestStatusArray = split("TBD,Passed,Failed,Blocked,Watch,N/A",",")
  	
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
		blnCommodityPM = rs("CommodityPM") 
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
		CurrentUserPartner = 0
		blnCommodityPM = false
	end if
	rs.Close
	
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	if blnCommodityPM then
		blnAdmin = true
	else
		blnAdmin = false
	end if
	
    Dim ProductID
    ProductID = Request.QueryString("ProductID")
    If (ProductID = "" Or Not IsNumeric(ProductID)) Then
        ProductID = 0
    end if

	rs.Open "spGetProductVersion " & clng(ProductID),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strProductID = "0"
  		strProductPartner = "0"
  		strProductName = ""
	else
		strProductID = rs("ID") & ""
  		strProductPartner = rs("PartnerID") & ""
  		strProductName = rs("DOTSname") & ""
	end if
	rs.Close

	'Verify Access is OK
    if ProductID > 0 then
	    if trim(CurrentUserPartner) <> "1" then
		    if trim(strProductPartner) <> trim(CurrentUserPartner) or trim(CurrentUserPartner) = "0" then
			    set rs = nothing
			    set cn=nothing
				
			    'Response.Redirect "../../NoAccess.asp?Level=1"
		    end if
	    end if
    end if

    rs.Open "spGetHardwareTeamAccessList " & CurrentUserID & "," & clng(ProductID),cn,adOpenStatic
        do while not rs.EOF
		    if rs("HWTeam") = "Commodity" and rs("Products") > 0 then
			    blnAdmin = true
                exit do
		    end if
		    rs.MoveNext
	    loop
	rs.Close

	if blnAdmin  then
		blnSETestLead = true
		blnODMTestLead = true
		blnWWANTestLead = true
		blnDEVTestLead = true
	else	
		blnSETestLead = false
		blnODMTestLead = false
		blnWWANTestLead = false
		blnDEVTestLead = false
		
		rs.Open "spGetTestLeadsAll",cn,adOpenStatic
		do while not rs.EOF
			if trim(CurrentUserID) = trim(rs("ID")) then
				if rs("Role") = "SE Test Lead" then
					blnSETestLead = true
				elseif rs("Role") = "ODM Test Lead" then
					blnODMTestLead = true
				elseif rs("Role") = "WWAN Test Lead" then
					blnWWANTestLead = true
    			elseif rs("Role") = "DEV Test Lead" then
					blnWWANTestLead = true
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
	end if

	if blnSETestLead = false and blnODMTestLead = false and blnWWANTestLead=false and blnDEVTestLead=false and blnAdmin=false then
			set rs = nothing
			set cn=nothing
				
			'Response.Redirect "../../NoAccess.asp?Level=1"
	end if    
%>

<font size=3 face=verdana><b>Update <%= strProductName%> Test Lead Status</b><BR><BR></font>
<form ID=frmMain action="MultiUpdateTestStatusSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>


<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<%if blnSETestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>SE&nbsp;Test&nbsp;Status:</b>&nbsp;&nbsp;&nbsp;</TD>
	<TD>
		<SELECT style="width:150" id=cboIntegrationStatus name=cboIntegrationStatus LANGUAGE=javascript onchange="return cboIntegrationStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<OPTION value=1>Passed</OPTION>
			<OPTION value=2>Failed</OPTION>
			<OPTION value=3>Blocked</OPTION>
		</SELECT>
	</TD>
</TR>
<%if blnSETestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>SE&nbsp;Test&nbsp;Notes:</b>&nbsp;<span style="Display:none" ID=RequireIntegrationNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD width=100%>
	<TEXTAREA rows=4 style="width:100%" id=txtIntegrationNotes name=txtIntegrationNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"></TEXTAREA>
	</TD>
</TR>

<%if blnODMTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>ODM&nbsp;Test&nbsp;Status:</b>&nbsp;&nbsp;&nbsp;</TD>
	<TD>
		<SELECT style="width:150" id=cboODMStatus name=cboODMStatus LANGUAGE=javascript onchange="return cboODMStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<OPTION value=1>Passed</OPTION>
			<OPTION value=2>Failed</OPTION>
			<OPTION value=3>Blocked</OPTION>
		</SELECT>
	</TD>
</TR>
<%if blnODMTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>ODM&nbsp;Test&nbsp;Notes:</b>&nbsp;<span style="Display:none" ID=RequireODMNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD width=100%>
	<TEXTAREA rows=4 style="width:100%" id=txtODMNotes name=txtODMNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"></TEXTAREA>
	</TD>
</TR>

<%if blnDEVTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>DEV&nbsp;Test&nbsp;Status:</b>&nbsp;&nbsp;&nbsp;</TD>
	<TD>
		<SELECT style="width:150" id=cboDEVStatus name=cboDEVStatus LANGUAGE=javascript onchange="return cboDEVStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<OPTION value=1>Passed</OPTION>
			<OPTION value=2>Failed</OPTION>
			<OPTION value=3>Blocked</OPTION>
		</SELECT>
	</TD>
</TR>
<%if blnDEVTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>DEV&nbsp;Test&nbsp;Notes:</b>&nbsp;<span style="Display:none" ID=RequireDEVNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD width=100%>
	<TEXTAREA rows=4 style="width:100%" id=txtDEVNotes name=txtDEVNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"></TEXTAREA>
	</TD>
</TR>

<%if blnWWANTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>COMM&nbsp;Test&nbsp;Status:</b>&nbsp;&nbsp;</TD>
	<TD>
		<SELECT style="width:150" id=cboWWANStatus name=cboWWANStatus LANGUAGE=javascript onchange="return cboWWANStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<OPTION value=1>Passed</OPTION>
			<OPTION value=2>Failed</OPTION>
			<OPTION value=3>Blocked</OPTION>
			<OPTION value=4>Watch</OPTION>
			<OPTION value=5>N/A</OPTION>
		</SELECT>
	</TD>
</TR>
<%if blnWWANTestLead then%>
	<TR bgcolor=cornsilk>
<%else%>
	<TR style=display:none bgcolor=cornsilk>
<%end if%>
	<TD nowrap valign=top><b>COMM&nbsp;Test&nbsp;Notes:</b>&nbsp;<span style="Display:none" ID=RequireWWANNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD width=100%>
	<TEXTAREA rows=4 style="width:100%" id=txtWWANNotes name=txtWWANNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"></TEXTAREA>
	</TD>
</TR>

</table>


<%
    function ScrubSQL(strWords) 

        dim badChars 
        dim newChars 
        dim i

    
        strWords=replace(strWords,"'","''")
    
        badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
        newChars = strWords 
    
        for i = 0 to uBound(badChars) 
          newChars = replace(newChars, badChars(i), "") 
        next 
    
        ScrubSQL = newChars 
  
    end function 
	
    Dim IDList
    IDList = Request.QueryString("IDList")
            
    dim pdids, pdrids
    pdids = ""
    pdrids = ""

    if InStr(IDList,"_") > 0 then
        dim arr
        arr = Split(IDList,",")
        dim arrID
        if UBound(arr) > 0 then 
            For i = 0 to uBound(arr)                        
                arrID = Split(arr(i),"_")
                if arrID(1) > 0 then
                    if pdrids <> "" then
                        pdrids = pdrids & ","
                    end if
                    pdrids = pdrids & arrID(1) 
                else 
                    if pdids <> "" then
                        pdids = pdids & ","
                    end if
                    pdids = pdids & arrID(0)                   
                end if                
            Next
        else 
            arrID = Split(arr(0),"_") 
            if arrID(1) > 0 then
                pdrids = arrID(1) 
            else 
                pdids = arrID(0)               
            end if        
        end if  
    
        if pdids = "" then
            pdids = "0"
        end if

        if pdrids = "" then
            pdrids = "0"
        end if

        strSQL = "SELECT v.id as versionid, pv.id as productid, pd.id as productdeliverableid, pv.Dotsname as product, v.deliverablename, v.partnumber, v.modelnumber, v.version, v.revision, v.pass, vd.name as Vendor, endoflifedate as eoadate, serviceeoadate, pd.integrationteststatus, pd.odmteststatus, pd.wwanteststatus, pd.DeveloperTestStatus, pd.integrationtestnotes, pd.odmtestnotes, pd.wwantestnotes, ProductDeliverableReleaseID = 0 " & _
		         "FROM product_deliverable pd with (NOLOCK) inner join " & _
                 "DeliverableVersion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _
                 "vendor vd with (NOLOCK) on v.vendorid = vd.id inner join " & _
                 "productversion pv with (NOLOCK) on pd.ProductVersionID = pv.ID " & _
		         "WHERE pd.id in (" & scrubsql(pdids) & ") "

        strSQL = strSQL & "Union select v.id as versionid, pv.id as productid, pd.id as productdeliverableid, pv.Dotsname as product, v.deliverablename, v.partnumber, v.modelnumber, v.version, v.revision, v.pass, vd.name as Vendor, endoflifedate as eoadate, serviceeoadate, pd.integrationteststatus, pd.odmteststatus, pd.wwanteststatus, pd.DeveloperTestStatus, pd.integrationtestnotes, pd.odmtestnotes, pd.wwantestnotes, ProductDeliverableReleaseID = pdr.id " & _
                 "from product_deliverable_release pdr inner join " & _
                 "product_deliverable pd on pd.id = pdr.productdeliverableid inner join " & _
                 "deliverableversion v on v.id = pd.DeliverableVersionID inner join " & _
                 "vendor vd on vd.id = v.vendorid inner join " & _
                 "productversion pv on pv.id = pd.productversionid inner join " & _
                 "productversionrelease pvr on pvr.id = pdr.releaseid " & _
                 "where pdr.id in (" & scrubsql(pdrids) & ") " & _
                 "ORDER BY v.deliverablename, v.id desc"
    else 
        
        strSQL = "SELECT v.id as versionid, pv.id as productid, pd.id as productdeliverableid, pv.Dotsname as product, v.deliverablename, v.partnumber, v.modelnumber, v.version, v.revision, v.pass, vd.name as Vendor, endoflifedate as eoadate, serviceeoadate, pd.integrationteststatus, pd.odmteststatus, pd.wwanteststatus, pd.DeveloperTestStatus, pd.integrationtestnotes, pd.odmtestnotes, pd.wwantestnotes, ProductDeliverableReleaseID = 0" & _
			     "FROM product_deliverable pd with (NOLOCK) inner join " & _
                 "DeliverableVersion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _  
                 "vendor vd with (NOLOCK) on v.vendorid = vd.id inner join " & _ 
                 "productversion pv with (NOLOCK) on pv.id = pd.productversionid " & _
			     "WHERE pd.productversionid = " & clng(request("ProductID")) & " " & _
			     "and v.id in (" & scrubsql(request("IDList")) & ") and pv.fusionrequirements <> 1 "

        strSQL = strSQL & "Union SELECT v.id as versionid, pv.id as productid, pd.id as productdeliverableid, pv.Dotsname + ' (' + pvr.name + ')' as product, v.deliverablename, v.partnumber, v.modelnumber, v.version, v.revision, v.pass, vd.name as Vendor, endoflifedate as eoadate, serviceeoadate, pdr.integrationteststatus, pdr.odmteststatus, pdr.wwanteststatus, pdr.DeveloperTestStatus, pdr.integrationtestnotes, pdr.odmtestnotes, pdr.wwantestnotes, pdr.ID as ProductDeliverableReleaseID " & _
			     "FROM product_deliverable pd with (NOLOCK) inner join " & _
                 "DeliverableVersion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _
                 "vendor vd with (NOLOCK) on v.vendorid = vd.id inner join " & _
                 "productversion pv with (NOLOCK) on pv.id = pd.productversionid inner join " & _
                 "product_deliverable_release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.id inner join " & _
                 "productversionrelease pvr with (NOLOCK) on pvr.id = pdr.ReleaseID " & _
			     "WHERE pd.productversionid = " & clng(request("ProductID")) & " " & _
			     "and v.id in (" & scrubsql(request("IDList")) & ") and pv.fusionrequirements = 1 " & _
			     "ORDER BY v.deliverablename, v.id desc;"   
    end if    

	rs.Open strSQL,cn,adOpenForwardOnly        
    
	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>No Deliverables Selected.</font>"
	else
		Response.Write "<BR><font size=2 face=verdana><b>Deliverables to Update:</b><BR></font>"
		Response.Write "<table class=VersionTable  ID=VersionTable border=1 width=""100%"" bordercolor=tan cellspacing=0 cellpadding=2>"
		Response.Write "<TR bgcolor=cornsilk><TD></TD><TD><b>Product</b></TD><TD><b>Deliverable</b></TD><TD><b>Available&nbsp;Until</b></TD><TD><b>Version</b></TD><TD><b>Vendor</b></TD><TD><b>Model</b></TD><TD><b>Part&nbsp;Number</b></TD>"
		
        if blnSETestLead then
			Response.Write "<TD><b>SE&nbsp;Status</b></TD>"
		end if
		
        if blnODMTestLead then
			Response.Write "<TD><b>ODM&nbsp;Status</b></TD>"
		end if

        if blnWWANTestLead then
			Response.Write "<TD><b>WWAN&nbsp;Status</b></TD>"
		end if
			
        if blnDEVTestLead then
			Response.Write "<TD><b>DEV&nbsp;Status</b></TD>"
		end if

        Response.Write "</TR>"
		do while not rs.EOF
		
			strVersion = rs("Version") & ""
			if trim(rs("Revision") & "") <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if trim(rs("Pass") & "") <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if

			if trim(request("TypeID"))="2" then
				strEOL = rs("ServiceEOADate") & ""
			else
				strEOL = rs("eoadate") & ""
			end if
			
			Response.Write "<TR>"
			Response.Write "<TD><INPUT type=""checkbox"" checked id=""lstID"" name=""lstID"" style=""WIDTH: 14px; HEIGHT: 14px"" size=""14"" value=""" & rs("VersionID") & "_" & rs("ProductID") & "_" & rs("productdeliverableid") & "_" & rs("ProductDeliverableReleaseID") & """></TD>"
            Response.Write "<TD>" & rs("product") & "</TD>"
			Response.Write "<TD>" & rs("DeliverableName") & "</TD>"
			Response.Write "<TD>" & strEOL & "&nbsp;</TD>"
			Response.Write "<TD>" & strVersion & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Vendor") & "</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "</TD>"
			Response.Write "<TD>" & rs("PartNumber") & "</TD>"
		
			if blnSETestLead then
				if isnull(rs("IntegrationTestStatus")) then
					Response.Write "<TD>&nbsp;</TD>"
				else
					Response.Write "<TD>" & TestStatusArray(rs("IntegrationTestStatus")) & "&nbsp;</TD>"
				end if
			end if
			if blnODMTestLead then
				if isnull(rs("ODMTestStatus")) then
					Response.Write "<TD>&nbsp;</TD>"
				else
					Response.Write "<TD>" & TestStatusArray(rs("ODMTestStatus")) & "&nbsp;</TD>"
				end if
			end if
			if blnWWANTestLead then
				if isnull(rs("WWANTestStatus")) then
					Response.Write "<TD>&nbsp;</TD>"
				else
					Response.Write "<TD>" & TestStatusArray(rs("WWANTestStatus")) & "&nbsp;</TD>"
				end if
			end if
    		if blnDEVTestLead then
				if isnull(rs("DeveloperTestStatus")) then
					Response.Write "<TD>&nbsp;</TD>"
				else
					Response.Write "<TD>" & TestStatusArray(rs("DeveloperTestStatus")) & "&nbsp;</TD>"
				end if
			end if


			rs.MoveNext
		loop
		Response.Write "</TR>"
		Response.Write "</table>"	
	end if

	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing
%>

<INPUT type="hidden" id=txtIDList name=txtIDList value="<%=request("IDList")%>">
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtSETestLead name=txtSETestLead value="<%=lcase(trim(blnSETestLead))%>">
<INPUT type="hidden" id=txtODMTestLead name=txtODMTestLead value="<%=lcase(trim(blnODMTestLead))%>">
<INPUT type="hidden" id=txtWWANTestLead name=txtWWANTestLead value="<%=lcase(trim(blnWWANTestLead))%>">
<INPUT type="hidden" id=txtDEVTestLead name=txtDEVTestLead value="<%=lcase(trim(blnDEVTestLead))%>">
<input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=request("TodayPageSection")%>" />
<input type="hidden" id="txtIndex" name="txtIndex" value="<%=request("Index")%>" />
</form>

</BODY>
</HTML>

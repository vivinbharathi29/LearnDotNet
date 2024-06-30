<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

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
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					break;
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



function window_onload() {
	frmStatus.txtComments.focus();
	//SupportRow.style.display="none";
}



function cboStatus_onchange(){	
}


//-->
</SCRIPT>
</HEAD>
<STYLE>
	TD
	{
	VERTICAL-ALIGN: top
	}
	
</STYLE>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

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
	
	dim cn
	dim rs
	dim p
	dim cm
	dim CurrentUser
	dim CurrentUserPartnerID
	dim strDeliverableList
	dim strSQl
	
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
		CurrentUserPartnerID = rs("PartnerID")
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	dim pdids, pdrids
    pdids = ""
    pdrids = ""

    if InStr(request("txtMultiID"),"_") > 0 then
        
        dim arr
        arr = Split(Request("txtMultiID"),",")
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
    end if

    if pdids = "" then
        pdids = "0"
    end if

     if pdrids = "" then
        pdrids = "0"
    end if

	if trim(Request("txtMultiID")) <> "" then
		 strSQl = "SELECT pd.TargetNotes, pd.DeveloperTestNotes, pd.RiskRelease, pd.AccessoryNotes, pd.PilotNotes, v.Active AS EOL, at.Name AS AccessoryStatus, v.EndOfLifeDate AS EOLDate, c.Commodity, pv.DOTSName AS Product, " & _
				  "pd.AccessoryStatusID, pd.AccessoryDate, pt.ID AS PilotStatusID, pt.Name AS PilotStatus, pd.PilotDate, t.ID AS StatusID, t.Status, pd.TestDate, pd.ID, r.Name, v.ID AS VersionID, v.Version, v.Revision, v.Pass, " & _
				  "v.ModelNumber, v.PartNumber, vd.Name AS vendor, v.Location, [productdeliverablereleaseid] = 0 " & _
				  "FROM Product_Deliverable AS pd WITH (NOLOCK) " & _
				  "INNER JOIN DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID " & _
				  "INNER JOIN DeliverableRoot AS r WITH (NOLOCK) ON r.ID = v.DeliverableRootID " & _
				  "INNER JOIN Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID " & _
				  "INNER JOIN PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID " & _
				  "INNER JOIN DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID " & _
				  "INNER JOIN ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _
				  "LEFT OUTER JOIN TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID " & _
				  "LEFT OUTER JOIN AccessoryStatus AS at WITH (NOLOCK) ON at.ID = pd.AccessoryStatusID " & _
				  "WHERE pd.ID in(" & scrubsql(pdids) & ") "
		
        if CurrentUserPartnerID <> 1 then
		    strSQl = strSQl & " AND pv.PartnerID = " & CurrentUserPartnerID & " " 
		end if

	    strSQl = strSQl & "Union " & _
                 "SELECT pdr.TargetNotes, pdr.DeveloperTestNotes, pdr.RiskRelease, pdr.AccessoryNotes, pdr.PilotNotes, v.Active AS EOL, at.Name AS AccessoryStatus, v.EndOfLifeDate AS EOLDate, c.Commodity, pv.DOTSName + ' (' + pvr.name + ')' AS Product, " & _
				 "pdr.AccessoryStatusID, pdr.AccessoryDate, pt.ID AS PilotStatusID, pt.Name AS PilotStatus, pdr.PilotDate, t.ID AS StatusID, t.Status, pdr.TestDate, pd.ID, r.Name, v.ID AS VersionID, v.Version, v.Revision, v.Pass, " & _
				 "v.ModelNumber, v.PartNumber, vd.Name AS vendor, v.Location, pdr.id as productdeliverablereleaseid " & _
				 "FROM Product_Deliverable AS pd WITH (NOLOCK) " & _
				 "INNER JOIN DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID " & _
				 "INNER JOIN DeliverableRoot AS r WITH (NOLOCK) ON r.ID = v.DeliverableRootID " & _
				 "INNER JOIN Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID " & _
    			 "INNER JOIN DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID " & _
				 "INNER JOIN ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _
                 "INNER JOIN Product_Deliverable_Release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.id " & _
                 "INNER JOIN ProductVersionRelease pvr with (NOLOCK) on pvr.id = pdr.releaseid " & _
                 "INNER JOIN PilotStatus AS pt WITH (NOLOCK) ON pdr.PilotStatusID = pt.ID " & _
				 "LEFT OUTER JOIN TestStatus AS t WITH (NOLOCK) ON pdr.TestStatusID = t.ID " & _
				 "LEFT OUTER JOIN AccessoryStatus AS at WITH (NOLOCK) ON at.ID = pdr.AccessoryStatusID " & _
				 "WHERE pdr.ID in(" & scrubsql(pdrids) & ") "


		if CurrentUserPartnerID <> 1 then
		    strSQL = strSQL & " AND pv.PartnerID = " & CurrentUserPartnerID & " " 
		end if

		strSQl = strSQL &  " ORDER BY r.name, vd.name, v.id desc;"

		rs.open strSQl, cn,adOpenForwardOnly
		do while not rs.EOF
			strVersion = rs("Version") & ""
			if trim(rs("Revision") & "") <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if	
			if trim(rs("Pass") & "") <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			
			strDeliverableList = strDeliverableList & "<TR><TD><INPUT checked type=""checkbox"" style=""Width:16;Height:16"" TestStatus=" & TestStatusID & " id=txtMultiID name=txtMultiID value=""2_" & rs("ID") & "_" & rs("productdeliverablereleaseid") & """></TD>"
			strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Product") & "&nbsp;</td>"
			strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Vendor") & "</td>"
			strDeliverableList = strDeliverableList & "<TD>" & rs("Name") & "&nbsp;[" & strVersion & "]</td>"
			strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("ModelNumber") & "&nbsp;</td>"
			strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("PartNumber") & "&nbsp;</td>"
			strDeliverableList = strDeliverableList & "<TD >" & rs("DeveloperTestNotes") & "&nbsp;</td>"

			rs.MoveNext
		loop
		rs.Close
	
	end if
        
    
	if strDeliverableList="" then
		Response.Write "Not enough information supplied to process your request."
	else
		



%>

<font face=verdana size=3>
<b>Developer&nbsp;Approval&nbsp;-&nbsp;Release&nbsp;to&nbsp;Production<BR></b>

	<form id="frmStatus" method="post" action="MultiDevApprovalSave.asp">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td valign=top width=10 nowrap><b>Status:&nbsp;</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td width=100%>
			<%if request("NewValue") = "1" then%>
				Approved for Release to Production
			<%else%>
				Not Approved for Release to Production
			<%end if%>	
		</TD>
	</TR>
	<tr>
		
		<%if request("NewValue") = "2" then%>
			<td valign=top width=10 nowrap><b>Comments:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<%else%>
			<td valign=top width=10 nowrap><b>Comments:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<%end if%>
		
		<td width=100%>
		<TEXTAREA rows=5 style="width: 100%" id=txtComments name=txtComments></TEXTAREA>
		
		</TD>
	</TR>
	<TR>
		<TD colspan=2 valign=top><b>Deliverables Selected:</b><BR>
			<TABLE bgcolor=white width=100% border=1>
			<THEAD>
			<TR bgcolor=gainsboro>
			<TD width=10>&nbsp;</td>
				<TD><b>Product</b></TD>
				<TD><b>Vendor</b></TD>
				<TD width=50%><b>Deliverable</b></TD>
				<TD><b>Model</b></TD>
				<TD><b>Part</b></TD>
				<TD width=50%><b>Dev.&nbsp;Comments</b></TD></TR>
			</THEAD>
				<%=strDeliverableList%>
			</TABLE>
		</TD>
	</TR>
</table>

<INPUT type="hidden" id=NewValue name=NewValue value="<%=request("NewValue")%>">
</form>

<%
	end if

	set rs = nothing
	cn.Close
	set cn = nothing
%>
</BODY>
</HTML>

<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<STYLE>
Body
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
td
{
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana
}
</STYLE>
<BODY  bgcolor=Ivory>
<%

	dim strDelName
	dim strprodname
	dim strRootID
	dim strProductID
	dim cn
	dim rs
	dim i
	dim strCategoryID
	dim blnLoadFailed
	dim strVersionsLoaded
	dim strVersionList
	dim VersionArray
	dim OutputArray
	dim strHeaderRow
	dim strVersion
	dim strModelNumber
	dim strPartNumber
	dim strVendor
	dim strSelectedProductIDs
	dim strProductsLoaded
	dim strLocation
	dim blnWorkflowComplete
		
	strSelectedProductIDs = ""
	strProductsLoaded = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	blnLoadFailed = false

	set rs = server.CreateObject("ADODB.recordset")

	if request("RootID") = "" or trim(request("RootID")) = "0" then
		rs.Open "spGetRootID " & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strRootID = 0
			blnLoadFailed = true
		else
			strRootID = rs("ID")
		end if
		rs.Close
	else
		strRootID = request("RootID")
	end if

	rs.Open "spGetDeliverableVersionProperties " & clng(request("VersionID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strDelName = ""
		strType = ""
		strCategoryID = ""
		blnLoadFailed = true
		strVersion = ""
		strModelNumber = ""
		strPartNumber = ""
		strVendor = ""
		strLocation = ""
	else
		strDelName = rs("name") & ""
		strType = rs("TypeID") & ""
		strCategoryID = rs("CategoryID") & ""
		strActive = rs("Active") & ""
		strVersion = rs("Version") & ""
		if rs("Revision") & "" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass") & "" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if
		strModelNumber = rs("ModelNUmber") & "&nbsp;"
		strPartNumber = rs("PartNumber") & "&nbsp;"
		strVendor = rs("Vendor") & "&nbsp;"
		strLocation = rs("Location") & ""
	end if
	
	rs.Close

	if instr(strLocation,"Workflow Complete") > 0 then
		blnWorkflowComplete = true
	else
		blnWorkflowComplete = false
	end if

	if blnLoadFailed then
		Response.Write "<BR><BR><font size=2 face=verdana>Not enough information supplied to display this page.</font>"
	else

		'Load Products Supported by Version
		rs.open "spGetProductsForVersion " & clng(request("VersionID")),cn,adOpenForwardOnly
		strSelectedProductIDs = ""
		do while not rs.EOF
			strSelectedProductIDs = strSelectedProductIDs  & rs("ID") & ","
			rs.MoveNext
		loop

		rs.Close	
	
		'get Vendor Buckets
		rs.Open "spListVendors4Root " & strRootID ,cn,adOpenForwardOnly
		strVersionList = ""
		strHeaderRow =  "<TR><TD><b>Support</b></TD><TD><b>Product</b></TD><TD><b>Total</b></TD>"
		do while not rs.EOF
			strHeaderRow = strHeaderRow &  "<TD><b>" & rs("Name") & "</b></TD>"	
			strVendorList = strVendorList & "," & rs("ID")
			rs.MoveNext
		loop
		rs.Close
		strHeaderRow = strHeaderRow &  "</TR>"
		
		if strVendorList <> "" then
			strVendorList = mid(strVendorList,2)
		end if

		VendorArray = split(strVendorList,",")
		OutputArray = split(strVendorList,",")

		rs.Open "spListDeliverableUsageByVendor " & clng(strRootID) ,cn,adOpenForwardOnly

		if rs.EOF and rs.BOF then
			Response.Write strRootID
			Response.Write "<BR><BR><font size=2 face=verdana>No products found.</font>"
		else
%>

<h3>Select Supported Products<h3>
<!--<h5></h5>-->
<TABLE bgcolor=cornsilk cellpadding=2 cellspacing=0 bordercolor=tan border=1><TR><TD><b>Deliverable</b></TD><TD><b>Vendor</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model Number</b></TD><TD><b>Part Number</b></TD></TR><TR><TD><%=strDelName%></TD><TD><%=strVendor%></TD><TD><%=strVersion%></TD><TD><%=strModelNumber%></TD><TD><%=strPartNUmber%></TD></TR></TABLE>

<form ID=frmMain action="ProductsSupportedSave.asp" method=post>
	<TABLE bgcolor=cornsilk bordercolor=tan cellpadding=2 cellspacing=0 border =1>
		<%=strHeaderRow%>
	
	<%
	dim strLastProduct
	dim blnProductChecked
	dim strCheckValue

	for i = 0 to ubound(OutputArray)
		OutputArray(i) = "<TD>&nbsp;</TD>"
	next

	strLastProduct = ""
	RowCount = 0
	do while not rs.EOF
		if strLastProduct <> rs("ProductID") then
			if strLastProduct <> "" then
				if Rowcount = 0 then
					Response.Write "<TD>" & Rowcount & "&nbsp;Versions</TD>"
				elseif Rowcount = 1 then
					Response.Write "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & strLastProduct & "&lstRoot=" & strRootID & """>" & Rowcount & "&nbsp;Version</TD>"
				else
					Response.Write "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & strLastProduct& "&lstRoot=" & strRootID & """>" & Rowcount & "&nbsp;Versions</TD>"
				end if
				for i = 0 to ubound(OutputArray)
					Response.write OutputArray(i)
					OutputArray(i) = "<TD>&nbsp;</TD>"
				next
				Response.Write "</tr>"
				RowCount = 0
			end if
			if instr("," & strSelectedProductIDs, "," & trim(rs("ProductID")) & ",") > 0 then
				strCheckValue = " checked "
				strProductsLoaded = strProductsLoaded & "," & rs("ProductID")
			else
				strCheckValue = " "
			end if
			Response.Write "<TR>"
			if blnWorkflowComplete then
				if trim(strCheckValue) = "" then
					Response.Write "<TD nowrap>No</td>"
				else
					Response.Write "<TD nowrap>Yes</td>"
				end if
				Response.Write "<TD style=""Display:none"" nowrap>"
			else
				Response.Write "<TD nowrap>"
			end if
			Response.Write "<INPUT " & strCheckValue & " type=""checkbox"" id=chkSupport name=chkSupport style=""WIDTH: 14px; HEIGHT: 14px"" size=14 value=""" & rs("ProductID") & """></TD><TD nowrap>" & rs("Product") & "</TD>"
			strLastProduct = rs("ProductID")

		end if
		Rowcount = Rowcount + rs("VersionCount")
		for i = 0 to ubound(VendorArray)
			if VendorArray(i) = trim(rs("VendorID")) then
				if rs("VersionCount") = 1 then
					OutputArray(i) = "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & rs("ProductID") & "&lstRoot=" & strRootID & "&lstVendor=" & rs("VendorID") & """>" & rs("VersionCount") & "&nbsp;Version</a></TD>"
				else
					OutputArray(i) = "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & rs("ProductID") & "&lstRoot=" & strRootID & "&lstVendor=" & rs("VendorID") & """>" & rs("VersionCount") & "&nbsp;Versions</TD>"
				end if
				exit for
			end if
		next
		rs.MoveNext
	loop
	rs.Close
	
		if strProductsLoaded <> "" then
			strProductsLoaded = mid(strProductsLoaded,2)
		end if

				if Rowcount = 1 then
					Response.Write "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & strLastProduct & "&lstRoot=" & strRootID & """>" & Rowcount & "&nbsp;Version</TD>"
				else
					Response.Write "<TD><a target=""_blank"" href=""HardwareMatrix.asp?lstProducts=" & strLastProduct & "&lstRoot=" & strRootID & """>" & Rowcount & "&nbsp;Versions</TD>"
				end if
				for i = 0 to ubound(OutputArray)
					Response.write OutputArray(i)
					OutputArray(i) = "<TD>&nbsp;</TD>"
				next
				Response.Write "</tr>"
	
	%>
	</TABLE>
<%
	if len(strVersionsLoaded) > 0 then
		strVersionsLoaded = mid(strVersionsLoaded,3)
	end if

%>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=clng(strProductID)%>">
<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=clng(strRootID)%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=clng(request("VersionID"))%>">
<INPUT type="hidden" id=txtProductsLoaded name=txtProductsLoaded value="<%=strProductsLoaded%>">
</form>


<%
	end if
end if
set rs= nothing
set cn=nothing
%>

</BODY>
</HTML>

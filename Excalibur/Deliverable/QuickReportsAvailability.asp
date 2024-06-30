<%@ Language=VBScript %>

<%
	if request("FileType")= 1  or request("FileType")= 2  then
		Response.ContentType = "application/vnd.ms-excel"
	else
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" language="javascript" src="../_ScriptLibrary/sort.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->

	function ChooseDates(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=10&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
	}


function window_onload() {
	lblLoad.style.display = "none";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	FONT-Family: Verdana;
	FONT-SIZE: xx-small;
	Vertical-Align: top;
}
h3{
	FONT-Family: Verdana;
	FONT-SIZE: medium;
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
<BODY LANGUAGE=javascript onload="return window_onload()">
<b>
<%if request("ProductStatus") = "3" then%>
	<Font size=3 face=verdana>Production Hardware Availability Audit
<%elseif request("ProductStatus") = "4" then%>
	<Font size=3 face=verdana>Post-Production Hardware Availability Audit
<%end if%>
</font>
</b><BR><BR>
<span ID=lblLoad><font size=2 face=verdana>This audit may take a couple minutes to complete.&nbsp;&nbsp;Please wait...</font></span>

<%
response.flush
	dim cn
	dim rs
	dim strLinkParams
	dim strIssue
	dim strProductID
	dim strProductStatusID
	dim strCategoryID
	
	if request("lstProducts") = "" then
		strProductID = "null"
	else
		strProductID = request("lstProducts")
	end if
	if request("ProductStatus") = "" then
		strProductStatusID = "null"
	else
		strProductStatusID = request("ProductStatus")
	end if
	if request("lstCategories") = "" then
		strCategoryID = "null"
	else
		strCategoryID = request("lstCategories")
	end if
	
	strLinkParams = ""

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.CommandTimeout = 240
	cn.IsolationLevel=256
	cn.Open
	
	ColumnCount = 0
	Select Case trim(Request("Report"))
	case "1"
		if trim(strProductStatusID) = "3" then
			rs.Open "spListSubassembliesExpiringProduction " & strProductID & "," & strCategoryID,cn,adOpenStatic
		else
			rs.Open "[spListSubassembliesExpiringPostProduction] " & strProductID & "," & strCategoryID,cn,adOpenStatic
		end if
	case else
		response.write "Not enough information supplied to run report."
	end select

	if trim(Request("Report")) = "1" then
		if rs.EOF and rs.BOF then
			Response.Write "<font size=2 face=verdana><BR>No deliverables found matching your criteria.</font>"
		elseif not (rs.EOF and rs.BOF) then
			Response.Write "<TABLE ID=""tblAudit"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
			Response.Write "<TR bgcolor=beige>"
			Response.write "<TD><b>Matrix</b></TD>"
			Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',1,0,1);"">Product</a></b></TD>"
			Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',2,0,1);"">Business</a></b></TD>"
			Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',3,0,1);"">ODM</a></b></TD>"
			if request("ProductStatus") = "4" then			
				Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',4,2,1);"">EOSL</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',5,0,1);"">Category</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',6,0,1);"">Deliverable</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',7,0,1);"">Subassembly</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',8,0,1);"">Months Remaining</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',9,0,1);"">Issue</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',10,0,1);"">Release&nbsp;Planned</a></b></TD>"
			else
				Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',4,0,1);"">Category</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'tblAudit',5,0,1);"">Deliverable</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',6,0,1);"">Subassembly</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',7,0,1);"">Months Remaining</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',8,0,1);"">Issue</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'tblAudit',9,0,1);"">Release&nbsp;Planned</a></b></TD>"
			end if
			Response.Write "</TR></THEAD>"
		end if
		dim RowCount
		RowCount= 0
		do while not rs.EOF
			RowCount = RowCount + 1
			if rs("WhyIncluded") & "" = "DateMissing or Expiring" then
				if isnull(rs("Months")) then
					strIssue = "Date Missing"
				elseif rs("Months") > 0 then
					strIssue = "Expired"
				else
					strIssue = "Expiring Soon"
				end if
			else
				strIssue = rs("WhyIncluded") & ""
			end if
			if isnull(rs("Months"))then
				if rs("WhyIncluded") & "" <> "DateMissing or Expiring" then			
					strMonths = "Expired"
				else
					strMonths = "Unknown"
				end if
			elseif rs("Months") > 0 then
				strMonths = "Expired"
			else
				strMonths = - rs("Months")
			end if
			strRootSubassembly = rs("Subassembly")
			if instr(strRootSubassembly,"-")>0 then
				strRootSubassembly = left(strRootSubassembly,instr(strRootSubassembly,"-")-1)
			end if
			Response.Write "<TR>"
			if trim(strProductStatusID) = "4" then
				Response.write "<TD><a target=""_blank"" href=""http://16.81.19.70/Deliverable/HardwareMatrix.asp?ReportFormat=5&lstSubassembly=" & strRootSubassembly & "&lstProducts=" & rs("ID") & """>View</a></TD>"
			else
				Response.write "<TD><a target=""_blank"" href=""http://16.81.19.70/Deliverable/HardwareMatrix.asp?ReportFormat=2&lstSubassembly=" & strRootSubassembly & "&lstProducts=" & rs("ID") & """>View</a></TD>"
			end if
			Response.Write "<TD>" & rs("Product") & "</TD>"
			Response.Write "<TD>" & rs("Business") & "</TD>"
			Response.Write "<TD>" & rs("ODM") & "</TD>"
			if request("ProductStatus") = "4" then
			    if isnull(rs("ServiceLifeDate")) then
				    Response.Write "<TD>&nbsp;</TD>"
			    else
				    Response.Write "<TD>" & rs("ServiceLifeDate") & "</TD>"
				end if
			end if
			Response.Write "<TD>" & rs("Category") & "</TD>"
			Response.Write "<TD>" & rs("Deliverable") & "</TD>"
			Response.Write "<TD>" & rs("Subassembly") & "</TD>"
			Response.Write "<TD>" & strMonths & "</TD>"
			Response.Write "<TD>" & strIssue & "</TD>"
			Response.Write "<TD>" & rs("NextTestComplete") & "&nbsp;</TD>"
			Response.Write "</TR>"
			rs.MoveNext
		loop
		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>"
		end if
	
	rs.Close
	end if	
	

	set rs = nothing
	set rs2 = nothing
	cn.Close
	set cn = nothing

	if RowCount > 0 then
		Response.Write "<font size=2 face=verdana><BR><BR>Deliverables Displayed: " & RowCount & "</font>"
	end if
	%>
</BODY>
</HTML>

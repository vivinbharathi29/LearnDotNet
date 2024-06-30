<%@ Language=VBScript %>
<%  Server.ScriptTimeout = 1800 %>
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

	function ChooseDates(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=10&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
	}

	function ChooseDates2(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=11&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
	}

	function ChooseDates3(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=13&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
	}

	function ChooseDates4(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=17&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
	}

	function ChooseDates5(StartDate,EndDate){
	var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate,"","dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.href = "QuickReports.asp?Report=15&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
		}	
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
A:link
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
<BODY>

<%

	dim cn
	dim rs
	dim rs2
	dim strVersion
	dim strLastProduct
	dim strProductList
	dim ColumnCount
	dim DateRangeStart
	dim DateRangeEnd
	dim strDateRange
	dim TempDate
	dim strLinkParams
	dim blnIsAccessoryReport
	dim blnIsPilotReport
	
	strLinkParams = ""
	blnIsAccessoryReport = false
	blnIsPilotReport = false
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	set rs2 = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	ColumnCount = 0
	Select Case trim(Request("Report"))
	case "1"
		Response.Write "<H3>Commodities - Development</H3>"
		rs.Open "spListDeliverablesInWorkflowStateHW 'Development'",cn,adOpenStatic
	case "2"
		Response.Write "<H3>Commodities - Engineering Development</H3>"
		rs.Open "spListDeliverablesInWorkflowStateHW 'Eng. Development'",cn,adOpenStatic
	case "3"
		Response.Write "<H3>Commodities - Core Team</H3>"
		rs.Open "spListDeliverablesInWorkflowStateHW 'Core Team'",cn,adOpenStatic
	case "4"
		Response.Write "<H3>Commodities - Investigating</H3>"
		rs.Open "spListDeliverableHWByStatus 1,1",cn,adOpenStatic
	case "5"
		Response.Write "<H3>Commodities - Failed Qualification</H3>"
		rs.Open "spListDeliverableHWByStatus 10,1",cn,adOpenStatic
	case "6"
		Response.Write "<H3>Commodities - Qualification Hold</H3>"
		rs.Open "spListDeliverableHWByStatus 7,1",cn,adOpenStatic
	case "19"
		Response.Write "<H3>Commodities - Risk Release</H3>"
		rs.Open "spListDeliverableHWByStatus 19,1",cn,adOpenStatic
	case "7"
		blnIsPilotReport = true
		Response.Write "<H3>Commodities - Pilot Failed</H3>"
		rs.Open "spListDeliverableHWByStatus 5,2",cn,adOpenStatic
	case "8"
		blnIsPilotReport = true
		Response.Write "<H3>Commodities - Pilot Hold</H3>"
		rs.Open "spListDeliverableHWByStatus 3,2",cn,adOpenStatic
	case "18"
		blnIsPilotReport = true
		Response.Write "<H3>Commodities - Pilot Complete</H3>"
		rs.Open "spListDeliverableHWByStatus 6,2",cn,adOpenStatic
	case "9"
		blnIsPilotReport = true
		Response.Write "<H3>Commodities - Factory Hold</H3>"
		rs.Open "spListDeliverableHWByStatus 7,2",cn,adOpenStatic
		ColumnCount = 1

	case "16"
		blnIsAccessoryReport = true
		Response.Write "<H3>Accessories - On Hold</H3>"
		rs.Open "spListDeliverableHWByStatus 3,3",cn,adOpenStatic
	case "14"
		blnIsAccessoryReport = true
		Response.Write "<H3>Accessories - Failed</H3>"
		rs.Open "spListDeliverableHWByStatus 5,3",cn,adOpenStatic
	case "10"
		ColumnCount = 1
		Response.Write "<H3>Commodities - Qualification Scheduled</H3>"
		if request("DateStart") <> "" and isdate(request("DateStart")) then
			DateRangeStart = request("DateStart")
		else
			DateRangeStart = ""
		end if

		if request("DateEnd") <> "" and isdate(request("DateEnd")) then
			DateRangeEnd = request("DateEnd")
			if DateRangeStart <> "" then
				if datediff("d",DateRangeStart,DateRangeEnd) < 0 then
					TempDate = DateRangeEnd
					DateRangeEnd = DateRangeStart
					DateRangeStart = TempDate
				end if
			end if
		else
			DateRangeEnd = ""
		end if
		
		strDateRange = "All"

		if DateRangeEnd <> "" and DateRangeStart <> "" then
			strDateRange = DateRangeStart & " - " & DateRangeEnd
		elseif DateRangeEnd <> "" then
			strDateRange = "Before " & DateRangeEnd
		elseif DateRangeStart <> "" then
			strDateRange = "After " & DateRangeStart
		end if

		if Request("FileType") = "" then
			Response.Write "<font size=2 face=verdana>Date Range: " & "<a href=""javascript: ChooseDates('" & DateRangeStart & "','" & DateRangeEnd & "')"">" & strDateRange & "</a><BR></font>"
		else
			Response.Write "<font size=2 face=verdana>Date Range: " & strDateRange & "<BR></font>"
		end if

		strLinkParams = ""
		if request("DateStart") <> "" then
			strLinkParams = strLinkParams & "&DateStart=" & request("DateStart")
		end if
		if request("DateEnd") <> "" then
			strLinkParams = strLinkParams & "&DateEnd=" & request("DateEnd")
		end if
		
		
		rs.Open "spListDeliverableHWByStatus 3,1,'" & DateRangeStart & "','" & DateRangeEnd & "'",cn,adOpenStatic
	case "11"
		blnIsPilotReport = true
		ColumnCount = 1
		Response.Write "<H3>Commodities - Pilot Scheduled</H3>"
		if request("DateStart") <> "" and isdate(request("DateStart")) then
			DateRangeStart = request("DateStart")
		else
			DateRangeStart = ""
		end if

		if request("DateEnd") <> "" and isdate(request("DateEnd")) then
			DateRangeEnd = request("DateEnd")
			if DateRangeStart <> "" then
				if datediff("d",DateRangeStart,DateRangeEnd) < 0 then
					TempDate = DateRangeEnd
					DateRangeEnd = DateRangeStart
					DateRangeStart = TempDate
				end if
			end if
		else
			DateRangeEnd = ""
		end if
		
		strDateRange = "All"

		if DateRangeEnd <> "" and DateRangeStart <> "" then
			strDateRange = DateRangeStart & " - " & DateRangeEnd
		elseif DateRangeEnd <> "" then
			strDateRange = "Before " & DateRangeEnd
		elseif DateRangeStart <> "" then
			strDateRange = "After " & DateRangeStart
		end if

		if Request("FileType") = "" then
			Response.Write "<font size=2 face=verdana>Date Range: " & "<a href=""javascript: ChooseDates2('" & DateRangeStart & "','" & DateRangeEnd & "')"">" & strDateRange & "</a><BR></font>"
		else
			Response.Write "<font size=2 face=verdana>Date Range: " & strDateRange & "<BR></font>"
		end if

		strLinkParams = ""
		if request("DateStart") <> "" then
			strLinkParams = strLinkParams & "&DateStart=" & request("DateStart")
		end if
		if request("DateEnd") <> "" then
			strLinkParams = strLinkParams & "&DateEnd=" & request("DateEnd")
		end if
				
		rs.Open "spListDeliverableHWByStatus 2,2,'" & DateRangeStart & "','" & DateRangeEnd & "'",cn,adOpenStatic
	case "15"
		blnIsAccessoryReport = true
		ColumnCount = 1
		Response.Write "<H3>Accessories - Scheduled</H3>"
		if request("DateStart") <> "" and isdate(request("DateStart")) then
			DateRangeStart = request("DateStart")
		else
			DateRangeStart = ""
		end if

		if request("DateEnd") <> "" and isdate(request("DateEnd")) then
			DateRangeEnd = request("DateEnd")
			if DateRangeStart <> "" then
				if datediff("d",DateRangeStart,DateRangeEnd) < 0 then
					TempDate = DateRangeEnd
					DateRangeEnd = DateRangeStart
					DateRangeStart = TempDate
				end if
			end if
		else
			DateRangeEnd = ""
		end if
		
		strDateRange = "All"

		if DateRangeEnd <> "" and DateRangeStart <> "" then
			strDateRange = DateRangeStart & " - " & DateRangeEnd
		elseif DateRangeEnd <> "" then
			strDateRange = "Before " & DateRangeEnd
		elseif DateRangeStart <> "" then
			strDateRange = "After " & DateRangeStart
		end if

		if Request("FileType") = "" then
			Response.Write "<font size=2 face=verdana>Date Range: " & "<a href=""javascript: ChooseDates5('" & DateRangeStart & "','" & DateRangeEnd & "')"">" & strDateRange & "</a><BR></font>"
		else
			Response.Write "<font size=2 face=verdana>Date Range: " & strDateRange & "<BR></font>"
		end if

		strLinkParams = ""
		if request("DateStart") <> "" then
			strLinkParams = strLinkParams & "&DateStart=" & request("DateStart")
		end if
		if request("DateEnd") <> "" then
			strLinkParams = strLinkParams & "&DateEnd=" & request("DateEnd")
		end if
				
		rs.Open "spListDeliverableHWByStatus 2,3,'" & DateRangeStart & "','" & DateRangeEnd & "'",cn,adOpenStatic
	case "12"
		Response.Write "<H3>Commodities - Supply Chain Restriction</H3>"
		rs.Open "spListCommodityRestrictions",cn,adOpenStatic
	case "13"
		blnIsPilotReport = true
		ColumnCount = 1
		Response.Write "<H3>Commodities - Pilot Complete</H3>"
		if request("DateStart") <> "" and isdate(request("DateStart")) then
			DateRangeStart = request("DateStart")
		else
			DateRangeStart = ""
		end if

		if request("DateEnd") <> "" and isdate(request("DateEnd")) then
			DateRangeEnd = request("DateEnd")
			if DateRangeStart <> "" then
				if datediff("d",DateRangeStart,DateRangeEnd) < 0 then
					TempDate = DateRangeEnd
					DateRangeEnd = DateRangeStart
					DateRangeStart = TempDate
				end if
			end if
		else
			DateRangeEnd = ""
		end if
		
		strDateRange = "All"

		if DateRangeEnd <> "" and DateRangeStart <> "" then
			strDateRange = DateRangeStart & " - " & DateRangeEnd
		elseif DateRangeEnd <> "" then
			strDateRange = "Before " & DateRangeEnd
		elseif DateRangeStart <> "" then
			strDateRange = "After " & DateRangeStart
		end if

		if Request("FileType") = "" then
			Response.Write "<font size=2 face=verdana>Date Range: " & "<a href=""javascript: ChooseDates3('" & DateRangeStart & "','" & DateRangeEnd & "')"">" & strDateRange & "</a><BR></font>"
		else
			Response.Write "<font size=2 face=verdana>Date Range: " & strDateRange & "<BR></font>"
		end if

		strLinkParams = ""
		if request("DateStart") <> "" then
			strLinkParams = strLinkParams & "&DateStart=" & request("DateStart")
		end if
		if request("DateEnd") <> "" then
			strLinkParams = strLinkParams & "&DateEnd=" & request("DateEnd")
		end if
				
		rs.Open "spListDeliverableHWByStatus 6,2,'" & DateRangeStart & "','" & DateRangeEnd & "'",cn,adOpenStatic
	
	case "17"
		blnIsAccessoryReport = true
		ColumnCount = 1
		Response.Write "<H3>Accessories - Complete</H3>"
		if request("DateStart") <> "" and isdate(request("DateStart")) then
			DateRangeStart = request("DateStart")
		else
			DateRangeStart = ""
		end if

		if request("DateEnd") <> "" and isdate(request("DateEnd")) then
			DateRangeEnd = request("DateEnd")
			if DateRangeStart <> "" then
				if datediff("d",DateRangeStart,DateRangeEnd) < 0 then
					TempDate = DateRangeEnd
					DateRangeEnd = DateRangeStart
					DateRangeStart = TempDate
				end if
			end if
		else
			DateRangeEnd = ""
		end if
		
		strDateRange = "All"

		if DateRangeEnd <> "" and DateRangeStart <> "" then
			strDateRange = DateRangeStart & " - " & DateRangeEnd
		elseif DateRangeEnd <> "" then
			strDateRange = "Before " & DateRangeEnd
		elseif DateRangeStart <> "" then
			strDateRange = "After " & DateRangeStart
		end if

		if Request("FileType") = "" then
			Response.Write "<font size=2 face=verdana>Date Range: " & "<a href=""javascript: ChooseDates4('" & DateRangeStart & "','" & DateRangeEnd & "')"">" & strDateRange & "</a><BR></font>"
		else
			Response.Write "<font size=2 face=verdana>Date Range: " & strDateRange & "<BR></font>"
		end if

		strLinkParams = ""
		if request("DateStart") <> "" then
			strLinkParams = strLinkParams & "&DateStart=" & request("DateStart")
		end if
		if request("DateEnd") <> "" then
			strLinkParams = strLinkParams & "&DateEnd=" & request("DateEnd")
		end if
				
		rs.Open "spListDeliverableHWByStatus 6,3,'" & DateRangeStart & "','" & DateRangeEnd & "'",cn,adOpenStatic
	
	case else
		response.write "Not enough information supplied to run report."
	end select

	if trim(Request("Report")) = "1" or trim(Request("Report")) = "2" or trim(Request("Report")) = "3" then

		if rs.EOF and rs.BOF then
			Response.Write "<font size=2 face=verdana><BR>No deliverables found matching your criteria.</font>"
		elseif not (rs.EOF and rs.BOF) then
			if Request("FileType") = "" then
				Response.Write "<TABLE BORDER=0 width=""100%""><TR><TD align=right><A target=""_blank"" href=""Quickreports.asp?FileType=1&Report=" & request("Report") & strLinkParams & """>Export</a></TD></TR></TABLE>"
				Response.Write "<TABLE ID=""ResultTable"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b><a href=""javascript: SortTable( 'ResultTable', 0,1,2);"">ID</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 1,0,1);"">Category</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 2,0,1);"">Vendor</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 3,0,1);"">Deliverable</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 4,0,1);"">HW/FW/Rev</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 5,0,1);"">Model</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 6,0,1);"">Part</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 7,0,1);"">Developer</a></b></TD>"
				Response.Write "<TD><b>Products</b></TD></TR></THEAD>"
			else
				Response.Write "<TABLE ID=""ResultTable"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b>ID</b></TD>"
				Response.write "<TD><b>Category</b></TD>"
				Response.Write "<TD><b>Vendor</b></TD>"
				Response.Write "<TD><b>Deliverable</b></TD>"
				Response.Write "<TD><b>HW/FW/Rev</b></TD>"
				Response.Write "<TD><b>Model</b></TD>"
				Response.Write "<TD><b>Part</b></TD>"
				Response.Write "<TD><b>Developer</b></TD>"
				Response.Write "<TD><b>Products</b></TD></TR></THEAD>"
			end if
		end if
		strLastVersion = ""
		strProductList = ""
		do while not rs.EOF
		
			if strLastVersion <> rs("ID") and strLastVersion <> "" then
				if strProductList = "" then
					strProductList = "&nbsp;"
				else
					strProductList = mid(strProductList,3)
				end if
				Response.Write "<TD>" & strProductList & "&nbsp;</td>"
				Response.Write "</TR>"
				strProductList = ""
			end if
			if strLastVersion <> rs("ID") then
				strVersion = rs("version") & ""
				if rs("Revision") & "" <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if rs("Pass") & "" <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				Response.Write "<TR><TD><a target=_blank href=""../Query/DeliverableVersionDetails.asp?ID=" & rs("ID") & """>" & rs("ID") & "</a></TD>"
				Response.Write "<TD>" & rs("Category") & "</TD>"
				Response.Write "<TD>" & rs("Vendor") & "</TD>"
				Response.Write "<TD>" & rs("Name") & "</TD>"
				Response.Write "<TD>" & strVersion & "</TD>"
				Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
				Response.Write "<TD nowrap>" & rs("PartNumber") & "&nbsp;</TD>"
				Response.Write "<TD>" & rs("Developer") & "</TD>"
			end if		
			strProductList = strProductList & ", " & rs("Product") 
			strLastVersion = rs("ID")
            Response.Flush
			rs.MoveNext
		loop

		if strLastVersion <> "" then
			if strProductList = "" then
				strProductList = "&nbsp;"
			else
				strProductList = mid(strProductList,3)
			end if
			Response.Write "<TD>" & strProductList & "&nbsp;</td>"
			Response.Write "</TR>"
		end if



		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>"
		end if

    elseif trim(Request("Report")) = "6" or trim(Request("Report")) = "19"  then
    
		if rs.EOF and rs.BOF then
			Response.Write "<font size=2 face=verdana><BR>No deliverables found matching your criteria.</font>"
		elseif not (rs.EOF and rs.BOF) then
			if Request("FileType") = "" then
				Response.Write "<TABLE BORDER=0 width=""100%""><TR><TD align=right><A target=""_blank"" href=""Quickreports.asp?FileType=1&Report=" & request("Report") & strLinkParams & """>Export</a></TD></TR></TABLE>"
				Response.Write "<TABLE ID=""ResultTable"" width=150% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b><a href=""javascript: SortTable( 'ResultTable', 0,1,2);"">ID</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 1,0,1);"">Product</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 2,0,1);"">ODM</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 3,0,1);"">PM</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 4,0,1);"">Part</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 5,0,1);"">Vendor</a></b></TD>"
				Response.Write "<TD style=""width:300px""><b><a href=""javascript: SortTable( 'ResultTable', 6,0,1);"">Deliverable</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 7,0,1);"">HW/FW/Rev</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 8,0,1);"">Model</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 9,2,1);"">Status&nbsp;Date</a></b></TD>"
				Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 10,2,1);"">Target&nbsp;Date</a></b></TD>"
				Response.Write "<TD><b>Comments</b></TD></TR></THEAD>"
			else
				Response.Write "<TABLE ID=""ResultTable"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b>ID</b></TD>"
				Response.write "<TD><b>Product</b></TD>"
				Response.Write "<TD><b>ODM</b></TD>"
				Response.Write "<TD><b>PM</b></TD>"
				Response.Write "<TD><b>Part</b></TD>"
				Response.Write "<TD><b>Vendor</b></TD>"
				Response.Write "<TD style=""width:300px""><b>Deliverable</b></TD>"
				Response.Write "<TD><b>HW/FW/Rev</b></TD>"
				Response.Write "<TD><b>Model</b></TD>"
				Response.write "<TD><b>Status&nbsp;Date</b></TD>"
				Response.write "<TD><b>Target&nbsp;Date</b></TD>"
				Response.Write "<TD><b>Comments</b></TD></TR></THEAD>"
			end if


		end if
		RowCount= 0
		do while not rs.EOF
			RowCount = RowCount + 1
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			strComments = ""
			if (rs("DevComments") & "") <> "" then
				strComments = "<b>Developer:</b>" & rs("DevComments")
			end if
			if trim(rs("QualComments") & "") <> "" then
				if strComments <> "" then
					strComments = strComments & "<BR>"
				end if
				strComments = strComments & "<b>Qual:</b> " & rs("QualComments")
			end if
			if 	blnIsPilotReport then
				if trim(rs("PilotComments") & "") <> "" then
					if strComments <> "" then
						strComments = strComments & "<BR>"
					end if
					strComments = strComments & "<b>Pilot:</b> " & rs("PilotComments")
				end if
			end if
			if 	blnIsAccessoryReport then
				if trim(rs("AccessoryComments") & "") <> "" then
					if strComments <> "" then
						strComments = strComments & "<BR>"
					end if
					strComments = strComments & "<b>Accessory:</b> " & rs("AccessoryComments")
				end if
			end if
			
			
			Response.Write "<TR><TD><a target=_blank href=""../Query/DeliverableVersionDetails.asp?ID=" & rs("DelID") & """>" & rs("DelID") & "</a></TD>"
			Response.Write "<TD>" & rs("Product") & "</TD>"
			Response.Write "<TD>" & rs("Partner") & "&nbsp;</TD>"
			Response.Write "<TD nowrap>" & rs("PM") & "&nbsp;</TD>"
			Response.Write "<TD nowrap>" & rs("PartNumber") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Vendor") & "</TD>"
			Response.Write "<TD>" & rs("Deliverable") & "</TD>"
			Response.Write "<TD>" & strVersion & "</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
			if isnull(rs("StatusUpdatedDate")) then
			    Response.Write "<TD>&nbsp;</TD>"
			else
			    Response.Write "<TD>" & formatdatetime(rs("StatusUpdatedDate"),vbshortdate) & "</TD>"
			end if
			if isnull(rs("TestDate")) then
			    Response.Write "<TD>&nbsp;</TD>"
		    else
			    Response.Write "<TD>" & formatdatetime(rs("TestDate"),vbshortdate) & "</TD>"
		    end if
			Response.Write "<TD>" & strComments & "&nbsp;</TD>"
			Response.Write "</TR>"
            Response.Flush
			rs.MoveNext
		loop
		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>"
		end if
    
    
	else
	
		if rs.EOF and rs.BOF then
			Response.Write "<font size=2 face=verdana><BR>No deliverables found matching your criteria.</font>"
		elseif not (rs.EOF and rs.BOF) then
			if Request("FileType") = "" then
				Response.Write "<TABLE BORDER=0 width=""100%""><TR><TD align=right><A target=""_blank"" href=""Quickreports.asp?FileType=1&Report=" & request("Report") & strLinkParams & """>Export</a></TD></TR></TABLE>"
				Response.Write "<TABLE ID=""ResultTable"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b><a href=""javascript: SortTable( 'ResultTable', 0,1,2);"">ID</a></b></TD>"
				if trim(Request("Report")) = "10" or trim(Request("Report")) = "11" or trim(Request("Report")) = "9" then
					Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', 1,2,1);"">Date</a></b></TD>"
				end if
				Response.write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 1 & ",0,1);"">Product</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 2 & ",0,1);"">Vendor</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 3 & ",0,1);"">Deliverable</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 4 & ",0,1);"">HW/FW/Rev</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 5 & ",0,1);"">Model</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 6 & ",0,1);"">Part</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 7 & ",0,1);"">ODM</a></b></TD>"
				Response.Write "<TD><b><a href=""javascript: SortTable( 'ResultTable', " & ColumnCount + 8 & ",0,1);"">ODM&nbsp;PM</a></b></TD>"
				Response.Write "<TD><b>Comments</b></TD></TR></THEAD>"
			else
				Response.Write "<TABLE ID=""ResultTable"" width=100% bgcolor=ivory border=1 cellspacing=0 cellpadding=2><THEAD>"
				Response.Write "<TR bgcolor=beige><TD><b>ID</b></TD>"
				if trim(Request("Report")) = "10" or trim(Request("Report")) = "11" or trim(Request("Report")) = "15" or trim(Request("Report")) = "9" then
					Response.write "<TD><b>Date</b></TD>"
				end if
				Response.write "<TD><b>Product</b></TD>"
				Response.Write "<TD><b>Vendor</b></TD>"
				Response.Write "<TD><b>Deliverable</b></TD>"
				Response.Write "<TD><b>HW/FW/Rev</b></TD>"
				Response.Write "<TD><b>Model</b></TD>"
				Response.Write "<TD><b>Part</b></TD>"
				Response.Write "<TD><b>ODM</b></TD>"
				Response.Write "<TD><b>ODM&nbsp;PM</b></TD>"
				Response.Write "<TD><b>Comments</b></TD></TR></THEAD>"
			end if


			'Response.Write "<TABLE ID=""ResultTable"" bgcolor=ivory border=1 cellspacing=0 cellpadding=2>"
			'Response.Write "<TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW/FW/Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Developer</b></TD><TD><b>Comments</b></TD></TR>"
		end if
		dim RowCount
		RowCount= 0
		do while not rs.EOF
			RowCount = RowCount + 1
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			strComments = ""
			if (rs("DevComments") & "") <> "" then
				strComments = "<b>Developer:</b>" & rs("DevComments")
			end if
			if trim(rs("QualComments") & "") <> "" then
				if strComments <> "" then
					strComments = strComments & "<BR>"
				end if
				strComments = strComments & "<b>Qual:</b> " & rs("QualComments")
			end if
			if 	blnIsPilotReport then
				if trim(rs("PilotComments") & "") <> "" then
					if strComments <> "" then
						strComments = strComments & "<BR>"
					end if
					strComments = strComments & "<b>Pilot:</b> " & rs("PilotComments")
				end if
			end if
			if 	blnIsAccessoryReport then
				if trim(rs("AccessoryComments") & "") <> "" then
					if strComments <> "" then
						strComments = strComments & "<BR>"
					end if
					strComments = strComments & "<b>Accessory:</b> " & rs("AccessoryComments")
				end if
			end if
			
			
			Response.Write "<TR><TD><a target=_blank href=""../Query/DeliverableVersionDetails.asp?ID=" & rs("DelID") & """>" & rs("DelID") & "</a></TD>"
			if trim(Request("Report")) = "10"  then
				if isdate(rs("TestDate")) then
					Response.Write "<TD>" & rs("TestDate") & "</TD>"
				else
					Response.Write "<TD>&nbsp;</TD>"
				end if
			elseif trim(Request("Report")) = "11"  or trim(Request("Report")) = "9" then
				if isdate(rs("PilotDate")) then
					Response.Write "<TD>" & formatdatetime(rs("PilotDate"),vbshortdate) & "</TD>"
				else
					Response.Write "<TD>&nbsp;</TD>"
				end if
			elseif trim(Request("Report")) = "15" then
				if isdate(rs("AccessoryDate")) then
					Response.Write "<TD>" & rs("AccessoryDate") & "</TD>"
				else
					Response.Write "<TD>&nbsp;</TD>"
				end if
			end if
			Response.Write "<TD>" & rs("Product") & "</TD>"
			Response.Write "<TD>" & rs("Vendor") & "</TD>"
			Response.Write "<TD>" & rs("Deliverable") & "</TD>"
			Response.Write "<TD>" & strVersion & "</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("PartNumber") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Partner") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("PM") & "&nbsp;</TD>"
			Response.Write "<TD>" & strComments & "&nbsp;</TD>"
			Response.Write "</TR>"
            Response.Flush	
			rs.MoveNext
		loop
		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>"
		end if
	
	end if	
	rs.Close


	set rs = nothing
	set rs2 = nothing
	cn.Close
	set cn = nothing

%>

<%
'fp = "c:\pagecounter.txt"'Server.MapPath("")
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set a = fs.OpenTextFile(fp)
'ct = Clng(a.ReadLine)
''if Session("ct") = "" then
''Session("ct") = ct
'ct = ct + 1
'a.close
'Set a = fs.CreateTextFile(fp, True)
'a.WriteLine(ct)
''end if 
'a.Close
''Response.Write ct
%>

<%
	if RowCount > 0 then
		Response.Write "<font size=2 face=verdana><BR><BR>Deliverables Displayed: " & RowCount & "</font>"
	end if
	%>
</BODY>
</HTML>

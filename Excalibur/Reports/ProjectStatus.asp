<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
	h3
	{
	FONT-SIZE: small;
	FONT-FAMILY: Verdana;
	}
	
	LI
	{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	}
	BODY
	{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	}
</STYLE>
<BODY>
<%

	dim strItems
	dim StartDate
	dim EndDate
	dim RoadmapCount
	dim cn
	dim rs
	dim strProductName
	dim strProductID	
	dim strResolution
	dim sreDescription

  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
  	set rs = server.CreateObject("ADODB.Recordset")

	StartDate = FormatDateTime(Now()-8,vbshortdate)
	EndDate = Now()

	rs.Open "spGetProductVersionName " & clng(Request("ID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "Product Not Found"
		strProductID = 0
	else
		strProductID =  clng(Request("ID"))
		strProductName = rs("Name")
	end if
	rs.Close

	if strProductID <> 0 then
		'Header
		Response.Write "<Center><font size=3 face=verdana><b>" & strProductName & " Status Report</b></font><BR><BR>"
		Response.Write "<font size=1 face=verdana>" & StartDate & " - " & formatDateTime(EndDate,vbshortdate) & "</font>"
		Response.write "</Center>"
	
	
		'Accomplishments
		rs.Open "spListActionsClosedSummary " & strProductID & ",2,'" & StartDate & "','" & EndDate & "'" ,cn,adOpenForwardOnly
		strItems = ""
		do while not rs.EOF
			strItems = strItems & "<LI>" & rs("Summary") & "</LI>"
			if trim(rs("Resolution") & "") <> "" then
				strItems = strItems & "<UL>"
				for each strResolution in split(rs("Resolution") & "",vbcrlf)
					if strResolution <> "" then
						strItems = strItems & "<LI>" & strResolution & "</LI>"
					end if
				next
				strItems = strItems & "</UL>"
			end if
			rs.MoveNext
		loop
		rs.Close
	
		Response.Write "<h3>Accomplishments</h3>"
		if strItems = "" then
			Response.Write "<Blockquote>No tasks were closed between " & StartDate & " and " & formatDateTime(EndDate,vbshortdate) & "</Blockquote>"
		else
			Response.Write "<UL>" & strItems & "</UL>"
		end if
		
		
		'Priorities
		rs.Open "spListActionRoadmap " & strProductID ,cn,adOpenForwardOnly
		strItems = ""
		RoadmapCount = 0
		do while not rs.EOF
			if rs("StatusReport") then
				strItems = strItems & "<LI>" & rs("Summary") & "</LI>"
				RoadmapCount = RoadmapCount + 1
				if RoadmapCount > 9 then
					Exit do
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
	
		Response.Write "<h3>Priorities</h3>"
		if strItems = "" then
			Response.Write "<Blockquote>No open priorities found.</blockquote>"
		else
			Response.Write "<UL>" & strItems & "</UL>"
		end if	
		
	
	
		'Open Issues
		rs.Open "spListActionItems " & strProductID & ",1,1",cn,adOpenForwardOnly
		strItems = ""
		do while not rs.EOF
			strItems = strItems & "<LI>" & rs("Summary") & "</LI>"
			if trim(rs("Description") & "") <> "" then
				strItems = strItems & "<UL>"
				for each strDescription in split(rs("Description") & "",vbcrlf)
					if strDescription <> "" then
						strItems = strItems & "<LI>" & strDescription & "</LI>"
					end if
				next
				strItems = strItems & "</UL>"
			end if
			
			rs.MoveNext
		loop
		rs.Close
	
		Response.Write "<h3>Issues</h3>"
		if strItems = "" then
			Response.Write "<Blockquote>No open issues found.</blockquote>"
		else
			Response.Write "<UL>" & strItems & "</UL>"
		end if	
	
	end if

	
	'Cleanup
	set rs = nothing
	cn.Close
	set cn=nothing
%>
</BODY>
</HTML>

<%@ Language=VBScript %>
<%
	if request("Type") = "Excel" then
		Response.ContentType = "application/vnd.ms-excel"
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%
	if trim(request("ID")) = "" then
		Response.write "No product specified."
	else
		cnString =Session("PDPIMS_ConnectionString")
	
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = cnString
		cn.Open
	
		set rs = server.CreateObject("ADODB.recordset")
		rs.ActiveConnection = cn
		
		rs.Open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			Response.Write rs("Name") & " Deliverables" 
		end if
		rs.Close
	'	strSQL = "Select v.id, v.deliverablename, v.version, v.revision, v.pass, pd.targetnotes, vd.name as vendor, v.vendorversion, v.filename, v.imagepath as location, c.name as category " & _
		strSQL = "Select v.id, v.deliverablename, v.version, v.revision, v.pass, pd.preinstall, pd.preload, pd.arcd, pd.web, pd.dropinbox, pd.selectiveRestore, pd.drdvd, pd.patch, pd.RACD_EMEA, pd.RACD_Americas,pd.RACD_APD, pd.DOCCD, pd.OSCD, pd.targetnotes, vd.name as vendor, v.vendorversion, v.filename, v.imagepath as location, c.name as category " & _
			  "from deliverableroot r with (NOLOCK), deliverableversion v with (NOLOCK), product_deliverable pd with (NOLOCK), vendor vd with (NOLOCK), deliverablecategory c with (NOLOCK) " & _
			  "where r.id = v.deliverablerootid " & _
			  "and v.id = pd.deliverableversionid " & _
			  "and pd.productversionid = " & clng(request("ID")) & " " & _
			  "and c.id = r.categoryid " & _
			  "and r.typeid = 2 " & _
			  "and vd.id = v.vendorid " & _
			  "and v.languages like '%US%' " & _
			  "and pd.targeted=1 " & _
			  "and c.id not in (170,171)"	
		rs.open strSQL,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.write "No deliverables found for this product."
		else
			Response.Write "<TABLE border=1>"
			Response.Write "<TR><TD><b>ID</b></TD><TD><b>Name</b></TD><TD><b>Version</b></TD><TD><b>TargetNotes</b></TD><TD><b>Vendor</b></TD><TD><b>VendorVersion</b></TD><TD><b>Path</b></TD><TD><b>Category</b></TD><TD><b>OS</b></TD><TD><b>Distribution</b></TD></TR>"
			do while not rs.EOF
				strVersion = rs("Version") & ""
				if trim(rs("Revision") & "") <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if trim(rs("Pass") & "") <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
			
				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spGetSelectedOS " & rs("ID"),cn,adOpenForwardOnly
				strOS = ""
				do while not rs2.EOF
					strOS = strOS & "," & rs2("Name")
					rs2.MoveNext
				loop
				rs2.Close
				set rs2 = nothing
				if len(strOS) > 0 then
					strOS = server.HTMLEncode(mid(strOS,2))
				end if
				
				strDistribution = ""
				if rs("Preinstall") then
					strDistribution = ",Preinstall"
				end if
				if rs("Preload") then
					strDistribution = strDistribution & ",Preload"
				end if
				if rs("DropInBox") then
					strDistribution = strDistribution & ",DIB"
				end if
				if rs("Web") then
					strDistribution = strDistribution & ",Web"
				end if
				if rs("SelectiveRestore") then
					strDistribution = strDistribution & ",SelectiveRestore"
				end if
				if rs("ARCD") then
					strDistribution = strDistribution & ",DRCD"
				end if
				if rs("DRDVD") then
					strDistribution = strDistribution & ",DRDVD"
				end if
				if rs("OSCD") then
					strDistribution = strDistribution & ",OSCD"
				end if
				if rs("DocCD") then
					strDistribution = strDistribution & ",DocCD"
				end if
				if rs("RACD_EMEA") then
					strDistribution = strDistribution & ",RACD_EMEA"
				end if
				if rs("RACD_AMERICAS") then
					strDistribution = strDistribution & ",RACD_Americas"
				end if
				if rs("RACD_APD") then
					strDistribution = strDistribution & ",RACD_APD"
				end if
				if trim(rs("Patch")&"") <> "0" then
					strDistribution = strDistribution & ",Patch"
				end if
				if len(strDistribution) > 0 then
					strDistribution = mid(strDistribution,2)
				end if
				
				Response.Write "<TR><TD>" & rs("ID") & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("DeliverableName") & "") & "</TD>"
				Response.Write "<TD>" & server.htmlencode(strVersion) & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("TargetNotes") & "") & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("Vendor") & "") & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("VendorVersion" & "")) & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("Location") & "") & "</TD>"
				Response.Write "<TD>" & server.htmlencode(rs("Category") & "") & "</TD>"
				Response.Write "<TD>" & strOS & "</TD>"
				Response.Write "<TD>" & strDistribution & "</TD>"
				Response.Write "</TR>"
				rs.movenext	
			loop
			Response.Write "</TABLE>"
		end if
		
		rs.Close
		cn.Close
		set rs = nothing
		set cn = nothing
	end if  
	 %>

</BODY>
</HTML>

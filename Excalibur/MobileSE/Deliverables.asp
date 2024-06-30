<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD
{
	FONT-SIZE: xx-small;
    COLOR: black;
    FONT-FAMILY: Verdana
}
</STYLE>
<BODY>

<font size=3 face=verdana><b>Product Deliverable Comparison</b></font><BR><BR>
<%
	if request("Report")="2" then
		Response.Write "<font face=verdana size=2>Display: Differences Only</font><BR><BR>"
	else
		Response.Write "<font face=verdana size=2>Display: All Deliverables</font><BR><BR>"
	end if 
%>
<BR>

<%
	dim rs	
	dim cn
	dim strSQL
	dim i
	dim strLastType
	dim strLastRoot
	dim ProductArray
	dim ProductIDArray
	dim ProductCount
	dim blnAllDups

	ProductArray = split(request("ID"),"_")
	ProductIDArray = split(request("ID"),"_")

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
			
	set rs = server.CreateObject("ADODB.recordset")

	rs.Open "Select v.ID, f.name, v.version from productfamily f with (NOLOCK), productversion v with (NOLOCK) where v.productfamilyid = f.id and v.id in (" & replace(request("ID"),"_",",") & ") order by f.name, v.version" ,cn,adOpenForwardOnly

	Response.Write "<TABLE bordercolor=tan border=1 bgcolor=Ivory width=""100%""><tr><td>&nbsp;</td>"
	i=0
	do while not rs.EOF
		Response.Write "<td>" &  rs("Name") & " " & rs("Version") & "</td>"
		'Response.Write rs("ID") & ":" &  & "<BR>"
		ProductArray(i) = rs("ID")
		ProductIDArray(i) = rs("ID")
		i=i+1
		rs.MoveNext
	loop
	ProductCount = i
	rs.Close
	Response.Write "</TR>"
	
	strSQL = "SELECT pd.Preinstall, pd.inimage, pd.targetnotes, pd.selectiverestore, pd.RACD_Americas, pd.RACD_EMEA, RACD_APD ,pd.doccd, pd.patch, pd.oscd, pd.preload, pd.dropinbox, pd.web, pd.arcd, pd.drdvd, pd.imagesummary, v.id, f.name + '" & " " & "' + v.version as Product, dv.version as DelVersion, dv.revision as DelRevision, dv.pass as DelPass, dr.name as Root, t.id as TypeID, t.name as Type  " & _
	         "FROM productfamily f with (NOLOCK), productversion v with (NOLOCK), product_deliverable pd with (NOLOCK), deliverableversion dv with (NOLOCK), deliverableroot dr with (NOLOCK), deliverabletype t with (NOLOCK) " & _
	         "WHERE f.ID = v.productfamilyID " & _
	         "AND v.ID = pd.productversionid " & _
	         "AND dv.id = pd.deliverableversionid " & _
	         "AND t.id = dr.typeid " & _
	         "AND dr.id = dv.DeliverableRootid " & _
	         "AND pd.Targeted=1 " & _
	         "AND v.id in (" & replace(request("ID"),"_",",") & ") " & _
			 "Order By t.name desc, dr.name, dv.version, dv.revision, dv.pass, v.id, f.name , v.version"



	rs.Open strSQl,cn,adOpenForwardOnly
	strLastType = ""
	strLastRoot = ""
	for i = 0 to ubound(ProductArray)
		ProductArray(i) = ""
	next
	do while not rs.EOF
		if strlastRoot <> rs("Root") and strLastRoot <> "" then
			blnAllDups = true
			for i = 0 to ubound(ProductArray) -1
				if replace(replace(ucase(trim(ProductArray(i)))," ",""),",","") <>  replace(replace(ucase(trim(ProductArray(i + 1 )))," ",""),",","") then
					blnAllDups = false
					exit for
				end if 
			next

			if (request("Report")= "2" and not blnAllDups) or request("Report")<> "2" then
				Response.Write "<TR><TD>" & strLastRoot & "</td>"
				for i = 0 to ubound(ProductArray)
					response.write "<TD valign=top>" & ProductArray(i) & "&nbsp;</TD>"
					ProductArray(i) = ""
				next
				response.write "</TR>"
			else
				for i = 0 to ubound(ProductArray)
					ProductArray(i) = ""
				next
			end if
		end if
		
		for i = 0 to ubound(ProductArray)
			if trim(ProductIDArray(i)) = trim(rs("ID")) then
				strVersion = rs("DelVersion")
				if trim(rs("DelRevision") & "") <> "" then
					strVersion = strVersion  & "," & rs("DelRevision")
				end if
				if trim(rs("DelPass") & "") <> "" then
					strVersion = strVersion  & "," & rs("DelPass")
				end if
				
				dim strDistribution
				strDistribution = ""

				if rs("Preinstall") then
					strDistribution = ", Preinstall"
				end if
				if rs("Preload") then
					strDistribution = strDistribution & ", Preload"
				end if	
				
				if rs("DropInBox") then
					strDistribution = strDistribution & ", DIB"
				end if
				if rs("Web") then
					strDistribution = strDistribution & ", Web"
				end if	
				if rs("DRDVD") then
					strDistribution = strDistribution & ", DRDVD"
				end if

				if rs("ARCD") then
					strDistribution = strDistribution & ", DRCD"
				end if

				if rs("RACD_Americas") then
					strDistribution = strDistribution & ", RACD-Americas"
				end if

				if rs("RACD_EMEA") then
					strDistribution = strDistribution & ", RACD-EMEA"
				end if

				if rs("RACD_APD") then
					strDistribution = strDistribution & ", RACD-APD"
				end if

				if rs("OSCD") then
					strDistribution = strDistribution & ", OSCD"
				end if
					
				if rs("DocCD") then
					strDistribution = strDistribution & ", Doc CD"
				end if

				if trim(rs("Patch")&"") <> "0" then
					strDistribution = strDistribution & ", Patch"
				end if	

				if rs("SelectiveRestore") then
					strDistribution = strDistribution & ", SelectiveRestore"
				end if

				if strDistribution <> "" then
					strDistribution = mid(strDistribution,3)
				end if
				if trim(request("Images")) = "1" then
					strImageSummary = rs("ImageSummary")
					if ucase(strImageSummary) = "ALL" or ucase(strImageSummary) = "ALL.." or ucase(strImageSummary) = "ALL." or ucase(strImageSummary) = "ALL .." or ucase(strImageSummary) = "ALL ." then
						strImageSummary  = "All"
					end if
					strImageSummary = "<BR><b>Images: </b>" & strImageSummary
				else
					strImageSummary = ""
				end if								
				if rs("TypeID") <> 1 then
					if trim(ProductArray(i)) = "" then
						ProductArray(i) = "<b>Version: </b>" & strVersion & strImageSummary & "<BR><b>Distributions: </b>" & strDistribution '& "<BR><b>Target Notes: </b>" & rs("TargetNotes")
					else
						ProductArray(i) = ProductArray(i) & "<hR><b>Version: </b>" & strVersion & strImageSummary & "<BR><b>Distributions: </b>" & strDistribution '& "<BR><b>Target Notes: </b>" & rs("TargetNotes")
					end if
				else
					ProductArray(i) = ProductArray(i) &  strVersion & "<BR>"
				end if
			end if
		next

		if strLastType <> rs("Type") then
			Response.write "<TR bgcolor=beige><TD colspan=" & ProductCount + 1 & ">" &  rs("Type") & "</TD></TR>"
			strLastType = rs("Type")
		end if

		strLastRoot = rs("Root")
		
		rs.MoveNext
	loop
	
	blnAllDups = true
	for i = 0 to ubound(ProductArray) -1
		if replace(replace(ucase(trim(ProductArray(i)))," ",""),",","") <>  replace(replace(ucase(trim(ProductArray(i + 1 )))," ",""),",","") then
			blnAllDups = false
			exit for
		end if 
	next
	
	if (request("Report")= "2" and not blnAllDups) or request("Report")<> "2" then
		Response.Write "<TR><TD>" & strLastRoot & "</td>"
		for i = 0 to ubound(ProductArray)
			response.write "<TD>" & ProductArray(i) & "&nbsp;</TD>"
		next
		response.write "</TR>"
	end if
	response.write "</TABLE>"
	rs.Close
	
	set rs=nothing
	cn.Close
	set cn=nothing



%>

</BODY>
</HTML>

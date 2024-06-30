<script runat="server" language="vbscript">

	dim cn 
	dim rs 
	dim strSQL
	dim strOutput
	dim strRows
	dim strVersion
	dim strCheckRoot
	dim VendorID 
    dim RootID
    dim ActionID
    dim CatID

    strSQL = ""
    VendorID = request("VendorID")
    RootID = request("RootID")
    ActionID = request("ActionID")
    CatID = request("CatID")

    PulsarProductIDList = Replace(request("PulsarProductIDList"),":",",")
    LegacyProductIDList = Replace(request("LegacyProductIDList"),":",",")

	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
    if len(PulsarProductIDList) > 0 then
	    strSQl = "Select distinct v.ID, v.partnumber, r.id as RootID, r.name, v.version, v.revision, v.pass, v.modelnumber " & _
			     "from product_deliverable pd with (NOLOCK) " &_
                 "inner join deliverableversion v with (NOLOCK) on pd.deliverableversionid = v.id " &_
                 "inner join deliverableroot r with (NOLOCK) on r.id = v.deliverablerootid " &_
                 "inner join productversion_release pvr with (NOLOCK) on pvr.ProductVersionID = pd.ProductVersionID " &_
                 "inner join product_deliverable_release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.ID and pvr.ReleaseID = pdr.ReleaseID " & _
			     "where pvr.ID in (" & PulsarProductIDList & ") " & _
			     "and v.vendorid = " & VendorID & " " & _
			     "and r.categoryid = " & CatID & " " & _
			     "and pdr.SupplyChainRestriction = " & ActionID & " " & _
			     "and v.active=1 " & _
			     "and v.filename not like 'HFCN_%' "
    end if

    if len(LegacyProductIDList) > 0 then 
        if strSQL <> "" then
            strSQL = strSQL & "Union "
        end if
		
        strSQl = strSQl & "Select distinct v.ID, v.partnumber, r.id as RootID, r.name, v.version, v.revision, v.pass, v.modelnumber " & _
			     "from product_deliverable pd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK) " & _
			     "where pd.deliverableversionid = v.id " & _
			     "and r.id = v.deliverablerootid " & _
			     "and pd.productversionid in (" & LegacyProductIDList & ") " & _
			     "and v.vendorid = " & VendorID & " " & _
			     "and r.categoryid = " & CatID & " " & _
			     "and pd.SupplyChainRestriction = " & ActionID & " " & _
			     "and v.active=1 " & _
			     "and v.filename not like 'HFCN_%' "
         
    end if

    strSQl = strSQl & "order by r.name, v.id"
	
	rs.open strSQL ,cn,adOpenForwardOnly
	strOutput = ""
	strRows = ""
	do while not rs.EOF
		if err.number <> 0 then
			exit do
		end if
		strVersion = rs("Version") & ""
		if rs("Revision") & "" <> "" then
			strVersion = strVersion & ", " & rs("Revision")
		end if
		if rs("Pass") & "" <> "" then
			strVersion = strVersion & ", " & rs("Pass")
		end if
		
		if trim(rs("RootID")) = trim(RootID) then
			strCheckRoot = "checked"
		else
			strCheckRoot = ""
		end if
		
		strRows = strRows & "<TR>"
		strRows = strRows & "<TD><INPUT style=""WIDTH:16;HEIGHT:16"" type=""checkbox"" id=chkVersions name=chkVersions " & strCheckRoot & " value=""" & rs("ID") & """></TD>"
		strRows = strRows & "<TD>" & rs("PartNumber") & "&nbsp;</TD>"
		strRows = strRows & "<TD>" & server.HTMLEncode(rs("Name") & "") & "</TD>"
		strRows = strRows & "<TD>" & server.HTMLEncode(strVersion) & "</TD>"
		strRows = strRows & "<TD>" & server.HTMLEncode(rs("ModelNumber") & "") & "&nbsp;</TD>"
		strRows = strRows & "</TR>"
		rs.MoveNext
	loop
	rs.Close
	
	if strRows <> "" then
		strOutput = "<TABLE class=VersionList style=""WIDTH:100%"" border=1 bgcolor=white cellpadding=2 cellspacing=1>"
		strOutput = strOutput & "<TR bgcolor=gainsboro ><TD width=1><INPUT style=""WIDTH:16;HEIGHT:16"" type=""checkbox"" id=chkAllVersions name=chkAllVersions language=javascript onclick=""ToggleVersions();""></TD><TD><b>Part</b></TD><TD><b>Deliverable</b></TD><TD><B>HW,FW,Rev</B></TD><TD><B>Model</B></TD></TR>"
		strOutput = strOutput &	strRows
		strOutput = strOutput &	"</TABLE>"
	else
		strOutput = "<font size=2 color=red face=verdana>No Versions Found.</font>"
	end if

    response.Write strOutput
	
	set rs = nothing
	cn.Close
	set cn=nothing

</script>
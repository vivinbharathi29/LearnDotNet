<%@ Language=VBScript %>
<HTML>
.SummaryTH
<HEAD>
function DisplayTargetIssues(){
function CompareLines(strTable){
//-->
</SCRIPT>
<BODY  LANGUAGE=javascript onload="return window_onload()">
<%
	function ScrubSQL(strWords) 
		strWords=replace(strWords,"'","''")
		for i = 0 to uBound(badChars) 
	dim StartDate
	if request("ProdID") = "" then
		CurrentUser = lcase(Session("LoggedInUser"))
		if instr(currentuser,"\") > 0 then
			'Response.Write "<font size=2 face=verdana><u><b>Results</b></u></font><BR><BR></div>"
			rs.Close
			skuCount = 0
							"from regions r with (NOLOCK), Images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK) " & _
							"where r.ID = i.RegionID " & _
							"and d.Id = i.ImageDefinitionID " & _
							"and i.ID in( " & scrubsql(request("lstImage") ) & ") " & _
							"and o.id = d.osid " & _
							"order by r.geoid, r.DisplayOrder, i.id;"
					rs.Open strSQl,cn,adOpenStatic
					Set p = cm.CreateParameter("@DefinitionID", 3, &H0001)
					rs.CursorType = adOpenForwardOnly
					Set rs = cm.Execute 
			elseif request("lstImageDefinitions") <> "" then
							"from regions r with (NOLOCK), Images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK) " & _
							"where r.ID = i.RegionID " & _
							"and d.Id = i.ImageDefinitionID " & _
							"and i.ImageDefinitionID in( " & scrubsql(request("lstImageDefinitions") ) & ") " & _
							"and o.id = d.osid " & _
							"order by r.geoid, r.DisplayOrder, i.id;"
					rs.Open strSQl,cn,adOpenStatic
		'	rs.Open "spListImagesForProductAll " & request("ProdID"),cn,adOpenForwardOnly
				if rs("ImageSnapshotsSaved") and trim(rs("lockeddeliverableList") & "") <> "" then 'rs("StatusID") > 1 and  request("ImageDefinitionID") <> "" then
				
					strSQl = "Select ID, DeliverableName, Version, Revision, Pass, 1 as Preinstall, 0 as preload, 0 as ARCD, 0 as selectiverestore, 1 as inimage, '' as Images " & _
					rs2.open strSQL, cn,adOpenStatic
				else
				'if request("ImageDefinitionID") <> "" then
    			set cn2 = server.CreateObject("ADODB.Connection")
						strDelConveyor = strDelConveyor & "2" & trim(rs2("DeliverableName")) & " " &  rs2("Version")
					do while not rs2.EOF
						if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
	end if
%>
<TABLE border=0 bordercolor=Indigo cellspacing=1 cellpadding=2>
        dim i
        dim blnFound 
        blnFound = false
        
        for i = 0 to ubound(MyArray)
            if trim(lcase(MyArray(i))) = trim(lcase(strValue)) then
                blnFound = true
                exit for
            end if    
        next
        InArray = blnFound
    end function
    
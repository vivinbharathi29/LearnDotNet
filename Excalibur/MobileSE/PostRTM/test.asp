<html><body>
<%
    RootID = 11668
	dim cn 
	dim rs 
	dim strConnect
	dim strStreetNames
	dim SeriesArray
	dim strSeriesText
	
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	if instr(RootID, "*") > 0 then
		RootID = left(RootID, instr(RootID, "*")-1)
	end if
	
	rs.open "spGetProductList4Root " & RootID ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		'getProduct = "<SELECT style=""width:400"" id=cboTable name=cboTable><OPTION selected value=""No Product"">No product has picked up this deliverable</OPTION>"
		getProduct = "No product has picked up this deliverable"		
	else
	
		do while not rs.EOF
    		set rs2 = server.CreateObject("ADODB.recordset")
		    rs2.open "spListBrands4Product " & rs("ID"),cn,adOpenForwardOnly
		    strStreetNames = ""
		    response.Write rs("ID") & "<BR>"
		    do while not rs2.EOF
    			if trim(rs2("SeriesSummary") & "") <> ""  then
				    SeriesArray = split(rs2("SeriesSummary"),",")
				    for each strSeriesText in SeriesArray
    					strStreetNames = strStreetNames & ", " & rs2("StreetName2") & " " & strSeriesText 
				    next
			    end if	
			    rs2.MoveNext
		    loop		
		    rs2.Close
		    set rs2 = nothing
		
		
		    if trim(strStreetNames) = "" then
			    strStreetNames = "TBD"
		    else
			    strStreetNames = mid(trim(strStreetNames),2)
		    end if				

	
		    if trim(rs("PostRTMStatus")) = "0" then
		        response.Write rs("ID") & "*" & "R" & "*" & rs("Name") & " " & rs("Version") & "*" & rs("AllowSMR") & "*" &  rs("SEPMID") & "*" & strStreetNames & "^"
		    elseif trim(rs("PostRTMStatus")) = "1" then
		        response.Write rs("ID") & "*" & "C" & "*" & rs("Name") & " " & rs("Version") & "*" & rs("AllowSMR") & "*" &  rs("SEPMID") & "*" & strStreetNames & "^"
		    end if
		    
		    rs.MoveNext
		loop
	end if
	rs.Close
'	getProduct = getProduct & "</SELECT>"
	
	set rs = nothing
	cn.Close
	set cn=nothing
%>
</body></html>
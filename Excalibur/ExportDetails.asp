<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.ExpiresAbsolute = Now() %>
<% Response.ContentType = "application/vnd.ms-excel"%>

<!-- #include file = "includes/noaccess.inc" -->

<HTML>
<HEAD>
</HEAD>
<BODY>

<%
	if request("Query") = "" then
	
		Response.Write "<font size=2 face=verdana>Not enough information supplied to display this page.</font>"
	else
	
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


	dim rs
	dim cn
	dim strSQL
	dim i
	dim strStatus
	dim strType
	dim strColumns
	dim strStatusID
	dim strTypeID
	dim strDate
	dim strBiosChange
	dim strProductList
	dim strdescription
    dim strjustification

	strColumns = lcase(request("lstSelected"))
	if right(strColumns,1) <> "," then
		strcolumns = strcolumns & ","
	end if

	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	'strSQL = request("Query")
	strProductList=request("lstSelectedProd")

    strStatusID = Request("hidStatus")
    strTypeID = Request("hidType")
    strBiosChange = Request("hidBios")
    
    If Trim(strStatusID) = "" Then
	    strStatusID = mid(request("Query"),instr(request("Query")," ")+1)
	    strStatusID = mid(strStatusID,instr(strStatusID,",")+1)
	    strTypeID = left(strStatusID,instr(strStatusID,",")-1)
	    strStatusID = mid(strStatusID,instr(strStatusID,",")+1)
	End If
	strColumns =  "," &  replace(strColumns," ","")
	strColumns = ScrubSQL(strColumns)
	strColumns = replace(strcolumns, "," & "product" & "," ,"," & "v.DOTSName as Product" & ",")
	strColumns = replace(strcolumns, "," & "id" & "," ,"," & "i.ID" & ",")
	strColumns = replace(strcolumns, "," & "type" & "," ,"," & "i.Type" & ",")
	strColumns = replace(strcolumns, "," & "status" & "," ,"," & "i.Status" & ",")
	strColumns = replace(strcolumns, "," & "owner" & "," ,"," & "e.name as Owner" & ",")
	strColumns = replace(strcolumns, "," & "summary" & "," ,"," & "i.Summary" & ",")
	strColumns = replace(strcolumns, "," & "description" & "," ,"," & "i.Description" & ",")
    strColumns = replace(strcolumns, "," & "details" & "," ,"," & "i.details" & ",")
	strColumns = replace(strcolumns, "," & "created" & "," ,"," & "i.Created" & ",")
	strColumns = replace(strcolumns, "," & "targetdate" & "," ,"," & "i.TargetDate" & ",")
	strColumns = replace(strcolumns, "," & "actualdate" & "," ,"," & "i.ActualDate" & ",")
	strColumns = replace(strcolumns, "," & "coreteamrep" & "," ,"," & "ct.Name as CoreTeamRep" & ",")
	strColumns = replace(strcolumns, "," & "submitter" & "," ,"," & "i.Submitter" & ",")
	strColumns = replace(strcolumns, "," & "notify" & "," ,"," & "i.Notify" & ",")
	strColumns = replace(strcolumns, "," & "bto" & "," ,"," & "i.BTO" & ",")
	strColumns = replace(strcolumns, "," & "na" & "," ,"," & "i.NA" & ",")
	strColumns = replace(strcolumns, "," & "la" & "," ,"," & "i.LA" & ",")
	strColumns = replace(strcolumns, "," & "apd" & "," ,"," & "i.APD" & ",")
	strColumns = replace(strcolumns, "," & "ckk" & "," ,"," & "i.CKK" & ",")
	strColumns = replace(strcolumns, "," & "emea" & "," ,"," & "i.EMEA" & ",")
	strColumns = replace(strcolumns, "," & "gcd" & "," ,"," & "i.GCD" & ",")
	strColumns = replace(strcolumns, "," & "affectscustomers" & "," ,"," & "i.AffectsCustomers" & ",")
	strColumns = replace(strcolumns, "," & "onstatusreport" & "," ,"," & "i.OnStatusReport" & ",")
	strColumns = replace(strcolumns, "," & "lastmodified" & "," ,"," & "i.LastModified" & ",")
	strColumns = replace(strcolumns, "," & "justification" & "," ,"," & "i.Justification" & ",")
	strColumns = replace(strcolumns, "," & "btodate" & "," ,"," & "i.BTODate" & ",")
	strColumns = replace(strcolumns, "," & "ctodate" & "," ,"," & "i.CTODate" & ",")
	strColumns = replace(strcolumns, "," & "approvals" & "," ,"," & "i.Approvals" & ",")
	strColumns = replace(strcolumns, "," & "actions" & "," ,"," & "i.Actions" & ",")
	strColumns = replace(strcolumns, "," & "resolution" & "," ,"," & "i.Resolution" & ",")
	strColumns = replace(strcolumns, "," & "release" & "," ,"," & "'&nbsp;' + i.ProductVersionRelease as Release" & ",")

	strColumns = mid(strcolumns,2,len(strColumns) -2 )	
	'strSQL = "Select v.DOTSName as Product, i.id, i.Type, i.Status, e.name as Owner, i.Summary, i.Description, i.Created, i.TargetDate, i.ActualDate, ct.Name as CoreTeamRep, i.Submitter,i.notify,i.BTO,i.CTO,i.NA,i.LA,i.APD,i.CKK,i.EMEA,i.GCD, i.AffectsCustomers, i.OnStatusReport,i.LastModified,i.Justification, i.BTODate,i.CTODate, i.Approvals,i.Actions,i.resolution " & _
	strSQL = "Select " & strColumns & " " & _
	        " FROM dbo.ProductVersion AS v WITH (NOLOCK)  " & _
	        " INNER JOIN dbo.ProductFamily AS f WITH (NOLOCK) ON v.ProductFamilyID = f.ID  " & _
	        " INNER JOIN dbo.DeliverableIssues AS i WITH (NOLOCK)  " & _
	        " INNER JOIN dbo.Employee AS e WITH (NOLOCK) ON i.OwnerID = e.ID  " & _
	        " LEFT OUTER JOIN dbo.CoreTeamRep AS ct WITH (NOLOCK) ON i.CoreTeamRep = ct.ID ON v.ID = i.ProductVersionID " & _
		    " WHERE i.ProductVersionID in (" & ScrubSQL(strProductList) & ") " & _
		    " AND i.Type = " & strTypeID
	
	if strStatusID <> "0" And Trim(strStatusID) <> "" then
		if CLng(strstatusid) = 1 then
			strSQL = strSQL & " AND ((i.Status = 1 or i.Status = 3 or i.status = 6) or ((v.sustaining = 1 or  i.CoreTeamRep = 12) and i.status=4 and i.ecndate is null)) "
		elseif CLng(strStatusID) = 2 then
		    strSQL = strSQL & " AND ((i.Status = 2  or i.Status = 5) or ((v.sustaining = 1 or  i.CoreTeamRep = 12) and i.status=4 and i.ecndate is not null) or ((v.sustaining <> 1 and  i.CoreTeamRep <> 12) and  i.Status = 4))"
		end if
	end if
	
	If Trim(strBiosChange) <> "" Then
	    SELECT CASE strBiosChange
	        CASE "1"
	            strSQL = strSQL & " AND BiosChange = 1 "
	        CASE "0"
	            strSQL = strSQL & " AND (BiosChange = 0 OR ((COALESCE(ImageChange, 0) | COALESCE(SkuChange, 0) | COALESCE(ReqChange, 0) | COALESCE(DocChange,0) | COALESCE(CommodityChange, 0) | COALESCE(OtherChange, 0) = 1 AND BiosChange = 1)))"
	    END SELECT
    End If
    
    strSQL = strSQL & " Order By v.DOTSName, i.ID"

'response.Write request.QueryString & "<br>"
'response.Write request.Form & "<br>"
'response.Write strBiosChange & "<br>"
'response.Write strStatusID & "<br>"
'response.Write strTypeID & "<br>"
'response.Write strSQL
'response.End

	rs.Open strSQL,cn,adOpenForwardOnly
  
	if not(rs.EOF and rs.BOF) then
		Response.Write "<TABLE border=1>"
		if Request("chkHeader") = "on" then
			Response.Write "<TR  bgcolor=cornsilk>"
			for i = 0 to rs.Fields.count -1		
				'if instr(strcolumns,lcase(rs.Fields(i).Name) & "," ) > 0 then
					Response.Write "<TD>" & rs.Fields(i).Name & "</TD>"
				'end if
			next
			response.write "</TR>"	
		end if
		do while not rs.EOF 
			Response.Write "<TR>"
			for i = 0 to rs.Fields.count -1		
				'if instr(strcolumns,lcase(rs.Fields(i).Name) & "," ) > 0 then
				if lcase(rs.Fields(i).Name) = "status" then
					select case rs.Fields(i).Value
					case 1
						strStatus = "Open"
					case 2
						strStatus = "Closed"
					case 3
						strStatus = "Need More Information"
					case 4
						strStatus = "Approved"
					case 5
						strStatus = "Disapproved"
					case 6
						strStatus = "Investigating"
					case else
						strStatus = "N/A"
					end select
					Response.Write "<TD>" & strStatus & "</TD>"			
				elseif lcase(rs.Fields(i).Name) = "type" then
					select case rs.Fields(i).Value
					case 1
						strType = "Issue"
					case 2
						strType = "Action"
					case 3
						strType = "Change Request"
					case 4
						strType = "Status Note"
					case else
						strType = "N/A"					
					end select
					Response.Write "<TD>" & strType & "</TD>"
                elseif lcase(rs.Fields(i).Name) = "description" then
					if isnull(rs.Fields(i).value) then
						strdescription = "&nbsp;"	
					else
						strdescription = replace(rs.Fields(i).value, vbcrlf, "<br style='mso-data-placement:same-cell;'/>")
					end if
					Response.Write "<TD>" & strdescription & "</TD>"
                elseif lcase(rs.Fields(i).Name) = "justification" then
					if isnull(rs.Fields(i).value) then
						strjustification = "&nbsp;"	
					else
						strjustification = replace(rs.Fields(i).value, vbcrlf, "<br style='mso-data-placement:same-cell;'/>")
					end if
					Response.Write "<TD>" & strjustification & "</TD>"
				elseif lcase(rs.Fields(i).Name) = "created" then
					if isnull(rs.Fields(i).value) then
						strDate = "&nbsp;"	
					else
						strDate = formatdatetime(rs.Fields(i).value,vbshortdate)
					end if
					Response.Write "<TD>" & strDate & "</TD>"
				elseif lcase(rs.Fields(i).Name) = "actualdate" then
					if isnull(rs.Fields(i).value) then
						strDate = "&nbsp;"	
					else
						strDate = formatdatetime(rs.Fields(i).value,vbshortdate)
					end if
					Response.Write "<TD>" & strDate & "</TD>"
				elseif lcase(rs.Fields(i).Name) = "lastmodified" then
					if isnull(rs.Fields(i).value) then
						strDate = "&nbsp;"	
					else
						strDate = formatdatetime(rs.Fields(i).value,vbshortdate)
					end if
					Response.Write "<TD>" & strDate & "</TD>"
				elseif lcase(rs.Fields(i).Name) = "affectscustomers" then
					Response.Write "<TD>" & replace(replace(rs("AffectsCustomers") & "","1","Yes"),"0","No") & "</TD>"			
				else
					Response.Write "<TD>" & rs.Fields(i).value & "</TD>"
				end if
				'end if
			next		
			response.write "</TR>"
			rs.MoveNext
		loop
		Response.Write "</TABLE>"
	else
		Response.Write "Requested information not found"
	end if
	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing

end if
%>
</BODY>
</HTML>

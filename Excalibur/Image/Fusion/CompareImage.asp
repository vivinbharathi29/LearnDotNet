<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<STYLE>
TD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: ivory;
}
TH
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: cornsilk;
}

.SummaryTH
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: LightSteelBlue;
}

.SummaryTD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: gainsboro;
}
</STYLE>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	lblProcessing.style.display = "none";
	ReportTitle.style.display = "";
	//var HideTables = txtHideTables.value.split(",");
	//var i;
	//for(i=0;i<HideTables.length;i++)
//		document.all("DIV" + HideTables[i]).style.display="none";
}

function DisplayTargetIssues(){
	TargetIssuesRow.style.display = "";	
	ImageIssuesRow.style.display = "none";	
}

function DisplayImageIssues(){
	TargetIssuesRow.style.display = "none";	
	ImageIssuesRow.style.display = "";	
}

function CompareLines(strTable){
	var i;
		document.all("frmCompare" + strTable).submit();
}

//-->

</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="return window_onload()">
<b><font ID=lblProcessing face=verdana size=2>Processing. This may take several minutes.  Please wait...</font></b>

<%
	Server.ScriptTimeout = 1200
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

	dim StartDate
	StartDate = now()

	dim cn
	dim cn2
	dim cm
	dim p
	dim rs
	dim rs2
	dim strDash
	dim strSKU
	dim strOutBuffer
	dim strDelConveyor
	dim strSKURevision
	dim strSQL
	dim skuCount
	dim errorcount
	dim totalerrorcount
	dim TotalCompared
	dim strConveyor
	dim TableCount
	dim CurrentUserPartner
	dim strHideTables
	dim strPreinstallTeam
	dim strServerLocation
	dim SkuHeader
	dim CheckingSKUNumber
	dim strImageID
	dim TableHeaderDisplayed
	dim strCompareType
	dim CompareTypeUpdated
	dim blnUseSnapshotDeliverables
	dim ExcaliburSKUArray
	dim currentuserid
	dim ExtraConveyorHeaderDisplayed
	
	TableCount =0

	if request("ProdID") = "" or true then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")

		'Get User
		dim CurrentDomain

		CurrentUser = lcase(Session("LoggedInUser"))

		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetUserInfo"
		
	
		Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		p.Value = Currentuser
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		p.Value = CurrentDomain
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
	
		set cm=nothing
		if (rs.EOF and rs.BOF) then
			set rs = nothing
			set cn=nothing
			Response.Redirect "../NoAccess.asp?Level=0"
		else
			CurrentUserPartner = rs("PartnerID")
		    currentuserid = rs("ID")
		end if 
		rs.Close		
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersionName"
		
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
		strPreinstallTeam = 0
'		rs.Open "spGetProductVersionName " & request("ProdID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
		else

			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					set rs = nothing
					set cn=nothing
					
					Response.Redirect "../NoAccess.asp?Level=0"
				end if
			end if
			
			'Response.Flush
		
			strPreinstallTeam = rs("PreinstallTeam")
			
		if cint(strPreinstallTeam) = 1 then
    			strConveyor = "houbnbcvr01.auth.hpicorp.net"'"16.101.60.73"
	    		strServerLocation = "Houston"
		'elseif cint(strPreinstallTeam) = 3 then
			'	strConveyor = "SGPACCCVR01.auth.hpicorp.net"
			'	strServerLocation = "Singapore"
			'elseif cint(strPreinstallTeam) = 4 then
			'	strConveyor = "BRAHPQCVR01.auth.hpicorp.net"
			'	strServerLocation = "Brazil"
'			elseif cint(strPreinstallTeam) = 5 then
'				strConveyor = "SHGCDCCVR01.auth.hpicorp.net"
'				strServerLocation = "China"
			else
				strConveyor = "16.159.144.23"'"tpopsgcvr3.auth.hpicorp.net"'"tpopsgcvr2.auth.hpicorp.net"
				strServerLocation = "Taiwan"
			end if
			strCompareType = rs("Name") & ""
			strproduct = rs("Name") & ""
			'Response.Write "<DIV ID=ReportTitle style=""display:none""><font size=3 face=verdana><b><center>" &  rs("name") & "</b></font><BR><BR><font size=2 face=verdana>Compare Excalibur Image Definitions to the " & strServerLocation & " Conveyor Images</font></center><BR>"

			'Response.Write "<font size=2 face=verdana><u><b>Results</b></u></font><BR><BR></div>"

			rs.Close

			skuCount = 0
			if request("ImageDefinitionID") <> "" then
				if request("lstImage") <> "" then
					strSQL = "Select r.ID, r.Name, d.ImageSnapshotsSaved, i.skunumber as FullSkuNumber, i.lockeddeliverableList, r.Dash, r.CountryCode,d.skunumber,r.dockits,r.restoremedia, r.Keyboard, o.name as os, r.OSLanguage,r.kwl, r.Geo, i.Priority, r.geoid, r.DisplayName, PowerCord, OtherLanguage, OptionConfig, d.statusid, d.id as ImageDefID, i.ID as ImageID " & _
							"from regions r with (NOLOCK), Images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK) " & _
							"where r.ID = i.RegionID " & _
							"and d.Id = i.ImageDefinitionID " & _
							"and i.ID in( " & scrubsql(request("lstImage") ) & ") " & _
							"and o.id = d.osid " & _
							"order by r.geoid, r.DisplayOrder, i.id;"
					rs.Open strSQl,cn,adOpenStatic
				else
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spListImagesForDefinition"


					Set p = cm.CreateParameter("@DefinitionID", 3, &H0001)
					p.Value = request("ImageDefinitionID")
					cm.Parameters.Append p
					rs.CursorType = adOpenForwardOnly
					rs.LockType=AdLockReadOnly

					Set rs = cm.Execute 
					Set cm=nothing
				end if

			elseif request("lstImageDefinitions") <> "" then
					strSQL = "Select r.ID, r.Name, d.ImageSnapshotsSaved, i.skunumber as FullSkuNumber, i.lockeddeliverableList, r.Dash, r.CountryCode,d.skunumber,r.dockits,r.restoremedia, r.Keyboard, o.name as os, r.OSLanguage,r.kwl, r.Geo, i.Priority, r.geoid, r.DisplayName, PowerCord, OtherLanguage, OptionConfig, d.statusid, d.id as ImageDefID, i.ID as ImageID " & _
							"from regions r with (NOLOCK), Images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK) " & _
							"where r.ID = i.RegionID " & _
							"and d.Id = i.ImageDefinitionID " & _
							"and i.ImageDefinitionID in( " & scrubsql(request("lstImageDefinitions") ) & ") " & _
							"and o.id = d.osid " & _
							"order by r.geoid, r.DisplayOrder, i.id;"
					rs.Open strSQl,cn,adOpenStatic
			else
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spListImagesForProductAll"
		
				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = request("ProdID")
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing
			end if

		'	rs.Open "spListImagesForProductAll " & request("ProdID"),cn,adOpenForwardOnly
			'Process each image for the selected product
			if request("ImageDefinitionID") <> "" or request("lstImageDefinitions") <> ""  then
				CompareTypeUpdated = false
			else
				CompareTypeUpdated = true
				Response.Write "<DIV ID=ReportTitle style=""display:none""><font size=3 face=verdana><b><center>" &  strCompareType & "</b></font><BR><BR><font size=2 face=verdana>Compare Excalibur Image Definitions to IRS Product Drops</font></center><BR>"
				Response.Write "<font size=2 face=verdana><u><b>Results</b></u></font><BR><BR></div>"
			end if
        	redim ExcaliburSKUArray(100)
        	MySKUCount=0
			do while not rs.EOF 
				if request("ImageDefinitionID") <> "" or request("lstImageDefinitions") <> ""  then
			 		 strImageID = rs("ImageID") & ""
				else
			 		 strImageID = rs("ID") & ""
				end if
			  if isnumeric(rs("Priority")) and rs("OS") <> "RedFlag Linux" and rs("OS") <> "FreeDOS"  and rs("OS") <> "SuSE Linux" and (rs("StatusID") < 2 or request("ImageDefinitionID") <> "") then
				if not CompareTypeUpdated then
					if request("lstImageDefinitions") <> "" then
						strCompareType = strCompareType & " (Selected Image Definitions)"
					else
						strCompareType = strCompareType & " (" & rs("Skunumber") & ")"
					end if
					CompareTypeUpdated = true
					Response.Write "<DIV ID=ReportTitle style=""display:none""><font size=3 face=verdana><b><center>" &  strCompareType & "</b></font><BR><BR><font size=2 face=verdana>Compare Excalibur Image Definitions to IRS product Drops</font></center><BR>"
					Response.Write "<font size=2 face=verdana><u><b>Results</b></u></font><BR><BR></div>"
				end if
				skuCount = skuCount + 1
				strDash = trim(rs("Dash") & "")
				strSKU = trim(lcase(rs("SKUNumber") & ""))
				
				if strDash = "" or strSKU = "" then
					strSKU = strImageID
				else
					strDash = mid(strDash,2)
					strDash = left(strDash,len(strDash)-1)
					strSKU = replace(strSKU,"xx",strDash) 
				end if
	
				CheckingSKUNumber = strSKU
				strOutBuffer = ""
				'Get Excalibur deliverables

				if rs("ImageSnapshotsSaved") and trim(rs("lockeddeliverableList") & "") <> "" then 'rs("StatusID") > 1 and  request("ImageDefinitionID") <> "" then
					if instr(rs("lockeddeliverableList") & "",":")> 0 then
						strSQl = left(rs("lockeddeliverableList") & "",instr(rs("lockeddeliverableList") & "",":")-2)
					else
						strSQL = left(rs("lockeddeliverableList") & "",len(rs("lockeddeliverableList") & "") -1)
					end if
				
					strSQl = "Select ID, DeliverableName, Version, Revision, Pass, 1 as Preinstall, 0 as preload, 0 as ARCD, 0 as selectiverestore, 1 as inimage, '' as Images " & _
							 "From DeliverableVersion with (NOLOCK) " & _
							 "where id in (" & strSQL & ") "
						'	 Response.Write "<BR>" & strSQL
					rs2.open strSQL, cn,adOpenStatic

				else
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
										
					cm.CommandText = "spListDeliverablesInImage"
	

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = strImageID
					cm.Parameters.Append p
					Set p = cm.CreateParameter("@PINTest", 3, &H0001)
					if request("PINTest") = "1" then
						p.Value = 1
					else
						p.Value = 0
					end if	
					cm.Parameters.Append p
		
					rs2.CursorType = adOpenForwardOnly
					rs2.LockType=AdLockReadOnly
					Set rs2 = cm.Execute 
					Set cm=nothing
				end if
				'rs2.Open "spListDeliverablesInImage " & rs("ID"),cn,adOpenForwardOnly
				do while not rs2.EOF
					if rs2("categoryid") <> 168 and rs2("categoryid") <> 169 and rs2("categoryid") <> 170 and rs2("categoryid") <> 171 and rs2("categoryid") <> 143 and rs2("categoryid") <> 179 and rs2("categoryid") <> 137 and ( rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore") ) and ( trim(rs2("Images") & "") = "" or instr(", " & rs2("Images") & ",", ", " & strImageID & ",")>0  or instr( rs2("Images") , "(" & strImageID & "=")>0 )  then
					'	templist = templist & "," & rs2("ID")
						strOutbuffer = strOutbuffer & "1" & trim(rs2("DeliverableName")) & " " &  rs2("Version")
						if  rs2("Revision") & "" <> "" then
							strOutbuffer = strOutbuffer & " " & rs2("Revision")
						end if
						if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
							strOutbuffer = strOutbuffer & asc(lcase(rs2("Pass"))) -87
						else
							strOutbuffer = strOutbuffer & rs2("Pass")
						end if
						strOutbuffer = strOutbuffer & vbcrlf
					end if
					rs2.MoveNext
				loop
				rs2.close		

				'if request("ImageDefinitionID") <> "" then
				'	Response.Write templist
				'end if
			'Excalibur deliverables have been loaded into strOutBuffer for this image 
				'Load Conevyor Deliverables here

    			'response.write strConveyor
               ' response.Flush
                'set cn2 = server.CreateObject("ADODB.Connection")
				'cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=T61WKbK9n82R;Persist Security Info=True;User ID=excalibur_feed;Initial Catalog=irs;Data Source=HPIRSPROD02.auth.hpicorp.net;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;"
				'cn2.Open
				'cn2.CommandTimeout = 300
				
				'rs2.Open "select max(skurevision) as Rev from sku with (NOLOCK) where skunumber = '" & ScrubSQL(strSKU) & "' and locked <> 0;",cn2,adOpenForwardOnly
				'if not (rs2.EOF and rs2.BOF) then
				'	strSKURevision =  rs2("rev")
				'end if
				'rs2.Close
				
				strDelConveyor = ""				
				'Response.Write ">" & strConveyor
				if true then 'strSKURevision <> "" then
					'Converoy query1
					strSQl = "spListComponentsInIRSImage '12WWODA86##'"
					rs2.Open strSQl,cn,adOpenForwardOnly
					do while not rs2.EOF
						strDelConveyor = strDelConveyor & "2" & trim(rs2("Name")) & " " &  trim(rs2("Version"))
						if  rs2("Revision") & "" <> "" then
							strDelConveyor = strDelConveyor & " " & trim(rs2("Revision"))
						end if
						if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
							strDelConveyor = strDelConveyor & asc(lcase(trim(rs2("Pass")))) -87
						else
							strDelConveyor = strDelConveyor & trim(rs2("Pass"))
						end if
						strDelConveyor = strDelConveyor & vbcrlf
						rs2.MoveNext
					loop
					rs2.Close

						
					'Sort Conveyor Deliverables
					LineArray = Split(lcase(strDelConveyor), vbcrlf)
					strDelConveyor = ""
					For i = UBound(LineArray) - 1 To 0 Step -1
						For j = 0 To i
							If mid(LineArray(j),2) > mid(LineArray(j + 1),2) Then
								temp = LineArray(j + 1)
								LineArray(j + 1) = LineArray(j)
								LineArray(j) = temp
							End If
						Next
					Next
						
					'Remove Dups and append to excalibur list
					for i = lbound(LineArray) to ubound(LineArray)					
						if i = ubound(LineArray) then
							strOutbuffer = strOutbuffer & LineArray(i) & vbcrlf
						elseif linearray(i) <> linearray(i+1) then
							strOutbuffer = strOutbuffer & LineArray(i) & vbcrlf
						end if
					next

				end if
				
				
'				cn2.close
'				set cn2 = nothing	
				
		'Compare Excalibur to Conveyor

			'Sort
				LineArray = Split(lcase(strOutBuffer), vbcrlf)
				if  UBound(LineArray) > 0 then
					TotalCompared = TotalCompared + UBound(LineArray)
				end if
				strOutBuffer = ""
				For i = UBound(LineArray) - 1 To 0 Step -1
					For j = 0 To i
						If mid(LineArray(j),2) > mid(LineArray(j + 1),2) Then
							temp = LineArray(j + 1)
    						LineArray(j + 1) = LineArray(j)
							LineArray(j) = temp
						End If
					Next
				Next
				
				'DisplayDifferences
				TableCount=TableCount + 1
				i = LBound(LineArray)
			'	if UBound(LineArray) > 0 then
			'		Response.Write "<DIV ID=DIV" & TableCount & "><form action=""ShowDifference.asp""  method=post target=""_blank"" id=frmCompare" & TableCount & ">"
			'		Response.Write "<font size=2 face=verdana><b>" & CheckingSKUNumber  & "</b></font>" ' : Discrepancies Found
			'	    Response.write "<TABLE border=1 bordercolor=tan bgcolor=ivory>"
			'		Response.write  "<TR><TH align=left><a href=""javascript: CompareLines(" & TableCount & ");"">Compare</a></TH><TH align=left>System</TH><TH align=left>Deliverable</TH></TR>"
			'	else
			'		strGoodSKUs = strGoodSKUs & "," & CheckingSKUNumber
			'	end if
				TableHeaderDisplayed = false	
				do while i < UBound(LineArray)
					if mid(lcase(LineArray(i)),2) = mid(lcase(LineArray(i+1)),2) then
						i=i+2
					else
						'Response.Write TableHeaderDisplayed & "<BR>"
						if not TableHeaderDisplayed then
							Response.Write "<DIV ID=DIV" & TableCount & "><form action=""ShowDifference.asp""  method=post target=""_blank"" id=frmCompare" & TableCount & ">"
							Response.Write "<font size=2 face=verdana><b>Image ID: " & CheckingSKUNumber  & " - Discrepancies Found</b></font>" 
						    Response.write "<TABLE width=100% border=1 bordercolor=tan bgcolor=ivory>"
							Response.write  "<TR><TH align=left><a href=""javascript: CompareLines(" & TableCount & ");"">Compare</a></TH><TH align=left>System</TH><TH align=left>Deliverable</TH></TR>"
							TableHeaderDisplayed = true
						end if
						if left(LineArray(i),1) = "1" then
							Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""E" & lcase(mid(LineArray(i),2))  & """></td>"
							Response.Write "<TD>Excalibur</td>"
							Response.Write "<TD>" & lcase(mid(LineArray(i),2))  & "</td></TR>"
						elseif left(LineArray(i),1) = "2" then
							Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""C" & lcase(mid(LineArray(i),2))  & """></td>"
							Response.Write "<TD>IRS</td>"
							Response.Write "<TD>" & lcase(mid(LineArray(i),2))  & "</td></TR>"
						end if
						totalErrorCount = totalErrorCount + 1
						i=i+1
					end if
				loop
				
				if not TableHeaderDisplayed then
					strGoodSKUs = strGoodSKUs & "," & CheckingSKUNumber
				end if
				
				if i=ubound(LineArray) then
						if not TableHeaderDisplayed then
							Response.Write "<DIV ID=DIV" & TableCount & "><form action=""ShowDifference.asp""  method=post target=""_blank"" id=frmCompare" & TableCount & ">"
							Response.Write "<font size=2 face=verdana><b>" & CheckingSKUNumber  & " : Discrepancies Found</b></font>" 
						    Response.write "<TABLE width=100% border=1 bordercolor=tan bgcolor=ivory>"
							Response.write  "<TR><TH align=left><a href=""javascript: CompareLines(" & TableCount & ");"">Compare</a></TH><TH align=left>System</TH><TH align=left>Deliverable</TH></TR>"
							TableHeaderDisplayed = true
						end if
					if left(LineArray(i),1) = "1" then
						Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""E" & lcase(mid(LineArray(i),2))  & """></td>"
						Response.Write "<TD>Excalibur</td>"
						Response.Write "<TD>" & lcase(mid(LineArray(i),2)) & "</td></TR>"
					elseif left(LineArray(i),1) = "2" then
						Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""C" & lcase(mid(LineArray(i),2))  & """></td>"
						Response.Write "<TD>IRS</td>"
						Response.Write "<TD>" & lcase(mid(LineArray(i),2)) & "</td></TR>"
					end if
					totalErrorCount = totalErrorCount + 1
				end if
				if TableHeaderDisplayed then
					Response.Write "</TABLE></form><BR></DIV>"
				end if
				
				if totalErrorCount = 0 then
					'Response.Write "...No Difference Found<BR>"
					strHideTables = strHideTables & "," & TableCount
				end if
			   end if
				rs.MoveNext
			loop
			rs.Close
		
			if strGoodSKUs <> "" then
				strGoodSKUs = mid(strGoodSKUs,2)
				if totalErrorCount > 0 then
					Response.Write "<Font face=verdana size=2>No discrepancies found on any other " & strCompareType & " SKUs.<BR><BR></font>"
				else
					Response.Write "<Font face=verdana size=2>No discrepancies found on any " & strCompareType & " SKUs.<BR><BR></font>"
				end if
			elseif skuCount = 0 then
				if not CompareTypeUpdated then
					Response.Write "<DIV ID=ReportTitle style=""display:none""></div>"
				end if	
				Response.Write "<Font face=verdana size=2>No active images found for this product.<BR><BR></font>"
			end if

            
			rs.Open "spListDeliverablesWithMissingImages " & request("ProdID"),cn,adOpenStatic
			strMissingImagesCount = 0
			strMissingImagesRows = ""
			do while not rs.EOF
				strMissingImagesCount = strMissingImagesCount + 1
				strVersion = rs("Version") & ""
				if trim(rs("Revision") & "") <> "" then
					strversion = strversion & "," & rs("Revision")				
				end if
				if trim(rs("Pass") & "") <> "" then
					strversion = strversion & "," & rs("Pass")				
				end if
				strMissingImagesRows = strMissingImagesRows & "<TR><TD>" & rs("deliverablename") & " [" & strVersion & "]</TD>"
                if rs("Targeted") then
                    strMissingImagesRows = strMissingImagesRows & "<TD>Yes</TD>"
                else
                    strMissingImagesRows = strMissingImagesRows & "<TD>No</TD>"
                end if
                if rs("InPINImage") then
                    strMissingImagesRows = strMissingImagesRows & "<TD>Yes</TD>"
                else
                    strMissingImagesRows = strMissingImagesRows & "<TD>No</TD>"
                end if
                if rs("InImage") then
                    strMissingImagesRows = strMissingImagesRows & "<TD>Yes</TD>"
                else
                    strMissingImagesRows = strMissingImagesRows & "<TD>No</TD>"
                end if
                strMissingImagesRows = strMissingImagesRows & "<TD>" & rs("PreinstallStep") & "</TD>"
                strMissingImagesRows = strMissingImagesRows & "</TR>"
				rs.MoveNext
			loop
			rs.Close


			rs.Open "spListTargetedVersionsNotInImage " & request("ProdID"),cn,adOpenStatic
			strMissingTargetCount = 0
			strMissingTargetRows = ""
			do while not rs.EOF
				strMissingTargetCount = strMissingTargetCount + 1
				strVersion = rs("Version") & ""
				if trim(rs("Revision") & "") <> "" then
					strversion = strversion & "," & rs("Revision")				
				end if
				if trim(rs("Pass") & "") <> "" then
					strversion = strversion & "," & rs("Pass")				
				end if
				strMissingTargetRows = strMissingTargetRows & "<TR><TD>" & rs("deliverablename") & " [" & strVersion & "]</TD></TR>"'<TD>" & rs("PreinstallStep") & "</TD></TR>"
				rs.MoveNext
			loop
			rs.Close

			
		end if'product found
		
		
		set rs = nothing
		set rs2 = nothing
		set cn = nothing

	end if

%>
</b>
<font size=2 face=verdana>
<%
Response.Write "<font size=2 face=verdana><u><b>Summary</b></u></font><BR><BR>"
if totalErrorCount > TotalCompared then
	totalErrorCount = TotalCompared
end if
if totalErrorCount = "" then
	totalErrorCount = "0"
end if
if TotalCompared = "" then
	TotalCompared = "0"
end if
%>

<TABLE border=0 bordercolor=Indigo cellspacing=1 cellpadding=2>
<TR><TD class=SummaryTH>SKUs Checked:</b></td><TD class=SummaryTD>&nbsp;<%=skuCount%>&nbsp;</td></tr>
<TR><TD class=SummaryTH>Deliverables&nbsp;Checked:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></td><TD class=SummaryTD>&nbsp;<%=TotalCompared%>&nbsp;</td></tr>
<TR><TD class=SummaryTH>Discrepancies Found:</b></td><TD class=SummaryTD>&nbsp;<%=totalErrorCount%>&nbsp;</td></tr>
<%
	if TotalCompared <> 0  then
		Response.Write "<TR><TD class=SummaryTH>Discrepancy Rate:</b></td><TD class=SummaryTD>&nbsp;" & round((TotalErrorCount/TotalCompared)*100,2) & "%&nbsp;</td></tr>"
	end if
%>
<TR><TD class=SummaryTH>Compare Time:</b></td><TD class=SummaryTD>&nbsp;<%=datediff("s",StartDate,now())%> seconds&nbsp;</td></tr>
</table><BR>
<%if request("ImageDefinitionID") <> "" then%>
	*This report does not compare inactive images, RedFlag Linux images, SuSE Linux images, or FreeDOS images
<%else%>
	*This report does not compare inactive images, RedFlag Linux images, SuSE Linux images, FreeDOS images, or images that have been released to the factory
<%end if%>
<%if request("lstImage") <> "" then%>
	<BR>*This report is only comparing selected images.
<%end if%>
</font>
<%
	if strHideTables <> "" then
		strHideTables = mid(strHideTables,2)
		strHideTables=""
	end if
%>
<font size=2 face=verdana>
<%
Response.Write "<font size=2 face=verdana><u><b><BR><BR>Potential " & strProduct & " Issues</b></u></font><BR><BR>"
%>
<TABLE border=0 bordercolor=Indigo cellspacing=1 cellpadding=2>
<TR><TD class=SummaryTH>Targeted Versions Not In Images:</td><TD class=SummaryTD>&nbsp;
<%if clng(strMissingTargetCount) > 0 then%>
	<a href="javascript: DisplayTargetIssues();"><%=strMissingTargetCount%></a>
<%else%>
	0
<%end if%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
<TR><TD class=SummaryTH>No Images Selected for Deliverable:</td><TD class=SummaryTD>&nbsp;
<%if clng(strMissingImagesCount) > 0 then%>
	<a href="javascript: DisplayImageIssues();"><%=strMissingImagesCount%></a>
<%else%>
	0
<%end if%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>

</table><BR>

<TABLE ID=TargetIssuesRow Style="display:none" width=100% border=1 bordercolor=tan bgcolor=ivory>
<TR><TH align=left><b>Deliverable</b></TH><!--<TH align=left><b>In Preinstall Section</b></TH>--></TR>
<%=strMissingTargetRows%>
</Table>

<TABLE ID=ImageIssuesRow Style="display:none" width=100% border=1 bordercolor=tan bgcolor=ivory>
<TR><TH align=left><b>Deliverable</b></TH><TH align=left><b>Targeted</b></TH><TH align=left><b>In&nbsp;PIN&nbsp;Image</b></TH><TH align=left><b>In&nbsp;Image</b></TH><TH align=left><b>PreInstall&nbsp;Step</b></TH></TR>
<%=strMissingImagesRows%>
</Table>

<%

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListImagesForProductAll"

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ProdID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

    do while not rs.eof
		strDash = trim(rs("Dash") & "")
		strSKU = trim(lcase(rs("SKUNumber") & ""))
		if strDash = "" or strSKU = "" then
			strSKU = strImageID
		else
			strDash = mid(strDash,2)
			strDash = left(strDash,len(strDash)-1)
			strSKU = replace(strSKU,"xx",strDash) 
		end if

	    if MySKUCount >= ubound(ExcaliburSKUArray) then
	        redim preserve ExcaliburSKUArray(MySKUCount+100)
    	 end if
		 ExcaliburSKUArray(MySKUCount) = lcase(trim(strSKU))
		 MySKUCount = MySKUCount + 1
        rs.movenext
    loop
    rs.close
    set rs = nothing

    set cn = server.CreateObject("ADODB.Connection")
    set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=" & strConveyor & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;"
	cn.Open

	strSQl =  "select distinct sku.skunumber, prod.prodname " & _
              "from bom with (NOLOCK) , prod with (NOLOCK), sku with (NOLOCK) " & _
              "where prod.ProdName like '%" & strproduct & "%' " & _
              "and prod.prodkey = bom.prodkey " & _
              "and sku.bomkey = bom.bomkey " & _
              "and sku.enabled = 1;"
    rs.open strSQl,cn,adOpenStatic
    ExtraConveyorHeaderDisplayed = false
    do while not rs.eof
        if not InArray(ExcaliburSKUArray,rs("skunumber") & "") then
            if not ExtraConveyorHeaderDisplayed then
                response.write "<BR><b><font size=2 face=verdana>SKUs in IRS that are not defined in Excalibur for this product.</b></font>"
                response.write "<table width=""100%"" bgcolor=ivory border=1 bordercolor=tan><TR bgcolor=cornsilk><TD width=""150""><b>SKU</b></TD><TD><b>Product</b></TD></tr>"
                ExtraConveyorHeaderDisplayed = true
            end if
            response.write "<TR><TD>" & rs("SkuNumber") & "</TD><TD>" & rs("Prodname") & "</TD></tr>"
        end if
        rs.movenext
    loop
    if ExtraConveyorHeaderDisplayed then
        response.Write "</table>"
    end if
    rs.close
    set rs = nothing
    cn.close
    set cn = nothing
'    if currentuserid = 31 then
'        for i = 0 to myskucount-1
'            response.Write "<BR>" & i & ":" & ExcaliburSKUArray (i)
'        next
'    end if
    
    function InArray(MyArray, strValue)
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
    
%>
<INPUT type="hidden" id=txtHideTables name=txtHideTables value=<%=strHideTables%>>
</BODY>
</HTML>
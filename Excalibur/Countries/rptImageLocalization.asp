<%@ Language=VBScript %>

	<%
	
'  Response.Buffer = True
 ' Response.ExpiresAbsolute = Now() - 1
  'Response.Expires = 0
  'Response.CacheControl = "no-cache"
	  
	if request("FileType")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("FileType")= 2 then
		Response.ContentType = "application/msword"
	end if

	%>
	
<HTML>
<HEAD>
<TITLE>Image Localization Matrix</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
//	if (typeof(txtMaxUpdated) != "undefined" && typeof(lblImageCount) != "undefined" && typeof(lblModDate) != "undefined" && typeof(txtImageCount ) != "undefined")
//		{
//		lblModDate.innerText = txtMaxUpdated.value;
//		}
}

function FilterMatrix(strNewFilter) {
	window.location.href = document.location.href + strNewFilter;
}

function Export(strID){
	window.open (window.location.href + "&FileType=" + strID);
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
a.Filter:visited
{
	COLOR: blue;
	TEXT-DECORATION: none;
}
a.Filter
{
	COLOR: blue;
	TEXT-DECORATION: none;
}
a.Filter:hover
{
	BACKGROUND-COLOR:thistle;
	COLOR: black;
}
a:visited
{
	COLOR:blue;
}
a:hover
{
	COLOR: red;
}


</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%

	dim cn
	dim rs
	dim cm
	dim p
	dim blnOK
	dim strProductName
	dim strCountNote
	dim DisplayedID
	dim CurrentUserID
	
  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	  'Create a recordset
	set rs = server.CreateObject("ADODB.recordset")
	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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
        CurrentUserID = rs("ID")
    end if 
    rs.Close

	
	dim strProductPartner
	strProductPartner = ""
	
	if request("ProdID") <> "" and isnumeric(request("ProdID")) then
		DisplayedID = request("ProdID")
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spGetProductVersion " & request("ProdID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<BR>Unable to find the selected product.<BR>"
			blnOK = false
		else
			strProductName = rs("name") & " " & rs("Version")
			strProductPartner = rs("PartnerID")
			blnOK = true
		end if
		rs.Close
	elseif request("Product") <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersionByName"
		

		Set p = cm.CreateParameter("@Name", 200, &H0001,255)
		p.Value = request("Product")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

	
'		rs.Open "spGetProductVersionByName '" & request("Product") & "'",cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<BR>Unable to find the selected product.<BR>"
			blnOK = false
		else
			DisplayedID = rs("ID") & ""
			strProductName = rs("name") & " " & rs("Version")
			strProductPartner = rs("PartnerID")
			blnOK = true
		end if
		rs.Close
	else
		Response.Write "<BR>Unable to find the selected product.<BR>"
		blnOK = false
	end if

	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strProductPartner) <> trim(CurrentUserPartner) then
			set rs = nothing
			set cn=nothing
			
			Response.Redirect "../NoAccess.asp?Level=0"
		end if
	end if



	if blnOK then 
	
	
	dim strFilterList
	
	strFilterList = ""
	
	 if request("Type") <> "" then
		strFilterList = "," & request("Type")
	 end if

	 if request("SW") <> "" then
		strFilterList = strFilterList & "," & request("SW") & " Bundle"
	 end if

	 if request("Brand") <> "" then
		strFilterList = strFilterList & "," & request("Brand")
	 end if

	 if request("OpSys") <> "" then
		strFilterList = strFilterList & "," & request("OpSys") 
	 end if

	 if request("Status") <> "" then
		strFilterList = strFilterList & "," & request("Status") 
	 end if

	 if request("RTMDate") <> "" then
		strFilterList = strFilterList & "," & request("RTMDate") 
	 end if

	
	strCountNote = ""
	if strFilterList <> "" then
		strFilterList = mid(strFilterList,2)
		strCountNote = "<br><BR><font color=red size=1 face=verdana>Not All Images Displayed. "
		if request("FileType") = "" then
			strCountNote = strCountNote & "<a href=""rptImageLocalization.asp?ProdID=" & DisplayedID & """>Show All</a>"
		end if
		strCountNote = strCountNote & "</font>"	
	end if
	
%>

<center><font face=verdana size=3><b> <%=strProductname & " Image Localization Matrix"%></b><BR><BR></font></center>
<%
	if trim(request("ProdID")) = "267" then
	%>
		<center><font size=2 face=verdana color=red>Note: These images are shared with Thruman 1.1</font></center><BR>
	<%
	elseif trim(request("ProdID")) = "268" then
	%>
		<center><font size=2 face=verdana color=red>Note: The consumer images are shared with Ford 1.1</font></center><BR>
	<%
	end if
%>
<%if strFilterList <> "" then%>
<font size=2 face=verdana><center>Images Displayed: <%=strFilterList%></center><BR></font>
<%end if%>
<font size=2 face=verdana><center><label ID=lblModDate><%=formatdatetime(now,vbshortdate)%></label></center><BR></font>
<%if request("FileType") = "" then%>
<table width=100% border=0><tr><td align=right><font size=1 face=verdana>	Export: <a href="javascript: Export(1);">Excel</a> | <a href="javascript: Export(2);">Word</a> | <a href="javascript: Export(1);">PDD Export</a></td></tr></table>
<%end if%>
<%

	Dim PreRow
	Dim MidRow
	Dim PostRow
	dim strRow
	dim strModelRow
	dim strOSRow
	dim strAppsRow
	dim strTypeRow
	dim strSKURow
	dim strModifiedRow
	dim strStatusRow
	dim strRTMDateRow
	dim strCommentsRow
	dim strHideRow
	dim blnFound
	dim MaxUpdated
	dim strImageIDList
	dim ImageArray
	dim OutArray
	dim strLastGeo
	dim strRegions
	dim ImageDefCount
	dim strLastRegion
	dim i
	dim blnFirst
	dim strRegionDef
	dim ImageCountArray
	dim ImageCountP1Array
	dim ImageCountP2Array
	dim ImageCountP3Array
	dim ImageCountP4Array
	dim ImageCountP5Array
	dim TotalImageCount
	dim strShading
	dim blnInclude
	dim strFilterColor
	dim strShowOption
	dim strTierheader
	dim strLangSub
	dim strTemp
	dim intNativeCount
	dim blnFoundImages
	
	strFilterColor = "Thistle"
	
	strImagesDisplayed = ""

	if request("Brand") = "" then
		strModelRow = "<TR BGCOLOR=white>"
	else
		strModelRow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
	if request("OpSys") = "" then
		strOSRow = "<TR BGCOLOR=white>"
	else
		strOSRow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
	if request("SW") = "" then
		strAppsRow = "<TR BGCOLOR=white>"
	else
		strAppsRow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
	if request("Type") = "" then	
		strTypeRow = "<TR BGCOLOR=white>"
	else
		strTypeRow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
	strSKURow = "<TR BGCOLOR=white>"
  	strModifiedRow= "<TR BGCOLOR=white>"
	if request("Status") = "" then
  		strStatusRow = "<TR BGCOLOR=white>"
	else
  		strStatusRow = "<TR BGCOLOR=" & strFilterColor & ">"
  	end if
	if request("RTMDate") = "" then
  		strRTMDateRow = "<TR BGCOLOR=white>"
	else
  		strRTMDateRow = "<TR BGCOLOR=" & strFilterColor & ">"
  	end if
  	strHideRow= "<TR>"
	blnFound = false


'	if request("Display") = "SKU" then
'		strShowOption = "<br><br><INPUT type=""radio"" id=optPriority name=optDisplay><font size=1 face=verdana>&nbsp;Show Priorities<br><INPUT type=""radio"" checked id=optSKU name=optDisplay><font size=1 face=verdana>&nbsp;Show SKU Numbers</font>" 
'	else
'		strShowOption = "<br><br><INPUT type=""radio"" checked id=optPriority name=optDisplay><font size=1 face=verdana>&nbsp;Show Priorities<br><INPUT type=""radio"" id=optSKU name=optDisplay><font size=1 face=verdana>&nbsp;Show SKU Numbers</font>" 
'	end if
	strShowOption = ""	
	TotalImageCount = 0
	ImageDefCount = 0
	strImageIDList = ""
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListImageDefinitionsByProduct"
		
'	Set p = cm.CreateParameter("@ID", 3, &H0001)
'	p.Value = DisplayedID
'	cm.Parameters.Append p

'	Set p = cm.CreateParameter("@HideInactive", 3, &H0001)
'	p.Value = 1
'	cm.Parameters.Append p

'	if DisplayedID = 268 then
'		Set p = cm.CreateParameter("@HideInactive", 3, &H0001)
'		p.Value = 267
'		cm.Parameters.Append p
'	end if	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = DisplayedID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Report", 3, &H0001)
	p.Value = 1
	cm.Parameters.Append p

	if DisplayedID = 268 then
		Set p = cm.CreateParameter("@IncludeProductID", 3, &H0001)
		p.Value = 267
		cm.Parameters.Append p
	end if	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	
	'rs.open "spListImageDefinitionsByProduct " & DisplayedID,cn,adOpenForwardOnly
	do while not rs.EOF
		blnInclude = true
		if request("Type") <> "" then
			if lcase(trim(rs("ImageType"))) <> lcase(trim(request("TYPE"))) then
				blnInclude = false
			end if
		end if
	
		if request("SW") <> "" then
			if lcase(trim(rs("SWType"))) <> lcase(trim(request("SW"))) then
				blnInclude = false
			end if
		end if

		if request("Brand") <> "" then
			if lcase(trim(rs("Brand"))) <> lcase(trim(request("Brand"))) then
				blnInclude = false
			end if
		end if
	
		if request("OpSys") <> "" then
			if lcase(trim(rs("OS"))) <> lcase(trim(request("OpSys"))) then
				blnInclude = false
			end if
		end if

		if request("Status") <> "" then
			if lcase(trim(rs("Status"))) <> lcase(trim(request("Status"))) then
				blnInclude = false
			end if
		end if

		if request("RTMDate") <> "" then
			if lcase(trim(rs("RTMDate") & "")) <> lcase(trim(request("RTMDate"))) then
				if not(lcase(trim(request("RTMDate"))) = "none" and lcase(trim(rs("RTMDate") & "")) = "") then
					blnInclude = false
				end if
			end if
		end if
	
	
		if blnInclude then
			'if rs("BGColor") & "" <> "" then
			'	strShading = "BGCOLOR=" & rs("BGColor")
			'else
				strShading = ""
			'end if
			strImageIDList = strImageIDList & "," & rs("ID")
			if request("Brand") = "" then
				if request("FileType") = "" then
					strModelRow = strModelRow & "<TD " & strShading & "><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Brand=" & rs("Brand") & "')"">" & rs("Brand") & "</a></font></td>"
				else
					strModelRow = strModelRow & "<TD " & strShading & "><font size=1 face=verdana>" & rs("Brand") & "</font></td>"
				end if
			else
				strModelRow = strModelRow & "<TD " & strShading & "><font size=1 face=verdana>" & rs("Brand") & "</font></td>"
			end if
			if request("OpSys") = "" then
				if request("FileType") = "" then
					strOSRow = strOSRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&OpSys=" & rs("OS") & "')"">"  & rs("OS") & "</a></font></td>"
				else
					strOSRow = strOSRow & "<TD><font size=1 face=verdana>"  & rs("OS") & "</font></td>"
				end if
			else
				strOSRow = strOSRow & "<TD><font size=1 face=verdana>" & rs("OS") & "</font></td>"
			end if
			if request("SW") = "" then
				if request("FileType") = "" then
					strAppsRow = strAppsRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&SW=" & rs("SWType") & "')"">" & rs("SWType") & "</a></font></td>"
				else
					strAppsRow = strAppsRow & "<TD><font size=1 face=verdana>" & rs("SWType") & "</font></td>"
				end if
			else
				strAppsRow = strAppsRow & "<TD><font size=1 face=verdana>" & rs("SWType") & "</font></td>"
			end if

			if request("Staus") = "" then
				if request("FileType") = "" then
					strStatusRow = strStatusRow & "<TD nowrap><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Status=" & rs("Status") & "')"">"  & rs("Status") & "</a></font></td>"
				else
					strStatusRow = strStatusRow & "<TD nowrap><font size=1 face=verdana>"  & rs("Status") & "</font></td>"
				end if
			else
				strStatusRow = strStatusRow & "<TD nowrap><font size=1 face=verdana>" & rs("Status") & "</font></td>"
			end if

			if request("RTMDate") = "" then
				if request("FileType") = "" then
					if trim(rs("RTMDate")& "") = "" then
						strRTMDateRow = strRTMDateRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&RTMDate=none')"">"  & rs("RTMDate") & "&nbsp;</a></font></td>"
					else
						strRTMDateRow = strRTMDateRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&RTMDate=" & rs("RTMDate") & "')"">"  & rs("RTMDate") & "&nbsp;</a></font></td>"
					end if
				else
					strRTMDateRow = strRTMDateRow & "<TD align=left><font size=1 face=verdana>"  & rs("RTMDate") & "</font></td>"
				end if
			else
				strRTMDateRow = strRTMDateRow & "<TD align=left><font size=1 face=verdana>" & rs("RTMDate") & "</font></td>"
			end if
			
			if request("Type") = "" then
				if request("FileType") = "" then
					strTypeRow = strTypeRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Type=" & rs("ImageType") & "')"">" & rs("ImageType") & "</a></font></td>"
				else
					strTypeRow = strTypeRow & "<TD><font size=1 face=verdana>" & rs("ImageType") & "</font></td>"
				end if
			else
				strTypeRow = strTypeRow & "<TD><font size=1 face=verdana>" & rs("ImageType") & "</font></td>"
			end if
			strSKURow = strSKURow & "<TD nowrap><font size=1 face=verdana>" & rs("SKUNumber") & "&nbsp;</font></td>"
			strCommentsRow = strCommentsRow & "<TD><font size=1 face=verdana>" & rs("Comments") & "&nbsp;</font></td>"
			strModifiedRow = strModifiedRow & "<TD align=left><font size=1 face=verdana>" & formatdatetime(rs("Modified"),vbshortdate) & "</font></td>"
			strHideRow = strHideRow & "<TD align=center><font size=1 face=verdana><a href="""">Hide</a></font></td>"
			if MaxUpdated = "" then
				MaxUpdated = formatdatetime(rs("Modified"),vbshortdate)
			elseif datediff("d",MaxUpdated,rs("Modified")) > 0 then
				MaxUpdated = formatdatetime(rs("Modified"),vbshortdate)
			end if
			blnFound= true
			strTotalCells = strTotalCells & "<TD><font size=2 face=verdana></font></TD>"
			ImageDefCount = ImageDefCount + 1
		end if
		rs.movenext
	loop
	rs.Close
	if len(strImageIDList) > 0 then
		strImageIDList = mid(strImageIDList,2) 'Strip comma
	end if

	ImageArray = split (strImageIDList,",")
	Redim OutArray(Ubound(ImageArray))
	Redim PriorityArray(Ubound(ImageArray))
	Redim ImageCountArray(Ubound(ImageArray))
	Redim ImageCountP1Array(Ubound(ImageArray))
	Redim ImageCountP2Array(Ubound(ImageArray))
	Redim ImageCountP3Array(Ubound(ImageArray))
	Redim ImageCountP4Array(Ubound(ImageArray))
	Redim ImageCountP5Array(Ubound(ImageArray))

	
	
	if blnFound then
		strLastGeo = ""
		strLastRegion = ""
		for i = lbound(OutArray) to ubound(OutArray)
			OutArray(i) = "&nbsp;"
			PriorityArray(i) = ""
			ImageCountP1Array(i) = 0
			ImageCountP2Array(i) = 0
			ImageCountP3Array(i) = 0
			ImageCountP4Array(i) = 0
			ImageCountP5Array(i) = 0
			ImageCountArray(i) = 0
		next

		blnFirst = true
		
		dim strSKUBGColor
		
		strLangSub = ""
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListImagesForProduct"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = DisplayedID
		cm.Parameters.Append p
	
		if DisplayedID = 268 then
			Set p = cm.CreateParameter("@IncludeID", 3, &H0001)
			p.Value = 267
			cm.Parameters.Append p
		end if	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.open "spListImagesForProduct " & DisplayedID,cn,adOpenForwardOnly
		do while not rs.EOF	
			strSKUBGColor = "white"
			if trim(strlastRegion) <> trim(rs("ID")) then
				if blnFirst then
					blnFirst = false
					strLastRegion = rs("ID")
				else
					'Output row 
				'	strSKUBGColor = "cornsilk"
					blnFoundImages = false
					for i = lbound(PriorityArray) to ubound(PriorityArray)
						if trim(PriorityArray(i)) <> "" then
							blnFoundImages = true
							exit for
						end if
					next
					if blnFoundImages then
						strRegions = strRegions & "<TR>"
						for i = lbound(PriorityArray) to ubound(PriorityArray)
							if trim(PriorityArray(i)) <> "" then
								Select case trim(PriorityArray(i))
								case "1"
									ImageCountP1Array(i) = ImageCountP1Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "2"
									ImageCountP2Array(i) = ImageCountP2Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "3"
									ImageCountP3Array(i) = ImageCountP3Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "4"
									ImageCountP4Array(i) = ImageCountP4Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "5"
									ImageCountP5Array(i) = ImageCountP5Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								end select							
							else
								strSKUBGColor = "white"			
							end if
							if request("Display") <> "SKU" then					
								strRegions = strRegions & "<TD nowrap BGCOLOR=" & strSKUBGColor & " align=center><font size=1 face=verdana>"&  OutArray(i) & "</font></td>"
							else
								strRegions = strRegions & "<TD nowrap BGCOLOR=" & strSKUBGColor & " align=center><font size=1 face=verdana>"	& "SKU" & "</font></td>"						
							end if
						next
						strRegions = strRegions & strRegionDef & "</TR>"
					end if		
									
					'reset row array
					for i = lbound(OutArray) to ubound(OutArray)
						OutArray(i) = "&nbsp;"
						PriorityArray(i) = ""
					next
					
					'Remember Last Region
					strLastRegion = rs("ID")
					strTemp = ""
					intNativeCount = 0
					strRegionDef = ""
				end if
			end if

			if rs("Geo") & "" <> strLastGeo then
				strRegions = strRegions & "<TR bgcolor=gainsboro><TD colspan=" & ImageDefCount + 10 & "><font size=2 face=verdana color=black><b>" & rs("Geo") & "</b></font></TD></TR>"
				strLastGeo = rs("Geo") & "" 
			end if
			
			for i = lbound(ImageArray) to ubound(ImageArray)
				if trim(ImageArray(i)) = trim(rs("ImageDefID")) then
					if trim(rs("Priority") & "" ) = "" then
						OutArray(i) = "&nbsp;"
						PriorityArray(i) = ""
					else
					strImagesDisplayed=strImagesDisplayed & "," & rs("ImageID")

						if isnumeric(rs("Priority")) then
							if request("FileType") = "" then
								OutArray(i) = "<a target=""_blank"" href=""../Image/DeliverableMatrix.asp?ID=" & rs("ImageID") & "&PINTest=" & request("PINTest") & """>" & rs("Priority") & "</a>"
							else
								OutArray(i) = rs("Priority") & ""
							end if
							intNativeCount = intNativeCount + 1
						else
							OutArray(i) = replace(rs("Priority")," " ,"") & "&nbsp;Image "						
							if instr(strTemp,replace(rs("Priority")," " ,"") )=0 then
								set rs2 = server.CreateObject("ADODB.recordset")
							
								rs2.Open "spGetImageLanguage4Sub " & rs("ImageDefID") & ",'" & replace(rs("Priority")," " ,"")  & "'" ,cn,adOpenForwardOnly
								if rs2.EOF and rs2.BOF then
									strTemp = strTemp & "<BR>(" & replace(rs("Priority")," " ,"") & ")"
								else
									strLangList = "<u>" & rs2("OSLanguage") & "</u>"
									if rs2("OtherLanguage") & "" <> "" then
										strLangList = strLangList & "," & rs2("OtherLanguage")
									end if
									strTemp = strTemp & "<BR>" & replace(rs("Priority")," " ,"") & " (" & strLangList & ")"
								end if
								rs2.Close
								set rs2=nothing
							end if
						end if
						PriorityArray(i) = rs("Priority")
					end if
					exit for
				end if
			next
		'	if strRegionDef = "" then
				strRegionDef = "<TD><font face=verdana size=1>" & rs("Name") & "&nbsp;</font></td><TD><font face=verdana size=1>" & rs("OptionConfig") & "&nbsp;</font></td><TD nowrap><font face=verdana size=1>" & rs("Dash")  & "</font></td>"
				strLangList = "<font face=verdana size=1><u>" & rs("OSLanguage") & "</u>"
				if trim(rs("OtherLanguage") & "") <> "" then
					strLangList = strLangList & "," & trim(rs("OtherLanguage") & "")
				end if
				if trim(strTemp) = "" then
					strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("Dash")," ","") & "&nbsp;(" & strLangList & ")" & "</font></td>"
				else
					if intNativeCount > 0 then
						strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("Dash")," ","") & " (" & strLangList & ")" & strTemp & "</font></td>"
					else
						strRegionDef = strRegionDef & "<TD><font face=verdana size=1>"  & mid(strTemp,5) & "</font></td>"
					end if
				end if
				if rs("KWL") & "" <> "" then
					strKWL = rs("KWL") &  ""
				else
					strKWL = "&nbsp;"
				end if
				'strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & rs("CountryCode") & "</font></td><TD " & strKWL & "><font face=verdana size=1>" & rs("Keyboard") & "</font></td><TD><font face=verdana size=1>" & strKWL & "</font></td><TD><font face=verdana size=1>" & rs("PowerCord") & "</font></td>"
				strLangSub = ""
		'	end if
			rs.MoveNext
		loop
		rs.Close

		strRegions = strRegions & "<TR>"
		for i = lbound(PriorityArray) to ubound(PriorityArray)
			if trim(PriorityArray(i)) <> "" then
				Select case trim(PriorityArray(i))
				case "1"
					ImageCountP1Array(i) = ImageCountP1Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "2"
					ImageCountP2Array(i) = ImageCountP2Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "3"
					ImageCountP3Array(i) = ImageCountP3Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "4"
					ImageCountP4Array(i) = ImageCountP4Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "5"
					ImageCountP5Array(i) = ImageCountP5Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				end select							
			end if
			strRegions = strRegions & "<TD nowrap align=center><font size=1 face=verdana>"	&  OutArray(i) & "</font></td>"
		next
		strRegions = strRegions & strRegionDef & "</TR>"

%>
	<table ID=ScopeTable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<!--<tr  bgcolor=Gainsboro> <TD  align=center><font bgcolor=black size=2><b>Scope</b></font></TD></TR>-->
<%
	if strFilterList <> "" and request("FileType") = "" then
		strCountNote = strCountNote & "<BR><BR><font face=verdana size=1>" & "<a href=""../Image/DeliverableMatrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&ImageFilter=" & strImagesDisplayed & "&FilterName=" & strFilterList & """>Show Deliverable Matrix for selected images</a></font>"
	end if
	strModelRow = strModelRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Model </b></font></td><td bgcolor=white rowspan=9 colspan=7><font size=1 face=verdana><b>Total Image Count:&nbsp;<label ID=lblImageCount>" & TotalImageCount & "</label></b></font>" & strCountNote & strShowOption & "</td></tr>"
	strOSRow = strOSRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>OS </b></font></td></tr>"
	strAppsRow = strAppsRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Apps&nbsp;Bundle&nbsp;</b></font></td></tr>"
	strTypeRow = strTypeRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>BTO/CTO </b></font></td></tr>"
	strStatusRow = strStatusRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Status </b></font></td></tr>"
	strRTMDateRow = strRTMDateRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>RTM&nbsp;Date </b></font></td></tr>"
	strSKURow = strSKURow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Image&nbsp;# </b></font></td></tr>"
	strCommentsRow = strCommentsRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Comments </b></font></td></tr>"
	strModifiedRow = strModifiedRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Modified </b></font></td></tr>"
	strHideRow = strHideRow & "<td align=center><font size=1 face=verdana><a href="""">Show All</a></font></td><td colspan=8></td></tr>"

%>

		<%=strModelRow%>
		<%=strOSRow%>
		<%=strAppsRow%>
		<%=strTypeRow%>
		<%=strStatusRow%>
		<%=strRTMDateRow%>
		<%=strCommentsRow%>
		<%=strSKURow%>
		<%=strModifiedRow%>
		<%
			strTotalRows = "<TR>"
			strTierHeader = "<TR>"
			for i = lbound(ImageCountArray) to ubound(ImageCountArray)
				strTotalRows = strTotalRows & "<TD style=""BORDER-TOP: black solid"" align=center><font size=1 face=verdana>" & ImageCountArray(i) & "</font></td>"
				strTierHeader = strTierHeader & "<TD style=""BORDER-TOP: black solid"" align=center><font size=1 face=verdana><b>Tier</b></font></td>"
			next
			'Response.Write strTotalRows & "<TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Localization</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>OS Lang</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Dash</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Option Config</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Country Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Keyboard</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Power Cords</b></font></TD></TR>"
			Response.Write strTierHeader & "<TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Localization</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>HP<BR>Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Dash<BR>Code</b></font></TD><TD width=80 style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Images</b></font></TD></TR>"'<TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Country<BR>Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Keyboard</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>KWL</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Power<BR>Cords</b></font></TD></TR>"
			
		%>
		<%=strRegions%>
		<%=strTotalRows & "<TD colspan=8 style=""BORDER-TOP: black solid""><font face=verdana size=1>Tier x</font></td></tr>"%>
		<%
			Response.Write "<TR>"
			for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
				Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP1Array(i) & "</font></td>"
			next
			Response.Write "<TD colspan=8><font face=verdana size=1>Tier 1</font></td></tr>"

			Response.Write "<TR>"
			for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
				Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP2Array(i) & "</font></td>"
			next
			Response.Write "<TD colspan=8><font face=verdana size=1>Tier 2</font></td></tr>"

			Response.Write "<TR>"
			for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
				Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP3Array(i) & "</font></td>"
			next
			Response.Write "<TD colspan=8><font face=verdana size=1>Tier 3</font></td></tr>"

			Response.Write "<TR>"
			for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
				Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP4Array(i) & "</font></td>"
			next
			Response.Write "<TD colspan=8><font face=verdana size=1>Tier 4</font></td></tr>"

			Response.Write "<TR>"
			for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
				Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP5Array(i) & "</font></td>"
			next
			Response.Write "<TD colspan=8><font face=verdana size=1>Tier 5</font></td></tr>"
		%>
	</TABLE>


<%

	else
		Response.Write "<font size=2 face=verdana><b><center>No Images Defined for this Product</center></b></font>"
	end if

end if

cn.Close
set rs = nothing
set cn = nothing
if strImagesDisplayed <> "" then
	strImagesDisplayed = mid(strImagesDisplayed,2)
end if

%>

<INPUT type="hidden" id=txtMaxUpdated name=txtMaxUpdated value="<%=MaxUpdated%>">
<INPUT type="hidden" id=txtImageCount name=txtImageCount value="<%=TotalImageCount%>">
</BODY>
</HTML>

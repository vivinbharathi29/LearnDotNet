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
       	Response.Redirect "../../NoAccess.asp?Level=0"
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
			
			Response.Redirect "../../NoAccess.asp?Level=0"
		end if
	end if



	if blnOK then 
	
	
	dim strFilterList
	
	strFilterList = ""
	
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

	 if request("Comments") <> "" then
		strFilterList = strFilterList & "," & request("Comments") 
	 end if

	 if request("SKUNumber") <> "" then
		strFilterList = strFilterList & "," & request("SKUNumber") 
	 end if

	 if request("Modified") <> "" then
		strFilterList = strFilterList & ",Modified:" & request("Modified") 
	 end if

    if request("OSReleaseName") <> "" then
		strFilterList = strFilterList & ",OSReleaseName:" & request("OSReleaseName") 
	 end if
	
	strCountNote = ""
	if strFilterList <> "" then
		strFilterList = mid(strFilterList,2)
		strCountNote = "<br><BR><font color=red size=1 face=verdana>Not All Images Displayed. "
		if request("FileType") = "" then
			strCountNote = strCountNote & "<a href=""Localization.asp?ProdID=" & DisplayedID & """>Show All</a>"
		end if
		strCountNote = strCountNote & "</font>"	
	end if
	

%>

<center><font face=verdana size=3><b> <%=strProductname & " Image Localization Matrix"%>
<%if request("ImageID") <> "" then%>
	&nbsp;(Single Image Definition)
<%end if%>


</b><BR><BR></font></center>
<%if strFilterList <> "" then%>
<font size=2 face=verdana><center>Images Displayed: <%=strFilterList%></center><BR></font>
<%end if%>
<font size=2 face=verdana><center><label ID=lblModDate><%=formatdatetime(now,vbshortdate)%></label></center><BR></font>
<%if request("FileType") = "" then%>
<table width=100% border=0><tr><td align=right><font size=1 face=verdana>	Export: <a href="javascript: Export(1);">Excel</a> | <a href="javascript: Export(2);">Word</a></td></tr></table>
<%end if%>
<%

	Dim PreRow
	Dim MidRow
	Dim PostRow
	dim strRow
	dim strModelRow
	dim strOSRow
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
	dim ImageCountP0Array
	dim ImageCountP1Array
	dim ImageCountP2Array
	dim ImageCountP3Array
	dim ImageCountP4Array
	dim ImageCountP5Array
	dim ImageCountP6Array
	dim ImageCountP7Array
	dim ImageCountP8Array
	dim ImageCountP9Array
	dim ImageCountP10Array
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
	dim Tier0Found
	dim Tier1Found
	dim Tier2Found
	dim Tier3Found
	dim Tier4Found
	dim Tier5Found
	dim Tier6Found
	dim Tier7Found
	dim Tier8Found
	dim Tier9Found
	dim Tier10Found
    dim strOSReleaseNameRow
	
	Tier0Found = false
	Tier1Found = false
	Tier2Found = false
	Tier3Found = false
	Tier4Found = false
	Tier5Found = false
	Tier6Found = false
	Tier7Found = false
	Tier8Found = false
	Tier9Found = false
	Tier10Found = false


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
	if request("SKUNumber") = "" then
		strSKURow = "<TR BGCOLOR=white>"
	else
		strSKURow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
	if request("Modified") = "" then
		strModifiedRow = "<TR BGCOLOR=white>"
	else
		strModifiedRow = "<TR BGCOLOR=" & strFilterColor & ">"
	end if
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
	if request("Comments") = "" then
  		strCommentsRow = "<TR BGCOLOR=white>"
	else
  		strCommentsRow = "<TR BGCOLOR=" & strFilterColor & ">"
  	end if
    if request("OSReleaseName") = "" then
  		strOSReleaseNameRow = "<TR BGCOLOR=white>"
	else
  		strOSReleaseNameRow = "<TR BGCOLOR=" & strFilterColor & ">"
  	end if
  	strHideRow= "<TR>"
	blnFound = false

	strShowOption = ""	
	TotalImageCount = 0
	ImageDefCount = 0
	strImageIDList = ""
	
	if request("ImageID") <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetImageDefinitionBrandsFusion"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ImageID")
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	else
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "[spListImageDefinitions4ImageMatrix]"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = DisplayedID
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	end if
	
	do while not rs.EOF
		blnInclude = true
        if rs("ImageTypeID") = 2 then
            blnInclude = false
        end if

		if request("Brand") <> "" then
			if lcase(trim(rs("Brand"))) <> lcase(trim(request("Brand"))) and lcase(trim(rs("Brand"))) <> "all supported brands" then
				blnInclude = false
			end if
		end if
	
		if request("OpSys") <> "" then
			if lcase(trim(rs("OS"))) <> lcase(trim(request("OpSys"))) then
				blnInclude = false
			end if
		end if

		if request("SKUNumber") <> "" then
			if lcase(trim(replace(rs("SKUNumber"),"#",""))) <> lcase(trim(request("SKUNumber"))) then
				blnInclude = false
			end if
		end if

   		if request("Modified") <> "" then
			if lcase(trim(formatdatetime(rs("Modified"),vbshortdate))) <> lcase(trim(request("Modified"))) then
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

		if request("Comments") <> "" then
			if lcase(trim(rs("Comments") & "")) <> lcase(trim(request("Comments"))) then
				if not(lcase(trim(request("Comments"))) = "none" and lcase(trim(rs("Comments") & "")) = "") then
					blnInclude = false
				end if
			end if
		end if

        if request("OSReleaseName") <> "" then
			if lcase(trim(rs("OSReleaseName"))) <> lcase(trim(request("OSReleaseName"))) then
				blnInclude = false
			end if
		end if
	
	
		if blnInclude then

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
			
'			strSKURow = strSKURow & "<TD nowrap><font size=1 face=verdana>" & rs("SKUNumber") & "&nbsp;</font></td>"

   			if request("SKUNumber") = "" then
				if request("FileType") = "" then
					strSKURow = strSKURow & "<TD nowrap><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&SKUNumber=" & replace(rs("SKUNumber"),"#","") & "')"">"  & rs("SKUNumber") & "&nbsp;</a></font></td>"
				else
					strSKURow = strSKURow & "<TD nowrap><font size=1 face=verdana>"  & rs("SKUNumber") & "</font></td>"
				end if
			else
				strSKURow = strSKURow & "<TD nowrap><font size=1 face=verdana>" & rs("SKUNumber") & "</font></td>"
			end if


			if request("Comments") = "" and request("FileType") = "" then
				if request("Comments") = "" then
					strCommentsRow = strCommentsRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Comments=" & rs("Comments") & "')"">" & rs("Comments") & "</a>&nbsp;</font></td>"
				else
					strCommentsRow = strCommentsRow & "<TD><font size=1 face=verdana>" & rs("Comments") & "</font></td>"
				end if
			else
				strCommentsRow = strCommentsRow & "<TD><font size=1 face=verdana>" & rs("Comments") & "</font></td>"
			end if
            
		'	strModifiedRow = strModifiedRow & "<TD align=left><font size=1 face=verdana>" & formatdatetime(rs("Modified"),vbshortdate) & "</font></td>"

			if request("Modified") = "" then
				if request("FileType") = "" then
					if trim(rs("Modified")& "") = "" then
						strModifiedRow = strModifiedRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Modified=none')"">"  & formatdatetime(rs("Modified"),vbshortdate) & "&nbsp;</a></font></td>"
					else
						strModifiedRow = strModifiedRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&Modified=" & formatdatetime(rs("Modified"),vbshortdate) & "')"">"  & formatdatetime(rs("Modified"),vbshortdate) & "&nbsp;</a></font></td>"
					end if
				else
					strModifiedRow = strModifiedRow & "<TD align=left><font size=1 face=verdana>"  & formatdatetime(rs("Modified"),vbshortdate) & "</font></td>"
				end if
			else
				strModifiedRow = strModifiedRow & "<TD align=left><font size=1 face=verdana>" & formatdatetime(rs("Modified"),vbshortdate) & "</font></td>"
			end if
            
            if request("OSReleaseName") = "" and request("FileType") = "" and rs("OSReleaseName")<>""  then
                strOSReleaseNameRow = strOSReleaseNameRow & "<TD><font size=1 face=verdana><a class=""Filter"" href=""javascript: FilterMatrix('&OSReleaseName=" & rs("OSReleaseName") & "')"">" & rs("OSReleaseName") & "&nbsp;</a></font></td>"
			else
				strOSReleaseNameRow = strOSReleaseNameRow & "<TD><font size=1 face=verdana>" & rs("OSReleaseName") & "&nbsp;</font></td>"
			end if


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
	Redim ImageCountP0Array(Ubound(ImageArray))
	Redim ImageCountP1Array(Ubound(ImageArray))
	Redim ImageCountP2Array(Ubound(ImageArray))
	Redim ImageCountP3Array(Ubound(ImageArray))
	Redim ImageCountP4Array(Ubound(ImageArray))
	Redim ImageCountP5Array(Ubound(ImageArray))
	Redim ImageCountP6Array(Ubound(ImageArray))
	Redim ImageCountP7Array(Ubound(ImageArray))
	Redim ImageCountP8Array(Ubound(ImageArray))
	Redim ImageCountP9Array(Ubound(ImageArray))
	Redim ImageCountP10Array(Ubound(ImageArray))

	
	
	if blnFound then
		strLastGeo = ""
		strLastRegion = ""
		for i = lbound(OutArray) to ubound(OutArray)
			OutArray(i) = "&nbsp;"
			PriorityArray(i) = ""
			ImageCountP0Array(i) = 0
			ImageCountP1Array(i) = 0
			ImageCountP2Array(i) = 0
			ImageCountP3Array(i) = 0
			ImageCountP4Array(i) = 0
			ImageCountP5Array(i) = 0
			ImageCountP6Array(i) = 0
			ImageCountP7Array(i) = 0
			ImageCountP8Array(i) = 0
			ImageCountP9Array(i) = 0
			ImageCountP10Array(i) = 0
			ImageCountArray(i) = 0
		next
		
		blnFirst = true
		
		dim strSKUBGColor
		
		strLangSub = ""
		
		if request("ImageID") <> "" then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spListImagesForDefinition"
			
	
			Set p = cm.CreateParameter("@DefID", 3, &H0001)
			p.Value = clng(request("ImageID"))
			cm.Parameters.Append p
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

		else
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spListImagesForProductFusion"
			
	
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
		end if
	
		do while not rs.EOF	
            if rs("Published") then
           	strSKUBGColor = "white"
    		if trim(strlastRegion) <> trim(rs("RegionID")) then
				if blnFirst then
					blnFirst = false
					strLastRegion = rs("RegionID")
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
								case "0"
									ImageCountP0Array(i) = ImageCountP0Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier0Found = true
								case "1"
									ImageCountP1Array(i) = ImageCountP1Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier1Found = true
								case "2"
									ImageCountP2Array(i) = ImageCountP2Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier2Found = true
								case "3"
									ImageCountP3Array(i) = ImageCountP3Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier3Found = true
								case "4"
									ImageCountP4Array(i) = ImageCountP4Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier4Found = true
								case "5"
									ImageCountP5Array(i) = ImageCountP5Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier5Found = true
								case "6"
									ImageCountP6Array(i) = ImageCountP6Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier6Found = true
								case "7"
									ImageCountP7Array(i) = ImageCountP7Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier7Found = true
								case "8"
									ImageCountP8Array(i) = ImageCountP8Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier8Found = true
								case "9"
									ImageCountP9Array(i) = ImageCountP9Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier9Found = true
								case "10"
									ImageCountP10Array(i) = ImageCountP10Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
									Tier10Found = true
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
					strLastRegion = rs("RegionID")
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
				if trim(ImageArray(i)) = trim(rs("DefinitionID")) then
					if trim(rs("Priority") & "" ) = "" then
						OutArray(i) = "&nbsp;"
						PriorityArray(i) = ""
					else
					strImagesDisplayed=strImagesDisplayed & "," & rs("ID")

						if isnumeric(rs("Priority")) then
							if trim(rs("Priority")) = "0" then
								OutArray(i) = "Ref. Only"
							elseif request("FileType") = "" then
								OutArray(i) = "X" '"<a target=""_blank"" href=""DeliverableMatrix.asp?ID=" & rs("ID") & "&PINTest=" & request("PINTest") & """>X</a>"  
							else
								OutArray(i) = "X" 'rs("Dash") 
							end if
							intNativeCount = intNativeCount + 1
						else
							set rs2 = server.CreateObject("ADODB.recordset")
							rs2.Open "spGetImageProperties4Sub "  & rs("DefinitionID") & ",'" & replace(rs("Priority")," " ,"")  & "'" ,cn,adOpenForwardOnly
							if rs2.eof and rs2.bof then
								OutArray(i) = rs("Priority") & ""
							elseif request("FileType") <> "" then
								OutArray(i) = rs("Priority") & " (Tier&nbsp;" & replace(rs2("Priority")," " ,"") & ")"					
							else
								OutArray(i) = "<a target=""_blank"" href=""DeliverableMatrix.asp?ID=" & rs2("ID") & "&PINTest=" & request("PINTest") & """>" & rs("Priority") & "</a> (Tier&nbsp;" & trim(rs2("Priority") & "")& ")"					
							end if
							rs2.Close
							
							if instr(strTemp,replace(rs("Priority")," " ,"") )=0 then
							
								rs2.Open "spGetImageLanguage4Sub " & rs("DefinitionID") & ",'" & replace(rs("Priority")," " ,"")  & "'" ,cn,adOpenForwardOnly
								if rs2.EOF and rs2.BOF then
									strTemp = strTemp & "<BR>(" & replace(rs("Priority")," " ,"") & ")"
								else
									strLangList = "<u>" & rs2("OSLanguage") & "</u>"
									if rs2("OtherLanguage") & "" <> "" then
										strLangList = strLangList & "," & rs2("OtherLanguage")
									end if
									strTemp = strTemp & "<BR>" & replace(rs("Priority")," " ,"")'& " (" & strLangList & ")"
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
				strRegionDef = "<TD><font face=verdana size=1>" & rs("RegionName") & "&nbsp;</font></td><TD><font face=verdana size=1>" & rs("OptionConfig") & "&nbsp;</font></td><TD><font face=verdana size=1>" & rs("GMCode") & "&nbsp;</font></td><TD><font face=verdana size=1>" & rs("CountryCode") & "</font></td>"
				strLangList = "<font face=verdana size=1><u>" & rs("OSLanguage") & "</u>"
				if trim(rs("OtherLanguage") & "") <> "" then
					strLangList = strLangList & "," & trim(rs("OtherLanguage") & "")
				end if
				'if trim(strTemp) = "" then
				'	strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("OptionConfig")," ","") & "</font></td>"
				'else
				'	if intNativeCount > 0 then
				'		strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("OptionConfig")," ","") & strTemp & "</font></td>"
				'	else
				'		strRegionDef = strRegionDef & "<TD><font face=verdana size=1>"  & mid(strTemp,5) & "</font></td>"
				'	end if
				'end if
				'Product Dash
				'strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("OptionConfig")," ","") & "</font></td>"
				strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & replace(rs("Dash")," ","") & "</font></td>"

				strRegionDef = strRegionDef & "<TD><font face=verdana size=1>" & strLangList & "</font></td>"

				if rs("KWL") & "" <> "" then
					strKWL = rs("KWL") &  ""
				else
					strKWL = "&nbsp;"
				end if
				strRegionDef = strRegionDef & "<TD " & strKWL & "><font face=verdana size=1>" & rs("Keyboard") & "</font></td><TD><font face=verdana size=1>" & strKWL & "</font></td><TD><font face=verdana size=1>" & rs("PowerCord") & "</font></td><TD><font face=verdana size=1>" & rs("RestoreMedia") & "&nbsp;</font></td>"
				strLangSub = ""
                end if
			rs.MoveNext
		loop
		rs.Close

		strRegions = strRegions & "<TR>"
		for i = lbound(PriorityArray) to ubound(PriorityArray)
			if trim(PriorityArray(i)) <> "" then
				Select case trim(PriorityArray(i))
				case "0"
					ImageCountP0Array(i) = ImageCountP0Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier0Found = true
				case "1"
					ImageCountP1Array(i) = ImageCountP1Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier1Found = true
				case "2"
					ImageCountP2Array(i) = ImageCountP2Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier2Found = true
				case "3"
					ImageCountP3Array(i) = ImageCountP3Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier3Found = true
				case "4"
					ImageCountP4Array(i) = ImageCountP4Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier4Found = true
				case "5"
					ImageCountP5Array(i) = ImageCountP5Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier5Found = true
				case "6"
					ImageCountP6Array(i) = ImageCountP6Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier6Found = true
				case "7"
					ImageCountP7Array(i) = ImageCountP7Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier7Found = true
				case "8"
					ImageCountP8Array(i) = ImageCountP8Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier8Found = true
				case "9"
					ImageCountP9Array(i) = ImageCountP9Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier9Found = true
				case "10"
					ImageCountP10Array(i) = ImageCountP10Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
					Tier10Found = true
				end select							
			end if
			strRegions = strRegions & "<TD nowrap align=center><font size=1 face=verdana>"	&  OutArray(i) & "</font></td>"
		next
		strRegions = strRegions & strRegionDef & "</TR>"

%>
	<table ID=ScopeTable border=1 borderColor=black cellpadding=2 cellSpacing=0 width="100%">
		<!--<tr  bgcolor=Gainsboro> <TD  align=center><font bgcolor=black size=2><b>Scope</b></font></TD></TR>-->
<%
	'if strFilterList <> "" and request("FileType") = "" then
	'	strCountNote = strCountNote & "<BR><BR><font face=verdana size=1>" & "<a href=""DeliverableMatrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&ImageFilter=" & strImagesDisplayed & "&FilterName=" & strFilterList & """>Show Deliverable Matrix for selected images</a></font>"
	'end if
	strModelRow = strModelRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Model </b></font></td><td bgcolor=white rowspan=7 colspan=9><font size=1 face=verdana><b>Total Images:&nbsp;<label ID=lblImageCount>" & TotalImageCount & "</label></b></font>" & strCountNote & strShowOption & "</td></tr>"
	strOSRow = strOSRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>OS </b></font></td></tr>"
	strStatusRow = strStatusRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Status </b></font></td></tr>"
	strRTMDateRow = strRTMDateRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>RTM&nbsp;Date </b></font></td></tr>"
	strSKURow = strSKURow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Product&nbsp;Drop</b></font></td></tr>"
	strCommentsRow = strCommentsRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Comments </b></font></td></tr>"
	strModifiedRow = strModifiedRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Modified </b></font></td></tr>"
    strOSReleaseNameRow = strOSReleaseNameRow & "<td bgcolor=gainsboro><font size=1 face=verdana><b>Releases for Operating System </b></font></td></tr>"
	strHideRow = strHideRow & "<td align=center><font size=1 face=verdana><a href="""">Show All</a></font></td><td colspan=10></td></tr>"

%>

		<%=strModelRow%>
		<%=strOSRow%>
		<%=strStatusRow%>
		<%=strRTMDateRow%>
		<%=strCommentsRow%>
		<%=strSKURow%>
		<%=strModifiedRow%>
        <%=strOSReleaseNameRow%>
		<%
			strTotalRows = "<TR>"
			strTierHeader = "<TR>"
			for i = lbound(ImageCountArray) to ubound(ImageCountArray)
				strTotalRows = strTotalRows & "<TD style=""BORDER-TOP: black solid"" align=center><font size=1 face=verdana>" & ImageCountArray(i) & "</font></td>"
				strTierHeader = strTierHeader & "<TD style=""BORDER-TOP: black solid"" align=center><font size=1 face=verdana><b>Images</b></font></td>"
			next
			'Response.Write strTotalRows & "<TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Localization</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>OS Lang</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Dash</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Option Config</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Country Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Keyboard</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Power Cords</b></font></TD></TR>"
			Response.Write strTierHeader & "<TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Localization</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>HP<BR>Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>GM<br>Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Country<BR>Code</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Product<br>DASH</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Languages</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Keyboard</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>KWL</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Power<BR>Cords</b></font></TD><TD style=""BORDER-TOP: black solid""><font size=1 face=verdana><b>Restore<BR>Solution</b></font></TD></TR>"
			
		%>
		<%=strRegions%>
		<%=strTotalRows & "<TD colspan=10 style=""BORDER-TOP: black solid""><font face=verdana size=1>Image Count</font></td></tr>"%>
		<%
        if false then
			if Tier0Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP0Array) to ubound(ImageCountP0Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP0Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 0</font></td></tr>"
			end if
			if Tier1Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP1Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 1</font></td></tr>"
			end if
			
			if Tier2Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP2Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 2</font></td></tr>"
			end if
			
			if Tier3Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP3Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 3</font></td></tr>"
			end if
			
			if Tier4Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP4Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 4</font></td></tr>"
			end if
			
			if Tier5Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP5Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 5</font></td></tr>"
			end if
			
			if Tier6Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP6Array) to ubound(ImageCountP6Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP6Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 6</font></td></tr>"
			end if
			
			if Tier7Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP7Array) to ubound(ImageCountP7Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP7Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 7</font></td></tr>"
			end if
			
			if Tier8Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP8Array) to ubound(ImageCountP8Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP8Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 8</font></td></tr>"
			end if
			
			if Tier9Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP9Array) to ubound(ImageCountP9Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP9Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 9</font></td></tr>"
			end if
			
			if Tier10Found then
				Response.Write "<TR>"
				for i = lbound(ImageCountP10Array) to ubound(ImageCountP10Array)
					Response.Write "<TD align=center><font size=1 face=verdana>" & ImageCountP10Array(i) & "</font></td>"
				next
				Response.Write "<TD colspan=11><font face=verdana size=1>Tier 10</font></td></tr>"
			end if
        end if
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

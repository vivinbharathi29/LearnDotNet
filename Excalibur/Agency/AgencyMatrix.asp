<!-- #include file="../includes/DataWrapper.asp" -->	
<!-- #include file="../includes/common.asp" -->	

<%
	Const CHECK_MARK = "<font face=wingdings size=3><strong>ü</strong></font>"

	Function BuildRow(RowData, Country, RecordSet)
		
		'LANGUAGE=javascript onmouseover=""return Row_OnMouseOver()"" onmouseout=""return Row_OnMouseOut()"" onclick=""return Row_OnClick()""
		If right(Country, 1) = " " Then
			RowData = RowData &" <TR BgColor=LightSteelBlue>"
		Else
			RowData = RowData & "<TR>"
		End If
		
		AddCell RowData, Country
		
		Dim Value, AgencyType
		
		Do Until RecordSet.EOF
			
			AddStatusCell RowData, RecordSet
			
			RecordSet.MoveNext

		Loop

		RowData = RowData & "</TR>" & vbcrlf
		
		BuildRow = RowData
		
	End Function
	
	Function AddCell(InputString, Value)
			InputString = InputString & "<TD class=AgencyCell>" & iif(Len(Trim(Value))=0,"&nbsp;", Trim(Value)) & "</TD>"
	End Function

	Function AddStatusCell(InputString, RecordSet)
		Dim AgencyType, Value, LinkCode
		Dim CssClass
		
		AgencyType = Trim(RecordSet("agency_type") & "")
		LinkCode = ""
		Value = GetStatusValue(RecordSet)
		
		Select Case AgencyType
			Case "C"
				CssClass = GetStatusClass(RecordSet)
				'Value = GetStatusValue(RecordSet)
				'If UCase(Trim(RecordSet("status_cd") & "")) <> "NR" Then
					LinkCode = "LANGUAGE=javascript onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()' onclick='return Status_OnClick(" & RecordSet("agency_status_id") & ")'"
				'End If
			Case "I"
				CssClass = "Illegal"
				Value = Trim(RecordSet("agency_name") & "")
			Case "E"
				CssClass = "Embargo"
				Value = Trim(RecordSet("agency_name") & "")
			Case "N"
				CssClass = "NotNeeded"
				If RecordSet("pcid") & "" = "" Then
					CssClass = "NS_NotNeeded"
				End If
				'If InStr(lcase(Request.ServerVariables("SCRIPT_NAME")),"dmview") > 0 then
					Value = Trim(RecordSet("agency_name") & "")
				'Else
				'	Value = CHECK_MARK
				'End If
			Case Else
				CssClass = "Unknown"
				Value = Trim(RecordSet("agency_name") & "")
		End Select
		
		If UCASE(RecordSet("por_dcr")) = "DCR" Then
			Value = Value & "&nbsp;*"
		End If

		InputString = InputString & "<TD " & LinkCode & " align=middle class=" & CssClass & ">" & Value & "</TD>"
		
	End Function
	
	Function GetStatusClass(RecordSet)
		Dim StatusCd, ProjectedDate
		Dim StatusClass
		Dim LevStatusCd, LevProjectedDate
		
		StatusCd = Trim(RecordSet("status_cd") & "")
		ProjectedDate = Trim(RecordSet("projected_date") & "")
		LevStatusCd = Trim(RecordSet("leveraged_status_cd") & "")
		LevProjectedDate = Trim(RecordSet("leveraged_projected_date") & "")
	
		Select Case UCase(StatusCd)
			Case "O"
				If isDate(ProjectedDate) Then
					If DateDiff("d", Now(), ProjectedDate) > 0 Then
						StatusClass = "Cert_Open"
					Else
						StatusClass = "Cert_Late"
					End If
				Else
					StatusClass = "Cert_Late"
				End If
            Case "SU"
				If isDate(ProjectedDate) Then
					If DateDiff("d", Now(), ProjectedDate) > 0 Then
						StatusClass = "Cert_Open"
					Else
						StatusClass = "Cert_Late"
					End If
				Else
					StatusClass = "Cert_Late"
				End If
			Case "C"
				StatusClass = "Cert_Complete"
			Case "NS"
				StatusClass = "NotSupported"
			Case "L"
				If LevStatusCd = "C" Then
					StatusClass = "Cert_Complete"
				ElseIf LevStatusCd = "NS" Then
				    StatusClass = "NotSupported"
				ElseIf isDate(LevProjectedDate) Then
					If DateDiff("d", Now(), LevProjectedDate) > 0 Then
						StatusClass = "Cert_Open"
					Else
						StatusClass = "Cert_Late"
					End If
				Else
					StatusClass = "Cert_Late"
				End If
			Case Else
				StatusClass = "NotNeeded"
		End Select
		
		If Len(Trim(RecordSet("pcid") & "")) = 0 Then
			StatusClass = "NS_" & StatusClass
		End If

		GetStatusClass = StatusClass
	End Function
	
	Function GetStatusValue(RecordSet)
		Dim StatusCd, ProjectedDate, LevStatusCd, LevProjectedDate, SupportedCountryYN
		
		StatusCd = Trim(RecordSet("status_cd") & "")
		ProjectedDate = Trim(RecordSet("projected_date") & "")
		LevStatusCd = Trim(RecordSet("leveraged_status_cd") & "")
		LevProjectedDate = Trim(RecordSet("leveraged_projected_date") & "")
		SupportedCountryYN = Trim(RecordSet("supported_country_yn") & "")
		GetStatusValue = "&nbsp;"

		Select Case UCase(StatusCd)
			Case "O"
				If isDate(ProjectedDate) Then
					GetStatusValue = FormatDateTime(ProjectedDate, vbShortDate)
				Else
					GetStatusValue = "OPEN"
				End If
            Case "SU"
				If isDate(ProjectedDate) Then
					GetStatusValue = FormatDateTime(ProjectedDate, vbShortDate)
				Else
					GetStatusValue = "Supported"
				End If
			Case "C"
				If InStr(lcase(Request.ServerVariables("SCRIPT_NAME")),"dmview") > 0 then
					GetStatusValue = "Complete"
				Else
					GetStatusValue = CHECK_MARK
				End If
			Case "L"
				GetStatusValue = "Leveraged<BR>"
				If LevStatusCd = "C" Then
					If InStr(lcase(Request.ServerVariables("SCRIPT_NAME")),"dmview") > 0 then
						GetStatusValue = GetStatusValue & "Complete"
					Else
						GetStatusValue = GetStatusValue & CHECK_MARK
					End If
				Else
					If LevStatusCd = "NS" Then
						GetStatusValue = GetStatusValue & "Not Supported"
					ElseIf isDate(LevProjectedDate) Then
						GetStatusValue = GetStatusValue & FormatDateTime(LevProjectedDate, vbShortDate)
					Else
						GetStatusValue = GetStatusValue & "OPEN"
					End If
				End If

			Case "NS"
				If UCase(SupportedCountryYN) = "Y" Then
					GetStatusValue = "Product<br>"
				Else
					GetStatusValue = "Country<br>"
				End If
				GetStatusValue = GetStatusValue & "Not Supported"
			Case "NR"
				GetStatusValue = "Not Requested"
			Case Else
				GetStatusValue = "Invalid Status"
		End Select
		
		If Len(Trim(RecordSet("status_notes"))) > 0 Then
			GetStatusValue = GetStatusValue + "<BR>&lt;Note&gt;"
		End If

	End Function

	Function DrawGroupingRow(Region, ColCount)
		Response.Write "<TR><TD class=""Region"" colspan=" & ColCount & ">" & Region & "</TD></TR>"
	End Function
	
	Function GetProductCountryList(ProductVersionID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs
		Dim i
		Dim asCountries()
		dim strCookie

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		strCookie = ""
		on error resume next
		strCookie = Request.Cookies("PMStatus")
		on error goto 0
		If strCookie = "All" Then
			Set cmd = dw.CreateCommandSP(cn, "spListCountries4ProductAll")
		Else
			Set cmd = dw.CreateCommandSP(cn, "spListCountries4Product")
		End If

		dw.CreateParameter cmd, "@prodID", adInteger, adParamInput, 8, ProductVersionID
		Set rs = dw.ExecuteCommandReturnRS(cmd)	

		i = 0
		
		Do Until rs.eof
			ReDim Preserve asCountries(i)
			If Trim(rs("prodid") & "") <> ProductVersionID Then
				asCountries(i) = join(Array(rs("countryid") & "", rs("country") & " ",rs("region") & ""), "|")
			Else
				asCountries(i) = join(Array(rs("countryid") & "", rs("country") & "",rs("region") & ""), "|")
			End If
			i = i + 1
			rs.movenext
			
		Loop
		
		GetProductCountryList = join(asCountries, "||")
		
		rs.close
		set rs = nothing
		set cn = nothing
		set dw = nothing
	
	End Function
	
	Function GetDeliverableCountryList(DeliverableRootID, RegionID, CountryID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs
		Dim i
		Dim asCountries()
		dim strCookie

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

		strCookie = ""
		on error resume next
		strCookie = Request.Cookies("PMStatus")
		on error goto 0
		
		If strCookie = "All" Or Trim(RegionID) <> "" Or Trim(CountryID) <> "" Then
			Set cmd = dw.CreateCommandSP(cn, "usp_ListCountries")
			dw.CreateParameter cmd, "@p_RegionID", adVarChar, adParamInput, 15, RegionID
			dw.CreateParameter cmd, "@p_CountryID", adInteger, adParamInput, 8, CountryID
		Else
			Set cmd = dw.CreateCommandSP(cn, "usp_ListDeliverableCountries")
			dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, DeliverableRootID
		End If
		
		Set rs = dw.ExecuteCommandReturnRS(cmd)	

		i = 0
		
		Do Until rs.eof
			ReDim Preserve asCountries(i)
			asCountries(i) = join(Array(rs("country_id") & "", rs("country_name") & "",rs("region") & ""), "|")
			
			i = i + 1
			rs.movenext
			
		Loop
		
		GetDeliverableCountryList = join(asCountries, "||")
		
		rs.close
		set rs = nothing
		set cn = nothing
		set dw = nothing
	
	End Function

	Function GetAgencyDeliverables(ProductVersionID, DeliverableRootID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs
		Dim i
		Dim asDeliverables()

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyDeliverables")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
		dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, DeliverableRootID
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		i = 0
		
		Do Until rs.eof
			ReDim Preserve asDeliverables(i)
			asDeliverables(i) = Join(Array(rs("deliverable_root_id") & "", rs("deliverable_name") & ""), "|")
			
			i = i + 1
			rs.movenext
			
		Loop
		
		GetAgencyDeliverables = Join(asDeliverables, "||")
		
		rs.close
		set rs = nothing
		set cn = nothing
		set dw = nothing
	
	End Function
	
	Function GetAgencyProducts(ProductVersionID, SubID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs
		Dim i
		Dim asProducts()

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyDeliverables")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, SubID
		dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, ProductVersionID
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		i = 0
		
		Do Until rs.eof
			ReDim Preserve asProducts(i)
			asProducts(i) = Join(Array(rs("product_version_id") & "", rs("product_name") & ""), "|")
			
			i = i + 1
			rs.movenext
			
		Loop
		
		GetAgencyProducts = Join(asProducts, "||")
		
		rs.close
		set rs = nothing
		set cn = nothing
		set dw = nothing
	
	End Function
	
	Function DrawHeaderRow(sDeliverables)
		Dim i
		Dim sOutput
		Dim asDeliverables
		Dim asDeliverable
		Dim sMasterUrl
		Dim sUrl
		
		sMasterUrl = Request.ServerVariables("SCRIPT_NAME") & "?ID=" & Request("ID") & "&Class=" & Request("Class")
		'Response.Write surl
		
		asDeliverables = split(sDeliverables, "||")

		If UBOUND(asDeliverables) = -1 Then
			sOutput = "<TR><TH Class=HeaderRow>There are no agency requirements recorded.</TH>"
		Else
			If Request("PDDExport") Then
				sOutput = "<TR><TH Class=HeaderRow aling=left>Country</TH>"
			Else
				sOutput = "<TR><TH Class=HeaderRow aling=left><a href=""" & sMasterUrl & """>Country (Remove Filters)</a></TH>"
			End If 
		End If
		
		For i = 0 to UBOUND(asDeliverables)
		    asDeliverable = split(asDeliverables(i), "|")
		    sUrl = sMasterUrl & "&SubID=" & asDeliverable(0)
			If Request("PDDExport") Then
				sOutput = sOutput & "<TH class=HeaderRow width=100>" & asDeliverable(1) & "</TH>"
			Else
				sOutput = sOutput & "<TH class=HeaderRow width=100><a href=""" & sUrl & """>" & asDeliverable(1) & "</a></TH>"
			End If
		Next
		
		sOutput = sOutput & "</TR>"
		
		Response.Write sOutput		
		
	End Function

	Function DrawPMViewMatrix(ProductVersionID, DeliverableRootID, RegionID, CountryID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs

		Dim asCountries
		Dim asCountry
		Dim sDeliverables
		Dim iCountryBound	

		'Response.Flush

		'
		' Get a list of Countries
		'
		If Trim(RegionID) <> "" Or Trim(CountryID) <> "" Then
			asCountries = split(GetDeliverableCountryList(DeliverableRootID, RegionID, CountryID), "||")
		Else
			asCountries = split(GetProductCountryList(ProductVersionID), "||")
		End If
		
		'
		' Get the list of the number of Agency Deliverables
		'
		sDeliverables = GetAgencyDeliverables(ProductVersionID, DeliverableRootID)
		
		DrawHeaderRow sDeliverables

		'
		' Loop through countries and get the status by country
		'
		Dim sRegion
		sRegion = ""
		
		If UBOUND(split(sDeliverables,"||")) >= 0 Then
			For iCountryBound = 0 to ubound(asCountries)
				asCountry = Split(asCountries(iCountryBound), "|")
				
				Set dw = New DataWrapper
				Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
				Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyStatus")
				dw.CreateParameter cmd, "@p_StatusID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
				dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, DeliverableRootID
				dw.CreateParameter cmd, "@p_DeliverableCategoryID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_MappingID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_StatusCd", adVarChar, adParamInput, 10, ""
				dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_AgencyID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_CountryID", adInteger, adParamInput, 8, asCountry(0)
				Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
				If sRegion <> asCountry(2) Then
					sRegion = asCountry(2)
					DrawGroupingRow Replace(Replace(sRegion, "<",""),">",""), UBOUND(Split(sDeliverables, "||")) + 2
				End If
				
				If Not rs.eof Then
					Response.Write BuildRow("", asCountry(1), rs)
				End If
				
				rs.Close
				
				Set rs = nothing
				Set cn = nothing
				Set cmd = nothing
				Set dw = nothing
			Next
		End If
	End Function

	Function DrawDMViewMatrix(DeliverableRootID, ProductVersionID, RegionID, CountryID)
		Dim cn
		Dim cmd
		Dim dw
		Dim rs

		Dim asCountries
		Dim asCountry
		Dim sProducts
		Dim iCountryBound
		
		'Response.Flush

		'
		' Get a list of Countries
		'
		asCountries = split(GetDeliverableCountryList(DeliverableRootID, RegionID, CountryID), "||")
		
		'
		' Get the list of the number of Agency Deliverables
		'
		sProducts = GetAgencyProducts(DeliverableRootID, ProductVersionID)
		
		DrawHeaderRow sProducts

		'
		' Loop through countries and get the status by country
		'
		Dim sRegion
		sRegion = ""
		
		If UBOUND(Split(sProducts, "||")) >= 0 Then
			For iCountryBound = 0 to ubound(asCountries)
				asCountry = Split(asCountries(iCountryBound), "|")
				
				Set dw = New DataWrapper
				Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
				Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyStatus")
				dw.CreateParameter cmd, "@p_StatusID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
				dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, DeliverableRootID
				dw.CreateParameter cmd, "@p_DeliverableCategoryID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_MappingID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_StatusCd", adVarChar, adParamInput, 10, ""
				dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_AgencyID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_CountryID", adInteger, adParamInput, 8, asCountry(0)
				Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
				If sRegion <> asCountry(2) Then
					sRegion = asCountry(2)
					DrawGroupingRow Replace(Replace(sRegion, "<",""),">",""), UBOUND(Split(sProducts, "||")) + 2
				End If
				
				If Not rs.eof Then
					Response.Write BuildRow("", asCountry(1), rs)
				End If
				
				rs.Close
				
				Set rs = nothing
				Set cn = nothing
				Set cmd = nothing
				Set dw = nothing
			Next
		End If
	End Function
%>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var OldBGColor;

function DateDiff( start, end, interval, rounding ) {

    var iOut = 0;
    
    // Create 2 error messages, 1 for each argument. 
    var startMsg = "Check the Start Date and End Date\n"
        startMsg += "must be a valid date format.\n\n"
        startMsg += "Please try again." ;
		
    var intervalMsg = "Sorry the dateAdd function only accepts\n"
        intervalMsg += "d, h, m OR s intervals.\n\n"
        intervalMsg += "Please try again." ;

    var bufferA = Date.parse( start ) ;
    var bufferB = Date.parse( end ) ;
    	
    // check that the start parameter is a valid Date. 
    if ( isNaN (bufferA) || isNaN (bufferB) ) {
        alert( startMsg ) ;
        return null ;
    }
	
    // check that an interval parameter was not numeric. 
    if ( interval.charAt == 'undefined' ) {
        // the user specified an incorrect interval, handle the error. 
        alert( intervalMsg ) ;
        return null ;
    }
    
    var number = bufferB-bufferA ;
    
    // what kind of add to do? 
    switch (interval.charAt(0))
    {
        case 'd': case 'D': 
            iOut = parseInt(number / 86400000) ;
            if(rounding) iOut += parseInt((number % 86400000)/43200001) ;
            break ;
        case 'h': case 'H':
            iOut = parseInt(number / 3600000 ) ;
            if(rounding) iOut += parseInt((number % 3600000)/1800001) ;
            break ;
        case 'm': case 'M':
            iOut = parseInt(number / 60000 ) ;
            if(rounding) iOut += parseInt((number % 60000)/30001) ;
            break ;
        case 's': case 'S':
            iOut = parseInt(number / 1000 ) ;
            if(rounding) iOut += parseInt((number % 1000)/501) ;
            break ;
        default:
        // If we get to here then the interval parameter
        // didn't meet the d,h,m,s criteria.  Handle
        // the error. 		
        alert(intervalMsg) ;
        return null ;
    }
    
    return iOut ;
}

function trim( varText )
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function Status_OnMouseOver() {

		OldBGColor = window.event.srcElement.style.backgroundColor;
		window.event.srcElement.style.backgroundColor="thistle";
		window.event.srcElement.style.cursor="hand";
}

function Status_OnMouseOut() {
		window.event.srcElement.style.backgroundColor=OldBGColor;
}

function Status_OnClick(StatusID) {
	var strID;
	
	strID = window.showModalDialog("./Agency/Agency.asp?ID=" + StatusID,"","dialogWidth:700px;dialogHeight:500px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
	{
		if (strID=='C')
		{
			window.event.srcElement.className = 'Cert_Complete';
			window.event.srcElement.innerText = 'COMPLETE';
		}
		else if(strID=='L')
		{
			window.event.srcElement.className = 'Cert_Complete';
			window.event.srcElement.innerText = 'LEVERAGED';
			//window.navigate(window.location.href);
//			window.location.reload();
		}
		else if(strID=='NS')
		{
			window.event.srcElement.className = 'NotSupported';
			window.event.srcElement.innerText = 'NOT SUPPORTED';
		}
		else if(strID=='NR')
		{
			window.event.srcElement.className = 'NotNeeded';
			window.event.srcElement.innerText = 'NOT REQUESTED';
		}
		else if(strID=='O')
		{
			window.event.srcElement.className = 'Cert_Late';
			window.event.srcElement.innerText = 'OPEN'
		}
		else if (strID == 'SU') {
		    window.event.srcElement.className = 'Cert_Late';
		    window.event.srcElement.innerText = 'SUPPORTED'
		}
		else
		{	
			var Now = new Date();
			var StartDate = new Date();
			if(strID != '') {
				if(!isNaN(Date.parse( strID ))) {
				StartDate = new Date(Date.parse(strID)) ;
		        }
			}
			if (DateDiff(Now, StartDate, 'd', true) <= 0) {
				window.event.srcElement.className = 'Cert_Late';
				}
			else {
				window.event.srcElement.className = 'Cert_Open';
				}
			window.event.srcElement.innerText = strID;
		}
	}
}

function Export(strID){
	if (txtCurrentFilter.value == "" &&  cboCategory.value==0 && txtProductList.value=="")
		window.open (window.location.href + "?FileType=" + strID);
	else	
		window.open (window.location.href + "&FileType=" + strID);
}

//-->
</SCRIPT>
<STYLE>
.HeaderRow
{
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
    BACKGROUND-COLOR: Beige;
}
.Illegal
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Embargo
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Unknown
{
    FONT-FAMILY: Verdana;
    COLOR: black;
    FONT-SIZE: xx-small;
    BACKGROUND-COLOR: gold;
}
.Cert_Late
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Cert_Open
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: springgreen;
}
.Cert_Complete
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    FONT-WEIGHT: bold;
    COLOR: DarkGreen;
}
.NotNeeded
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
}
.NotSupported
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Changed
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: yellow;
}
.AgencyCell
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
}
.Region
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: MediumAquamarine;
}
.NS_AgencyCell
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Late
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Open
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Complete
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: DarkGreen;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_NotNeeded
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_NotSupported
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
</STYLE>
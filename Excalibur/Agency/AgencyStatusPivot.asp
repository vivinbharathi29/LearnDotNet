<%@ Language=VBScript %>
<%Option Explicit%>
<%

Response.AddHeader "Pragma", "No-Cache"

if request("FileType")= 1 then
	Response.ContentType = "application/vnd.ms-excel"
elseif request("FileType")= 2 then
	Response.ContentType = "application/msword"
end if
%>
<!-- #include file="../includes/DataWrapper.asp" -->	
<!-- #include file="../includes/common.asp" -->	
<%
	Dim m_IsAdmin
	Dim m_IsSysAdmin

	Function Main()
		Dim dw
		Dim cn
		Dim cmd
		Dim rs
		Dim CurrentUser
		Dim CurrentUserID
		Dim CurrentDomain

		CurrentUser = lcase(Session("LoggedInUser"))

		if InStr(CurrentUser,"\") > 0 then
			CurrentDomain = Left(CurrentUser, instr(CurrentUser,"\") - 1)
			CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
		end if
		
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spGetUserInfo")
		dw.CreateParameter cmd, "@UserName", adVarChar, adParamInput, 80, CurrentUser
		dw.CreateParameter cmd, "@Domain", adVarChar, adParamInput, 30, CurrentDomain
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") & ""
		else
			CurrentUserID = 0
		end if
		rs.Close

		m_IsSysAdmin = dw.UserIsAdmin(cn, CurrentUserID)
		
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing
		
		Call DrawAgencyMatrix()
	
	End Function

	Function BuildRow(RowData, RecordSet)
		
		RowData = RowData &"<TR bgcolor=ivory>"
		
		Dim field
				
		For Each filed in rs.fields
		
			Select Case ucase(field.name)
				Case "country_id"
				Case "country_name"
					AddCell RowData, field.value
				Case Else
					AddStatusCell RowData, field.value
			End Select
		
		Next

		RowData = RowData & "</TR>" & vbcrlf
		
		BuildRow = RowData
		
	End Function
	
	Function AddCell(InputString, Value)
		InputString = InputString & "<TD>" & iif(Len(Trim(Value))=0,"&nbsp;", Value) & "</TD>"
	End Function

	Function AddStatusCell(InputString, DataString)
		Dim asData
		asData = Split(DataString, "|")
				
		Dim AgencyType, Value, LinkCode
		Dim CssClass
		
		AgencyType = asData(0)
		LinkCode = ""
		Value = Trim(asData(4) & "")
		
		Select Case AgencyType
			Case "C"
				CssClass = GetStatusClass(asData(2), asData(3))
				Value = GetStatusValue(asData(2), asData(3))
				LinkCode = "LANGUAGE=javascript onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()' onclick='return Status_OnClick(" & asData(1) & ")'"
			Case "I"
				CssClass = "Illegal"
			Case "E"
				CssClass = "Embargo"
			Case "N"
				CssClass = "NotNeeded"
			Case Else
				CssClass = "Unknown"
		End Select

		InputString = InputString & "<TD " & LinkCode & " align=middle class=" & CssClass & ">" & Value & "</TD>"
		
	End Function
	
	Function GetStatusClass(StatusCd, ProjectedDate)
		
		GetStatusClass = ""
	
		If (UCase(StatusCd) = "O" OR UCase(StatusCd) = "SU") And isDate(ProjectedDate) Then
			If DateDiff("d", Now(), ProjectedDate) < 0 Then
				GetStatusClass = "Certification_Late"
			'ElseIf DateDiff("d", Now(), ProjectedDate) <= 30 Then
			'	GetStatusClass = "Certification_Warn"
			Else
				GetStatusClass = "Certification_Open"
			End If
		Else
			GetStatusClass = "Certification_Late"
		End If
		
		If UCase(StatusCd) = "C" Or UCase(StatusCd) = "L"Then
			GetStatusClass = "Certification_Complete"
		End If
		
		If UCase(StatusCd) = "NS" Then
			GetStatusClass = "NotSupported"
		End If
		
		
	End Function
	
	Function GetStatusValue(StatusCd, ProjectedDate)
		GetStatusValue = "&nbsp;"

		If (UCase(StatusCd) = "O" OR UCase(StatusCd) = "SU")  And isDate(ProjectedDate) Then
			GetStatusValue = FormatDateTime(ProjectedDate, vbShortDate)
		Else
			GetStatusValue = "OPEN"
		End If
		
		If UCase(StatusCd) = "C" Or UCase(StatusCd) = "L" Then
			GetStatusValue = "Complete"
		End If
		
		If UCase(StatusCd) = "NS" Then
			GetStatusValue = "Not Supported"
		End If

	End Function

	Function DrawHeaderRow(rs)
		dim sOutput, field
				
		sOutput = "<TR bgcolor=Beige>"
		
		For Each field In rs.Fields
			If field.name <> "country_id" Then
				sOutput = sOutput & "<TH>" & field.name & "</TH>"
			End If
		Next
		
		sOutput = sOutput & "</TR>"
		
		Response.Write sOutput		
	End Function

	Function DrawAgencyMatrix()
		
		Dim cn
		Dim cmd
		Dim dw
		Dim rs

		Dim strRow
		Dim asCountries
		Dim asCountry
		Dim sDeliverables
		Dim iDeliverableCount	
		Dim iCountryBound	
		
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_AgencyStatusPivot")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("ID")
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof & rs.bof Then
			Response.Write "No Records Returned"
			Response.End
		End If

		DrawHeaderRow rs
		Response.End
				
		Do Until rs.eof
	
			If Not rs.eof Then
				strRow = BuildRow(strRow, rs)
			End If
			
			rs.movenext			
			
		Loop
		
		Response.Write strRow
	
		rs.Close
		
		Set rs = nothing
		Set cn = nothing
		Set cmd = nothing
		Set dw = nothing

	End Function
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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

	if (window.event.srcElement.className=="Status" || window.event.srcElement.className=="VersionPart" || window.event.srcElement.className=="Header" || window.event.srcElement.className=="Sub")
		{
		OldBGColor = window.event.srcElement.style.backgroundColor;
		window.event.srcElement.style.backgroundColor="thistle";
		window.event.srcElement.style.cursor="hand";
		}
	else if (window.event.srcElement.className!="Static")
		{
		window.event.srcElement.parentElement.style.color="red";
		window.event.srcElement.parentElement.style.cursor="hand";
		}
}

function Status_OnMouseOut() {
	if (window.event.srcElement.className=="Status"  || window.event.srcElement.className=="VersionPart" || window.event.srcElement.className=="Header" || window.event.srcElement.className=="Sub")
		window.event.srcElement.style.backgroundColor=OldBGColor;
	else if (window.event.srcElement.className!="Static" )
		window.event.srcElement.parentElement.style.color="black";
}

function Status_OnClick(StatusID) {
	var strID;
	
		strID = window.showModalDialog("Agency.asp?ID=" + StatusID,"","dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
		if (typeof(strID) != "undefined")
		{
			if (strID=='C')
			{
				window.event.srcElement.className = 'Certification_Complete';
				window.event.srcElement.innerText = 'Complete';
			}
			else if(strID=='L')
			{
				window.event.srcElement.className = 'Certification_Complete';
				window.event.srcElement.innerText = 'Leveraged';
			}
			else if(strID=='NS')
			{
				window.event.srcElement.className = 'NotSupported';
				window.event.srcElement.innerText = 'Not Supported';
			}
			else if(strID=='O')
			{
				window.event.srcElement.className = 'Certification_Late';
				window.event.srcElement.innerText = 'Open'
			}
			else if (strID == 'SU') {
			    window.event.srcElement.className = 'Certification_Late';
			    window.event.srcElement.innerText = 'Supported'
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
					window.event.srcElement.className = 'Certification_Late';
					}
				else {
					window.event.srcElement.className = 'Certification_Open';
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
<LINK rel="stylesheet" type="text/css" href="agency.css">
</HEAD>
<BODY>

<table width=100%><tr><td colspan=3>

<%if request("FileType") = "" then%>
</td><td align=right>
<table width=100% border=0><tr><td align=right><font size=1 face=verdana>	Export: <a href="javascript: Export(1);">Excel</a></td></tr></table>
<%end if%>
<Table  width=100% border=0 cellpadding=2 cellspacing=0 ><TR><TD ID=CategoryFilter align=right></TD></TR></table>
</td></tr></table>
<table width=100%  border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>
<%
Call Main
%>
</TR></table></DIV>
<BR><BR><BR><BR>
<font size=1>Generated <%=formatdatetime(now)%></font>
</BODY>
</HTML>

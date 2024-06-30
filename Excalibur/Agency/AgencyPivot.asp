<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/common.asp" -->

<%
Const CHECK_MARK = "<font face=wingdings size=3><strong>?/strong></font>"
AppRoot = Session("ApplicationRoot")
    
Dim isAgencyDataMaintainer : isAgencyDataMaintainer = false
Dim m_ProductID
Dim m_from_where : m_from_where = ""
Dim m_Platform : m_Platform = ""

If InStr(Request.ServerVariables("URL") & "", "pmview") > 0 Then
    m_from_where = "Certification"
ElseIf InStr(Request.ServerVariables("URL") & "", "dmview") > 0 Then
    m_from_where = "Agency"
End If

Function BuildRow(RecordSet, HeaderRecordSet)
    Dim rowData, strRow
    Dim statusID, agencyType, agencyName, statusCd, projectedDt, levStatusCd, levProjectedDt
    Dim supportedCountryYN, statusNote, porDcr, pcid, modifiedBy
    Dim field
    Dim tmpCell

    tmpCell = "N/A"

	If RecordSet.Fields(3).Value & "" = "" Then
		supportedCountryYN = "N"
	Else
		supportedCountryYN = "Y"
	End If
	
    If RecordSet.Fields(3).Value & "" = "" Then
        strRow = " <TR BgColor=LightSteelBlue>"
    Else
        strRow = "<TR>"
    End If

    'strRow = AddCell(strRow, RecordSet("Country"))

    For Each field In RecordSet.Fields

        If field.name = RecordSet.Fields(0).name Then
			'Do Nothing                                    
        ElseIf field.name = RecordSet.Fields(1).name Then
            strRow = strRow & AddCell(field.value)         
        ElseIf field.name = RecordSet.Fields(2).name Then
            'Do Nothing                                     
        ElseIf field.name = RecordSet.Fields(3).name Then
            'Do Nothing                                     
        ElseIf InStr(field.value & "","|") <= 0 Then
            If InStr(field.value & "","-") > 0 Then
                dim txt
                tmpCell = Trim(field.value)
                txt = Split(tmpCell,"-")
                Select case CInt(txt(0))
                    case 0
                        tmpCell = "Not Supported"
                    case 1
                        tmpCell = "Supported"
                    case 2
                        tmpCell = txt(1)
                End select
            Else 
                tmpCell = "N/A"
            End If

            If isAgencyDataMaintainer = True Then
                strRow = strRow & "<TD onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()' onclick='ShowReleaseStatus(" & RecordSet.Fields(0).Value & ",""" &_
                                                                                                                                                        RecordSet.Fields(1).Value & """," &_
                                                                                                                                                        field.name & ",""" &_ 
                                                                                                                                                        HeaderRecordSet.Fields(field.name).Value & """," &_ 
                                                                                                                                                        m_ProductID & ",""" &_ 
                                                                                                                                                        field.value & """, """ &_ 
                                                                                                                                                        m_Platform & """);'>" & tmpCell & "</TD>"
            Else 
                strRow = strRow & "<TD onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()'>" & tmpCell & "</TD>"
            End If	
		Else
            rowData = Split(field.value, "|")

            statusID = rowData(0)
            agencyType = rowData(1)
            agencyName = rowData(2)
            statusCd = rowData(3)
            projectedDt = rowData(4)
            levStatusCd = rowData(5)
            levProjectedDt = rowData(6)
            supportedCountryYN = rowData(7)
            statusNote = rowData(8)
            porDcr = rowData(9)
            pcid = rowData(10)
            modifiedBy = rowData(11)

            strRow = strRow & AddStatusCell(statusID, agencyType, agencyName, statusCd, projectedDt, levStatusCd, levProjectedDt, supportedCountryYN, statusNote, porDcr, pcid, modifiedBy)
        End If

    Next

    strRow = strRow & "</TR>" & vbcrlf

    BuildRow = strRow

End Function

Function AddCell(Value)
    AddCell = "<TD class=AgencyCell>" & iif(Len(Trim(Value))=0,"&nbsp;", Trim(Value)) & "</TD>"
End Function

Function AddStatusCell(StatusID, AgencyType, AgencyName, StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote, PorDcr, PCID, ModifiedBy)
    Dim Value, LinkCode
    Dim CssClass
    Dim StarMarked

    LinkCode = ""
    Value = GetStatusValue(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote)

    Select Case AgencyType
        Case "C", "P"
            CssClass = GetStatusClass(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, PCID)
            If isAgencyDataMaintainer = true Then
                LinkCode = "LANGUAGE=javascript onmouseover='return Status_OnMouseOver()' onmousedown='return Status_OnMouseDown()' onmouseout='return Status_OnMouseOut()' onclick='return Status_OnClick(" & StatusID & ")'"
            Else
                LinkCode = "LANGUAGE=javascript onmouseover='return Status_OnMouseOver()' onmouseout='return Status_OnMouseOut()'"
            End If
        Case "I"
            CssClass = "Illegal"
            Value = AgencyName
        Case "E"
            CssClass = "Embargo"
            Value = AgencyName
        Case "N"
            CssClass = "NotNeeded"
            If PCID = "" Then
                CssClass = "NS_NotNeeded"
            End If
            Value = AgencyName
        Case Else
            CssClass = "Unknown"
            Value = AgencyName
    End Select

    If UCASE(PorDcr) = "DCR" Then
        Value = Value & "&nbsp;*"
    End If

    If ModifiedBy <> "" Then
        StarMarked = "*"
    End If

    AddStatusCell = "<TD " & LinkCode & " StatusID=""" & StatusID & """ data-statusid=""" & StatusID & """align=middle class=" & CssClass & ">" & StarMarked & Value & "</TD>"
 
End Function

Function GetStatusClass(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, PCID)
    Dim StatusClass

    Select Case UCase(StatusCd)
        Case "O"
            If isDate(ProjectedDt) Then
                If DateDiff("d", Now(), ProjectedDt) > 0 Then
                    StatusClass = "Cert_Open"
                Else
                    StatusClass = "Cert_Late"
                End If
            Else
                StatusClass = "Cert_Late"
            End If
        Case "SU"
            If isDate(ProjectedDt) Then
                If DateDiff("d", Now(), ProjectedDt) > 0 Then
                    StatusClass = "Cert_Open"
                Else
                    StatusClass = "Cert_Late"
                End If
            Else
                StatusClass = "Cert_Late"
            End If
        Case "C", "P"
            StatusClass = "Cert_Complete"
        Case "NS"
            StatusClass = "NotSupported"
        Case "L"
            If LevStatusCd = "C" Then
                StatusClass = "Cert_Complete"
            ElseIf LevStatusCd = "NS" Then
                StatusClass = "NotSupported"
            ElseIf isDate(LevProjectedDt) Then
                If DateDiff("d", Now(), LevProjectedDt) > 0 Then
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

    If Len(Trim(PCID)) = 0 Then
        StatusClass = "NS_" & StatusClass
    End If

    GetStatusClass = StatusClass
End Function

Function GetStatusValue(StatusCd, ProjectedDt, LevStatusCd, LevProjectedDt, SupportedCountryYN, StatusNote)

    GetStatusValue = "&nbsp;"

    Select Case UCase(StatusCd)
        Case "O"
            If isDate(ProjectedDt) Then
                GetStatusValue = FormatDateTime(ProjectedDt, vbShortDate)
            Else
                GetStatusValue = "OPEN"
            End If
         Case "SU"
            If isDate(ProjectedDt) Then
                GetStatusValue = FormatDateTime(ProjectedDt, vbShortDate)
            Else
                GetStatusValue = "Supported"
            End If
        Case "C"
			GetStatusValue = "Certified"
        Case "P"
			GetStatusValue = "Partial"
        Case "L"
            GetStatusValue = "Leveraged<BR>"
            If LevStatusCd = "C" Then
				GetStatusValue = GetStatusValue & "Certified"
            Else
                If LevStatusCd = "NS" Then
                    GetStatusValue = GetStatusValue & "Not Supported"
                ElseIf LevStatusCd = "SU" Then
                    GetStatusValue = GetStatusValue & "Supported"
                ElseIf isDate(LevProjectedDt) Then
                    GetStatusValue = GetStatusValue & FormatDateTime(LevProjectedDt, vbShortDate)
                Else
                    GetStatusValue = GetStatusValue & "OPEN " & LevProjectedDt
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
        Case "NC"
            GetStatusValue = "No Cert Needed"
        Case Else
            GetStatusValue = "Invalid Status"
    End Select

    If Len(Trim(StatusNote)) > 0 Then
        GetStatusValue = GetStatusValue + "<BR>&lt;Note&gt;"
    End If

End Function

Function DrawGroupingRow(Region, ColCount)
    Response.Write "<TR><TD class=""Region"" colspan=" & ColCount & ">" & Region & "</TD></TR>"
End Function
	
Function DrawHeaderRow(RecordSet)
	Dim i
	Dim sOutput
	Dim asDeliverables
	Dim asDeliverable
	Dim sMasterUrl
	Dim sUrl
	Dim field
	
	sMasterUrl = Request.ServerVariables("SCRIPT_NAME") & "?ID=" & Request("ID") & "&Class=" & Request("Class")
	'Response.Write surl
	
'	Set RecordSet = Server.CreateObject("ADODB.RecordSet")
	
	Response.Write "<THEAD><TR class=""FrozenHeader"">"
	If RecordSet.Fields.Count = 0 Then
		Response.Write "<TH Class=HeaderRow>There are no agency requirements recorded.</TH>"
		Exit Function
	End If
	For Each field IN RecordSet.Fields
		If field = RecordSet.Fields(0) Then
			Response.Write "<TH Class=HeaderRow>" & field.value & "</TH>"
		Else
			Response.Write "<TH Class=HeaderRow width=100>" & field.value & "</TH>"
		End If
	Next
	Response.Write "</TR></THEAD>"

End Function

Function DrawPMViewMatrix(ProductVersionID, AgencyDataMaintainer)
	Dim cn
	Dim cmd
	Dim dw
	Dim rs
	dim strCookie

    isAgencyDataMaintainer = AgencyDataMaintainer
    m_ProductID = ProductVersionID

	strCookie = ""
	on error resume next
	strCookie = Request.Cookies("PMStatus")
	on error goto 0
	
	Dim showAll
	If strCookie = "All" Then
		showAll = True
	Else
		showAll = False
	End If

	Dim colCount
	Dim sRegion
	sRegion = ""

	'Response.Flush

	'
	' Get Data
	'
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "rpt_AgencyPMView")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
    dw.CreateParameter cmd, "@p_ModifiedBy", adVarChar, adParamInput, 15, m_from_where
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
    ' set header
	colCount = rs.Fields.Count
    Set rs_header = rs
	DrawHeaderRow(rs)
		
	If colCount = 0 Then
		Response.End
	End If

    ' set Regulatory start
    Set rs = rs.NextRecordset
    m_Platform = rs.Fields(2).value

    Response.Write "<TR><TD></TD><TD style='text-align:center;' colspan=" & rs("ReleaseCount") & ">System and Adapters</TD></TR>"
    IF AgencyDataMaintainer = True Then
        Response.Write "<TR>" &_
                           "<TD></TD>" &_ 
                           "<TD style='text-align:center;' colspan=" & rs("ReleaseCount") & ">" &_ 
                               "Regulatory Reference: <input type='text' id='RegulatoryText' value='" & rs("regulatorymodel") & "' maxlength='15' />&nbsp;" &_ 
                               "<input id='BtnSaveRegulatory' type='button' value='Save Regulatory' onclick='SaveRegulatory();' />" &_
                           "</TD>" &_ 
                       "</TR>"
    ELSE 
        Response.Write "<TR>" &_ 
                          "<TD></TD>" &_
                          "<TD style='text-align:center;' colspan=" & rs("ReleaseCount") & ">" &_
                          "Regulatory Reference: <label>" & rs("regulatorymodel") & "</label>" &_
                       "</TR>"
    END IF
    ' set Regulatory end
		
	Set rs = rs.NextRecordset
	
   ' set body
    'Response.Write "Rows Count: " & rs.fields.Count

	Do Until rs.EOF
		
		If sRegion <> rs("Region") Then
			sRegion = rs("Region")
			DrawGroupingRow Replace(Replace(sRegion, "<",""),">",""), colCount
		End If

		If showAll Then
			Response.Write BuildRow(rs, rs_header)
		ElseIf rs.Fields(3).Value & "" <> "" Then
			Response.Write BuildRow(rs, rs_header)
		End If
			
		rs.MoveNext
	Loop

	rs.Close
				
	Set rs = nothing
	Set cn = nothing
	Set cmd = nothing
	Set dw = nothing

End Function

Function DrawDMViewMatrix(DeliverableRootID, AgencyDataMaintainer)
	Dim cn
	Dim cmd
	Dim dw
	Dim rs
	dim strCookie
	Dim showAll

    isAgencyDataMaintainer = AgencyDataMaintainer

	strCookie = ""
	on error resume next
	strCookie = Request.Cookies("DMStatus")
	on error goto 0

	If strCookie = "All" Then
		showAll = True
	Else
		showAll = False
	End If

	Dim colCount
	Dim sRegion
	sRegion = ""

	'Response.Flush

	'
	' Get Data
	'
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "rpt_AgencyDMView")
	dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, DeliverableRootID
    dw.CreateParameter cmd, "@p_ModifiedBy", adVarChar, adParamInput, 15, m_from_where
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
	colCount = rs.Fields.Count
    Set rs_header = rs
	DrawHeaderRow(rs)
		
	If colCount = 0 Then
		Response.End
	End If
		
	Set rs = rs.NextRecordset
	rs.sort = "Region, Country"
	Do Until rs.EOF
		
		If sRegion <> rs("Region") Then
			sRegion = rs("Region")
			DrawGroupingRow Replace(Replace(sRegion, "<",""),">",""), colCount
		End If

		If showAll Then
			Response.Write BuildRow(rs,rs_header)
		ElseIf rs.Fields(3).Value & "" <> "" Then
			Response.Write BuildRow(rs,rs_header)
		End If
			
		rs.MoveNext
	Loop


		
	rs.Close
				
	Set rs = nothing
	Set cn = nothing
	Set cmd = nothing
	Set dw = nothing
	

End Function
%>
<script id="clientEventHandlersJS" language="javascript">
<!--
    var OldBGColor;

    function DateDiff(start, end, interval, rounding) {

        var iOut = 0;

        // Create 2 error messages, 1 for each argument. 
        var startMsg = "Check the Start Date and End Date\n"
        startMsg += "must be a valid date format.\n\n"
        startMsg += "Please try again.";

        var intervalMsg = "Sorry the dateAdd function only accepts\n"
        intervalMsg += "d, h, m OR s intervals.\n\n"
        intervalMsg += "Please try again.";

        var bufferA = Date.parse(start);
        var bufferB = Date.parse(end);

        // check that the start parameter is a valid Date. 
        if (isNaN(bufferA) || isNaN(bufferB)) {
            alert(startMsg);
            return null;
        }

        // check that an interval parameter was not numeric. 
        if (interval.charAt == 'undefined') {
            // the user specified an incorrect interval, handle the error. 
            alert(intervalMsg);
            return null;
        }

        var number = bufferB - bufferA;

        // what kind of add to do? 
        switch (interval.charAt(0)) {
            case 'd': case 'D':
                iOut = parseInt(number / 86400000);
                if (rounding) iOut += parseInt((number % 86400000) / 43200001);
                break;
            case 'h': case 'H':
                iOut = parseInt(number / 3600000);
                if (rounding) iOut += parseInt((number % 3600000) / 1800001);
                break;
            case 'm': case 'M':
                iOut = parseInt(number / 60000);
                if (rounding) iOut += parseInt((number % 60000) / 30001);
                break;
            case 's': case 'S':
                iOut = parseInt(number / 1000);
                if (rounding) iOut += parseInt((number % 1000) / 501);
                break;
            default:
                // If we get to here then the interval parameter
                // didn't meet the d,h,m,s criteria.  Handle
                // the error. 		
                alert(intervalMsg);
                return null;
        }

        return iOut;
    }

    function trim(varText) {
        var i = 0;
        var j = varText.length - 1;

        for (i = 0; i < varText.length; i++) {
            if (varText.substr(i, 1) != " " &&
                varText.substr(i, 1) != "\t")
                break;
        }


        for (j = varText.length - 1; j >= 0; j--) {
            if (varText.substr(j, 1) != " " &&
                varText.substr(j, 1) != "\t")
                break;
        }

        if (i <= j)
            return (varText.substr(i, (j + 1) - i));
        else
            return ("");
    }

    function Status_OnMouseOver() {
        OldBGColor = window.event.srcElement.style.backgroundColor;
        window.event.srcElement.style.backgroundColor = "thistle";
        window.event.srcElement.style.cursor = "hand";

        var node = window.event.srcElement;
        while (node.nodeName.toUpperCase() != "TD") {
            node = node.parentElement;
        }

        window.status = node.StatusID;
    }

    function Status_OnMouseOut() {
        window.event.srcElement.style.backgroundColor = OldBGColor;
    }

    function Status_OnMouseDown() {
        if (event.button == 2) {
            return false;
        }
    }

    function Status_OnClick(StatusID) {
        ShowAgencyStatus(StatusID);
    }

    function ShowAgencyBatchEdit(StatusID, DeliverableRootID) {
        modalDialog.open({ dialogTitle: 'Agency Batch Update', dialogURL: './Agency/BatchEditFrames.asp?ID=' + StatusID + '&DeliverableRootID=' + DeliverableRootID + '', dialogHeight: 675, dialogWidth: 1000, dialogResizable: true, dialogDraggable: true });

        //window.showModalDialog("./Agency/BatchEditFrames.asp?ID=" + StatusID + "&DeliverableRootID=" + DeliverableRootID, "", "dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        //window.location.reload();
    }

    function ShowAgencyStatus(StatusID) {

       modalDialog.open({ dialogTitle: 'Agency Status', dialogURL: './Agency/Agency.asp?ID=' + StatusID + '&from_where=<%=m_from_where %>' , dialogHeight: 500, dialogWidth: 700, dialogResizable: true, dialogDraggable: true });
    }

    function ShowReleaseStatus(CountryID, CountryName, ReleaseID, ReleaseName, ProductID, currentData, platform) {
        modalDialog.open({
            dialogTitle: 'Release Status',
            dialogURL: './Agency/Agency.asp?CountryID=' + CountryID +
                                            "&ReleaseID=" + ReleaseID +
                                            "&ReleaseName=" + ReleaseName +
                                            "&ProductID=" + ProductID +
                                            "&currentData=" + currentData +
                                            "&platform=" + platform +
                                            "&CountryName=" + CountryName,
            dialogHeight: 500,
            dialogWidth: 700,
            dialogResizable: true,
            dialogDraggable: true
        });
    }

    function ShowAgencyStatusResults(strID) {
        //The old window.event.className and innerText no longer worked, replace with JQuery
        var iStatusID = globalVariable.get('agency_status_id');
        //The status has changed, clear the Status Cell's class and text: ---
        $("#TableAgency td[data-statusid='" + iStatusID + "']").attr('class', '');
        $("#TableAgency td[data-statusid='" + iStatusID + "']").text('');

        //Based on the selected Status or Date, clear update Status Cell's class and text: ---
        if (typeof (strID) != "undefined") {
            if (strID == 'C') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Complete');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('COMPLETE');
            }
            if (strID == 'P') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Complete');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('PARTIAL');
            }
            else if (strID == 'L') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Complete');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('LEVERAGED');
            }
            else if (strID == 'NS') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('NotSupported');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('NOT SUPPORTED');
            }
            else if (strID == 'NR') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('NotNeeded');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('NOT REQUESTED');
            }
            else if (strID == 'NC') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('NotNeeded');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('NO CERT NEEDED');
            }
            else if (strID == 'O') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Late');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('OPEN');
            }
            else if (strID == 'SU') {
                $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Late');
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text('SUPPORTED');
            }
            else {
                var Now = new Date();
                var StartDate = new Date();
                if (strID != '') {
                    if (!isNaN(Date.parse(strID))) {
                        StartDate = new Date(Date.parse(strID));
                    }
                }
                if (DateDiff(Now, StartDate, 'd', true) <= 0) {
                    $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Late');
                }
                else {
                    $("#TableAgency td[data-statusid='" + iStatusID + "']").addClass('Cert_Open');
                }
                $("#TableAgency td[data-statusid='" + iStatusID + "']").text(strID);
            }
        }
    }

    function Export(strID) {
        if (txtCurrentFilter.value == "" && cboCategory.value == 0 && txtProductList.value == "")
            window.open(window.location.href + "?FileType=" + strID);
        else
            window.open(window.location.href + "&FileType=" + strID);
    }

    function SaveRegulatory() {
        $.ajax({
            type: 'POST',
            url: window.location.protocol + '//' + window.location.hostname + 
                 '/PulsarAPI/svc/Product/UpdateRegulatoryModel?regulatoryModelText=' +
                 $("#RegulatoryText").val() + '&productVersionId=<%=Request.QueryString("ID")%>',
            success: function (result) {
                if (result == true) {
                    alert("Update Regulatory Model success.");
                }
                else {
                    alert("No data be updated.");
                }
            },
            error: function() {
                alert("Update Regulatory Model fail.");
            }
        });
    }
//-->
</script>
<style>
    .HeaderRow {
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        FONT-FAMILY: Verdana;
        BACKGROUND-COLOR: Beige;
    }

    .Illegal {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: white;
        BACKGROUND-COLOR: red;
    }

    .Embargo {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: white;
        BACKGROUND-COLOR: red;
    }

    .Unknown {
        FONT-FAMILY: Verdana;
        COLOR: black;
        FONT-SIZE: xx-small;
        BACKGROUND-COLOR: gold;
    }

    .Cert_Late {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: white;
        BACKGROUND-COLOR: red;
    }

    .Cert_Open {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: springgreen;
    }

    .Cert_Complete {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        FONT-WEIGHT: bold;
        COLOR: DarkGreen;
    }

    .NotNeeded {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
    }

    .NotSupported {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: white;
        BACKGROUND-COLOR: red;
    }

    .Changed {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: yellow;
    }

    .AgencyCell {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
    }

    .Region {
        FONT-FAMILY: Verdana;
        FONT-WEIGHT: bold;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: MediumAquamarine;
    }

    .NS_AgencyCell {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: LightSteelBlue;
    }

    .NS_Cert_Late {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: LightSteelBlue;
    }

    .NS_Cert_Open {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: LightSteelBlue;
    }

    .NS_Cert_Complete {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: DarkGreen;
        BACKGROUND-COLOR: LightSteelBlue;
    }

    .NS_NotNeeded {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: LightSteelBlue;
    }

    .NS_NotSupported {
        FONT-FAMILY: Verdana;
        FONT-SIZE: xx-small;
        COLOR: black;
        BACKGROUND-COLOR: LightSteelBlue;
    }
</style>

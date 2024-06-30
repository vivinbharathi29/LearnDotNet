<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<%

Response.AddHeader "Pragma", "No-Cache"

Dim m_Table
Dim m_ListBox
Dim m_IsSysAdmin
Dim m_IsDeliverableOwner
Dim m_EditModeOn

m_IsSysAdmin = false
m_IsDeliverableOwner = false
m_EditModeOn = false

Sub Main()
	If ucase(request("SaveMode")) = ucase("true") Then
		SaveChanges
	Else
		DrawScreen
	End If
End Sub

Sub SaveChanges()
	Dim Security, sUserFullName
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsDeliverableOwner = Security.IsDeliverableOwner("", Trim(Request("drid")))
	sUserFullName = Security.CurrentUser()
	
	If m_IsSysAdmin Or m_IsDeliverableOwner Then
		m_EditModeOn = True
	End If
	
	Set Security = Nothing

	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Dim obj
	Dim objCheckbox
	Dim dw
	Dim cn
	Dim cmd
	Dim RecordCount
	Dim asRecordData
	Dim ReturnValue
	Dim StatusValue
	Dim LeveragedValue
	Dim ProjectedDate
	Dim bUpdateRecord

	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	
	for each obj in Request.Form
		If Left(obj,3) = "hid" And IsNumeric(right(obj, len(obj)-3)) Then
			objCheckbox = "cbx" & right(obj, len(obj)-3)
			
			Response.Write obj & " : " & Request.Form(obj) & "<BR>"
			Response.Write objCheckbox & " : " & Request.Form(objCheckbox) & "<BR>"
						 
			asRecordData = Split(Request.Form(obj),"|")
			
			
			
			Select Case asRecordData(4)
				Case "NS"
					bUpdateRecord = False
				Case "C"
					bUpdateRecord = False
				Case "L"
					If Request.Form(objCheckbox) = "on" Then
						If asRecordData(8) = asRecordData(1) Then
							bUpdateRecord = False
						Else
							bUpdateRecord = True
						End If
					Else
						If asRecordData(8) <> asRecordData(1) Then
							bUpdateRecord = False
						Else
							bUpdateRecord = True
						End If
					End If
				Case "O"
					If Request.Form(objCheckbox) = "on" Then
						bUpdateRecord = True
					Else
						bUpdateRecord = False
					End If
                Case "SU"
					If Request.Form(objCheckbox) = "on" Then
						bUpdateRecord = True
					Else
						bUpdateRecord = False
					End If
				Case Else
			End Select

			If Request.Form(objCheckbox) = "on" Then
				StatusValue = "L"
				LeveragedValue = asRecordData(1)
				ProjectedDate = ""
			Else
				StatusValue = "SU"
				LeveragedValue = ""
				ProjectedDate = asRecordData(7)
			End If
			response.Write bUpdateRecord & "<br>"
            'bUpdateRecord = false

			If bUpdateRecord Then
			
				Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAgencyStatus")
				dw.CreateParameter cmd, "@p_AgencyStatusID", adInteger, adParamInput, 8, asRecordData(5)
                dw.CreateParameter cmd, "@p_SelectedProducts", adVarChar, adParamInput, 5000, ""
                dw.CreateParameter cmd, "@p_SelectedCountries", adVarChar, adParamInput, 5000, ""
                dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
				dw.CreateParameter cmd, "@p_StatusCd", adChar, adParamInput, 5, StatusValue
				dw.CreateParameter cmd, "@p_ProjectedDate", adDate, adParamInput, 8, ProjectedDate
				dw.CreateParameter cmd, "@p_ActualDate", adDate, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_CertificationNo", adVarChar, adParamInput, 50, ""
				dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, LeveragedValue
				dw.CreateParameter cmd, "@p_Notes", adVarChar, adParamInput, 50, ""
				dw.CreateParameter cmd, "@p_TestOrganizer", adInteger, adParamInput, 8, ""
				dw.CreateParameter cmd, "@p_TestBudget", adInteger, adParamInput, 8, ""
				RecordCount = dw.ExecuteNonQuery(cmd)
			
                Response.Write "Records Updated:" & RecordCount & "<br>" & "<br>"

				Set cmd = nothing
			End If
						
			If trim(asRecordData(0)) = trim(request("cid")) Then
				ReturnValue = asRecordData(1) & "|" & asRecordData(6)
			End If
			
		End If
	next
	'Response.End
	Response.Write "<input type=hidden id=returnValue value=""" & ReturnValue & """>"
	
End Sub

Sub DrawScreen()
	If Len(Trim(Request("DRID"))) = 0 Or Len(Trim(Request("PVID"))) = 0 Or Len(Trim(Request("CID"))) = 0 Then
		Response.Write "<H3>Insufficient information provided.</H3>"
		Response.End
	End If

	Dim Security, sUserFullName
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsDeliverableOwner = Security.IsDeliverableOwner("", Trim(Request("drid")))
	sUserFullName = Security.CurrentUser()
	
	If m_IsSysAdmin Or m_IsDeliverableOwner Then
		m_EditModeOn = True
	End If
	
	Set Security = Nothing
	
	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListAgencyLeveragePlatforms")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID")
	dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request("DRID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
	If rs.eof And rs.bof Then
		Response.Write "<H3>There are no completed certifications for this device.</H3>"
		Response.End
	End If
	
	Dim Rows
	Dim ReturnCode
	Dim ProductVersionID
	Dim ListBox
	Dim sRegion

	ListBox = "<OPTION Value="""">-- Please Make A Selection --</OPTION>"

	Do Until rs.EOF
		'
		' ListBox Code Goes Here
		'
		If Trim(Request("lbxPlatform")) = Trim(rs("product_version_id")) Then
			ListBox = ListBox & "<OPTION Value=""" & rs("product_version_id") & """SELECTED>" & rs("product_version_name") & "</OPTION>"
		Else
			ListBox = ListBox & "<OPTION Value=""" & rs("product_version_id") & """>" & rs("product_version_name") & "</OPTION>"
		End If

		rs.movenext
	Loop
	
	m_ListBox = "<SELECT id=lbxPlatform name=lbxPlatform onchange='window.frmLeverage.submit();'>" & ListBox & "</SELECT>"
	
	rs.close
	set rs = nothing
	
	If len(trim(Request("lbxPlatform"))) > 0 Then
	
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyLeverageCountries")
		dw.CreateParameter cmd, "@p_SourceProductVersionID", adInteger, adParamInput, 8, Request("PVID")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("lbxPlatform")
		dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request("DRID")
		Set rs = dw.ExecuteCommandReturnRS(cmd)	

		If Not rs.eof then
			Rows = "<TABLE cellpadding=2 cellspacing=0><TR bgcolor=wheat><TH><input type=checkbox id=chkAll name=chkAll onClick=""return SelectAll();""></TH><TH align=left>Country</TH>" & _
				"<TH align=left>Status</TH><TH align=left>Projected Date</TH></TR>" '& _
				'"<TR><TD colspan=4><a href=""javascript:SelectAll()"">Toggle All</a></TD></TR>"
		Else
			Rows = "<TABLE cellpadding=2 cellspacing=0><TR bgcolor=wheat><TH>No Leveragable Countries Found</TH></TR>"
		End If

		sRegion = ""
		
		Do Until rs.eof
			If sRegion <> rs("region") Then
				Rows = Rows & _
					"<TR><TD class=""Region"" ColSpan=4>" & Replace(Replace(rs("region"), "<", ""), ">", "") & "</TD></TR>"
				sRegion = rs("region")
			End If
			Rows = Rows & _
				"<TR><TD style=""border-top: black thin solid;"">" & _
				"<input type=hidden id=hid" & rs("agency_status_id") & " name=hid" & rs("agency_status_id") & _
				" value=""" & rs("Country_Id") & "|" & rs("Agency_Status_Id") & "|" & trim(rs("Agency_Status_Cd")) & "|" & rs("Agency_Status_Txt") & "|" & trim(rs("Current_Agency_Status_Cd")) & "|" & trim(rs("current_agency_status_id")) & "|" & trim(rs("product_version_name")) & "|" & rs("Current_Agency_Status_Dt") & "|" & rs("current_leveraged_id") & """>"
			Rows = Rows & _
				"<INPUT onclick='return checkIndeterminate();' type=checkbox id=cbx" & rs("agency_status_id") & " name=cbx" & rs("agency_status_id")
			If Trim(rs("current_agency_status_cd")) = "L" Then
				If rs("current_leveraged_id") = rs("agency_status_id") Then
					Rows = Rows & " CHECKED "
				Else
					Rows = Rows & " INDETERMINATE=-1 VALUE=INDETERMINATE"
				End If
			ElseIf Trim(rs("current_agency_status_cd")) = "NS" Then
				Rows = Rows & " DISABLED "
			End If
			Rows = Rows & _
				"></TD><TD nowrap style=""border-top: black thin solid;"">" & rs("Country") & "&nbsp;&nbsp;</TD>" & _
				"<TD nowrap style=""border-top: black thin solid;"">" & rs("Agency_Status_Txt") & "</TD>" & _
				"<TD width=""100%"" style=""border-top: black thin solid;"">" & rs("Agency_Status_Dt") & "&nbsp;" & "</TD></TR>" & vbcrlf
				
			rs.MoveNext
			
		Loop

		Rows = Rows & "</TABLE>"
	
		m_Table = "<table ID=tblMain WIDTH=100% BGCOLOR=cornsilk BORDER=1 CELLSPACING=0 CELLPADDING=2 bordercolor=tan style=""display:none""><TR><TH width=""50"">&nbsp;</TH><TH width=""400"">Program Name</TH></TR>" & Rows & "</table>"
		
		rs.close
		Set rs = nothing
		
	End If
	
End Sub

%>
<HTML>
<HEAD>
<TITLE>Leveragable Programs</TITLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.6.1.min.js" type="text/javascript"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
<!--

function checkIndeterminate()
{
	if (window.event.srcElement.value == "INDETERMINATE")
	{
		if (confirm("Another platform has already been chosen for this country.\nClick OK to Leverage the currently selected platform or Cancel to keep the existing selection."))
			window.event.srcElement.checked = true;
		else
			window.event.srcElement.indeterminate = true;
	}
}

$(function() {
    var lowerWindow = window.parent.frames["LowerWindow"].document;
    if ($("#returnValue").length) {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.LeverageSearchResult($("#returnValue").val());
            parent.window.parent.modalDialog.cancel();
        } else {
            window.returnValue = $("#returnValue").val();
            window.close();
        }
    }

    if ($("#frmButtons", lowerWindow).length) {
        if ($("#hidEdit").length) {
            if ($("#hidEdit").val().toLowerCase() == "false") {
                $("#cmdOK", lowerWindow).attr('disabled', 'disabled');
            }
        }
    }
});

function SelectAll()
{
with (window.frmLeverage){
	var field;
	for (field in window.frmLeverage)
	{
		if(field.substr(0,3) == "cbx")
		{
				window.frmLeverage[field].checked = chkAll.checked;
		}
	}
}
}
//-->
</SCRIPT>
<STYLE>
.Region
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: MediumAquamarine;
    border-top: black thin solid;
}
</STYLE>
</HEAD>
<BODY bgcolor="ivory">
<form id="frmLeverage" method="post" action=AgencyLeverageMain.asp >
<% Call Main() %>

<%= m_ListBox%>
<BR><BR>
<%= m_Table%>
<input type=hidden id=SaveMode name=SaveMode value=false >
<input type=hidden id=drid name=drid value="<%= Request("drid")%>">
<input type=hidden id=pvid name=pvid value="<%= Request("pvid")%>">
<input type=hidden id=cid name=cid value="<%= Request("cid")%>">
<input type=hidden id=hidEdit name=hidEdit value="<%= m_EditModeOn%>">
</form>
</BODY>
</HTML>

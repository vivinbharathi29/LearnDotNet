<%@ Language=VBScript %>
<%Option Explicit%>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 

<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim m_IsSysAdmin
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_IsSEPMProductsEditor

Dim m_UserFullName
Dim m_EditModeOn
Dim m_ScheduleID
	
Sub Main()
'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	Dim sUserFullName
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	
	m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
	m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	m_UserFullName = Security.CurrentUser()
	
'
'	Security has been checked in the previous dialog
'
	m_EditModeOn = True
	
	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

	m_ScheduleID = Request("ScheduleID")

	If m_ScheduleID = "" Then
		Response.Write "Insufficient Data to process request"
		Response.End
	End If 

	If Request.Form("PostBack") = "True" Then
		Call SaveData()
	End If
End Sub

Sub SaveData()

	Dim dw, cn, cmd, iRowsChanged, sShowOnReports_YN, prodVersionIDs, rs, strSQL, prodGroupIDs, sItemDescription, sCustomItemDefinition, sRTMitemOSReleaseId, OSRolloutType

    prodVersionIDs = ""
    prodGroupIDs = ""
    OSRolloutType = NULL
    Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    If Request.Form("selProdGroup") <> "" Then
        prodGroupIDs = Request.Form("selProdGroup")
        strSQL = "spGetProgramTree '" & prodGroupIDs & "'"

        set rs = server.CreateObject("ADODB.recordset")
        rs.Open strSQL,cn,adOpenForwardOnly

        do while not rs.eof
            prodVersionIDs = rs("ProdID") & "," & prodVersionIDs    'ex: 1,12,8,25, always ends with comma
            rs.movenext
        loop
        rs.close
    End If

    If Request.Form("selProd") <> "" Then
        Dim selProds, i
        selProds = Split(Request.Form("selProd"), ",")
        For i=0 to UBound(selProds)
            If InStr("," & prodVersionIDs, "," & selProds(i) & ",") < 1 Then        ' make sure there is no duplicate ProdVid
                prodVersionIDs = prodVersionIDs & Request.Form("selProd") & ","   'ex: 1,12,8,25,22,50,31, still ends with comma
            End If
        Next
    End If

    If InStr("," & prodVersionIDs, "," & Request("ProdVID") & ",") < 1 Then        ' make sure prodVersionIDs does not already contain current ProdVid
        prodVersionIDs = prodVersionIDs & Request("ProdVID")            'ex: 1,12,8,25,22,50,31,20  assume ProdVID is 20
    Else
		prodVersionIDs = Left(prodVersionIDs, Len(prodVersionIDs) - 1)
    End If

	If Request.Form("cbxShowOnReports") = "on" Then
		sShowOnReports_YN = "Y"
	Else
		sShowOnReports_YN = "N"
	End If 

	If Request.Form("cbxISRtmItem") = "RTM" OR Request.Form("cbxRTP") = "RTP" OR Request.Form("cbxUpgrade") = "Upgrade" Then
		sRTMitemOSReleaseId = Request.Form("selItemOSRelease")
		sItemDescription = Request.Form("SelectedCate") & " " & Request.Form("ReleasesOSName")
	    sCustomItemDefinition = Request.Form("SelectedCate") & " " & Request.Form("ReleasesOSName")
        
        If Request.Form("cbxRTP") = "RTP" Then
            OSRolloutType = 1
        ElseIf Request.Form("cbxUpgrade") = "Upgrade" Then
            OSRolloutType = 2
        End If

	Else 
		sItemDescription = Request.Form("txtItemDescription")
		sCustomItemDefinition = Request.Form("txtCustomItemDefinition")
	End If

	Set dw = New DataWrapper
	cn.BeginTrans

    Set cmd = dw.CreateCommandSP(cn, "usp_InsertScheduleDataIntoProducts")
	dw.CreateParameter cmd, "@p_ProductVersionIDs", adVarChar, adParamInput, 5000, prodVersionIDs
	dw.CreateParameter cmd, "@p_ScheduleDefinitionDataID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(m_UserFullName)
	dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, "Y"
	dw.CreateParameter cmd, "@p_ItemDescription", adVarChar, adParamInput, 500, sItemDescription
	dw.CreateParameter cmd, "@p_ItemPhase", adInteger, adParamInput, 8, Request.Form("selItemPhase")
	dw.CreateParameter cmd, "@p_ItemOwner", adInteger, adParamInput, 8, Request.Form("selItemOwner")
	dw.CreateParameter cmd, "@p_ItemNotes", adVarChar, adParamInput, 5000, NULL
	dw.CreateParameter cmd, "@p_PorStartDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_PorEndDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_ProjectedStartDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_ProjectedEndDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_ActualStartDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_ActualEndDt", adDate, adParamInput, 0, NULL
	dw.CreateParameter cmd, "@p_ShowOnReports_YN", adChar, adParamInput, 1, sShowOnReports_YN
	dw.CreateParameter cmd, "@p_Milestone_YN", adChar, adParamInput, 1, Request.Form("rbMilestone")
	dw.CreateParameter cmd, "@p_CustomItemDefinition", adVarChar, adParamInput, 500, sCustomItemDefinition
	dw.CreateParameter cmd, "@p_RtmItem_OsReleaseId", adVarChar, adParamInput, 8, sRTMitemOSReleaseId
	dw.CreateParameter cmd, "@p_RtmWave", adInteger, adParamInput, 8, Request.Form("RTMWave")
    dw.CreateParameter cmd, "@p_OSRolloutType", adInteger, adParamInput, 8, OSRolloutType
    
    iRowsChanged = dw.ExecuteNonQuery(cmd)

    If iRowsChanged < 1 Then
		cn.RollbackTrans	
		Response.Write "Error Saving Schedule Item"
		Response.End
	Else
		cn.CommitTrans
        dim prodVersionWithoutSchedule
        strSQL = "usp_SelectProductWithoutSchedule '" & prodVersionIDs & "'"
        set rs = server.CreateObject("ADODB.recordset")
        rs.Open strSQL,cn,adOpenForwardOnly
        prodVersionWithoutSchedule = ""
        do while not rs.eof
            prodVersionWithoutSchedule = prodVersionWithoutSchedule & "<li>" & rs("ProductVersionName") & "</li>"
            rs.movenext
        loop
        rs.close
        
        if prodVersionWithoutSchedule <> "" then
            Response.Write "The following products do not have any schedule:"
            Response.Write "<ul>" & prodVersionWithoutSchedule & "</ul>"
            Response.End
        else
            Response.Write "<input type=hidden id=CloseOnLoad value=True>"        
        end if
		
	End If

    Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
    
End Sub

Sub FillPhaseList()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListSchedulePhases")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("schedule_phase_id") & """>" & rs("phase_name") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub

Sub FillOwnerList()
	Dim dw, cn, cmd, rs
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "select item_description, schedule_definition_data_id from schedule_definition_data WHERE GenericOwner = 1")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("schedule_definition_data_id") & """>" & rs("item_description") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub

Sub FillProductList()
	Dim dw, cn, cmd, rs
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "usp_ListProducts")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("ProductVersionID") & """>" & rs("ProductVersionName") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub  

Sub FillProductGroupList()
	Dim dw, cn, cmd, rs
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "spListPrograms2")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("ID") & """>" & rs("FullName") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub  

Sub FillOSRelease()
	Dim dw, cn, cmd, rs
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSQL(cn, "select ID, Description from OSRelease")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	Do Until rs.eof
		Response.Write "<OPTION Value=""" & rs("ID") & """>" & rs("Description") & "</OPTION>"
		rs.MoveNext
	Loop
	
	rs.close
	
	Set rs = Nothing
	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing	
	
End Sub

%>
	
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Add Custom Item</title>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cmdCancel_onclick() {
        if (window.location != window.parent.location) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.close();
        }
    }

    function VerifySave() {
        with (window.frmMilestone) {
            if (!validateTextInput(selItemPhase, 'Phase')) { return false; }
            if (document.getElementById("cbxISRtmItem").checked == true ||
                document.getElementById("cbxRTP").checked == true || 
                document.getElementById("cbxUpgrade").checked == true) {
				if (!validateTextInput(selItemOSRelease, 'Phase')) { return false;}
			} else {
				if (!validateTextInput(txtItemDescription, 'Description')) { return false; }
				if (!validateTextInput(txtCustomItemDefinition, 'Definition')) { return false; }
			}
			if (!validateTextInput(selItemOwner, 'Owner')) { return false; }
            if (!((rbMilestone[0].checked) || (rbMilestone[1].checked))) { alert('Task or Milestone Selection Required'); return false; }
        }
        return true;
    }

    function cmdOK_onclick() {
        if (VerifySave()) {
            window.frmMilestone.cmdCancel.disabled = true;
            window.frmMilestone.cmdOK.disabled = true;
            window.frmMilestone.submit();
        }
    }

    function window_onLoad() {
        if (window.CloseOnLoad) {
            if (window.location != window.parent.location) {
                parent.window.parent.modalDialog.cancel(true);
            } else {
                window.returnValue = 1;
                window.close();
            }
        }
    }
	
	function cbxISRtmItem_onclicked() {
		var checkBox = document.getElementById("cbxISRtmItem");
		
		if (checkBox.checked == true){
			document.getElementById("txtItemDescription").value="";
			document.getElementById("txtCustomItemDefinition").value="";
			document.getElementById("txtItemDescription").disabled = true;
			document.getElementById("txtCustomItemDefinition").disabled = true;
            document.getElementById("selItemOSRelease").disabled = false;
            document.getElementById("cbxRTP").disabled = true;
            document.getElementById("cbxUpgrade").disabled = true;
            document.getElementById("SelectedCate").value = checkBox.value;  
		} else {
			document.getElementById("selItemOSRelease").selectedIndex = 0;
			document.getElementById("txtItemDescription").disabled = false;
            document.getElementById("txtCustomItemDefinition").disabled = false;
            document.getElementById("selItemOSRelease").value = "";
            document.getElementById("selItemOSRelease").disabled = true;
            document.getElementById("cbxRTP").disabled = false;
            document.getElementById("cbxUpgrade").disabled = false;
            document.getElementById("SelectedCate").value = "";
		}
    }

    function CbxRTPItem_onclicked() {
        var checkBox = document.getElementById("cbxRTP");

        if (checkBox.checked == true) {
            document.getElementById("cbxISRtmItem").disabled = true;
            document.getElementById("cbxUpgrade").disabled = true;
            document.getElementById("selItemOSRelease").disabled = false;
            document.getElementById("txtItemDescription").value="";
			document.getElementById("txtCustomItemDefinition").value="";
			document.getElementById("txtItemDescription").disabled = true;
			document.getElementById("txtCustomItemDefinition").disabled = true;
            document.getElementById("SelectedCate").value = checkBox.value;
            document.getElementById("RTMWave").parentNode.parentElement.style.display = "block";
        }
        else {
            document.getElementById("cbxISRtmItem").disabled = false;
            document.getElementById("cbxUpgrade").disabled = false;
            document.getElementById("selItemOSRelease").value = "";
            document.getElementById("selItemOSRelease").disabled = true;
            document.getElementById("txtItemDescription").disabled = false;
			document.getElementById("txtCustomItemDefinition").disabled = false;
            document.getElementById("SelectedCate").value = "";
            document.getElementById("RTMWave").parentNode.parentNode.style.display = "none";
        }
    }

    function CbxUpgradeItem_onclicked() {
        var checkBox = document.getElementById("cbxUpgrade");

        if (checkBox.checked == true) {
            document.getElementById("cbxISRtmItem").disabled = true;
            document.getElementById("cbxRTP").disabled = true;
            document.getElementById("selItemOSRelease").disabled = false;
            document.getElementById("txtItemDescription").value="";
			document.getElementById("txtCustomItemDefinition").value="";
			document.getElementById("txtItemDescription").disabled = true;
			document.getElementById("txtCustomItemDefinition").disabled = true;
            document.getElementById("SelectedCate").value = checkBox.value;
        }
        else {
            document.getElementById("cbxISRtmItem").disabled = false;
            document.getElementById("cbxRTP").disabled = false;
            document.getElementById("selItemOSRelease").value = "";
            document.getElementById("selItemOSRelease").disabled = true;
            document.getElementById("txtItemDescription").disabled = false;
			document.getElementById("txtCustomItemDefinition").disabled = false;
            document.getElementById("SelectedCate").value = "";
        }
    }

	function selItemOSRelease_onChanged() {
		var e = document.getElementById("selItemOSRelease");
		var selectTxt = e.options[e.selectedIndex].text;
        document.getElementById("ReleasesOSName").value = selectTxt;
    }	

    function filterDigital(e, pnumber) {
        if (!/^\d+$/.test(pnumber)) {
            var newValue = /^\d+/.exec(e.value);
            if (newValue != null) {
                e.value = newValue;
            }
            else {
                e.value = "";
            }
        }
        return false;
    }
    //-->
</SCRIPT>
<LINK href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY bgcolor=ivory leftMargin=9 topMargin=9 OnLoad="window_onLoad()">
<% Call Main() %>
	<h3>Add Custom Item</h3>
	<Form ID="frmMilestone" method="post">
	    <table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		    <tr>
			    <td nowrap valign="top"><b>Phase:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td>
			        <SELECT id=selItemPhase name=selItemPhase>
				        <OPTION value="">--- Select Phase ---</OPTION>
				        <% Call FillPhaseList() %>
			        </SELECT>
			    </td>
		    </tr>
            <tr>
                <td nowrap valign="top"><b>RTM Item:</b></td>
                <td>
                    <input type="checkbox" id="cbxISRtmItem" name="cbxISRtmItem" value="RTM" onclick="return cbxISRtmItem_onclicked()">
                </td>
            </tr>
            <tr>
                <td style="white-space: nowrap; vertical-align:top;"><b>RTP:</b></td>
                <td>
                    <input type="checkbox" id="cbxRTP" name="cbxRTP" value="RTP" onclick="return CbxRTPItem_onclicked()">
                </td>
            </tr>
            <tr>
                <td style="white-space: nowrap; vertical-align:top;"><b>Upgrade:</b></td>
                <td>
                    <input type="checkbox" id="cbxUpgrade" name="cbxUpgrade" value="Upgrade" onclick="return CbxUpgradeItem_onclicked()">
                </td>
            </tr>
		    <tr id="DescriptionDr">
			    <td width="150" nowrap><b>Item Description:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td width="100%">
			    <INPUT type="text" id=txtItemDescription name=txtItemDescription size=40 maxlength=500>
			    </td>
		    </tr>
		    <tr id="DefinitionDr">
			    <td width="150" nowrap><b>Definition:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td width="100%">
			    <INPUT type="text" id=txtCustomItemDefinition name=txtCustomItemDefinition size=40 maxlength=500>
			    </td>
		    </tr>			
		    <tr>
			    <td nowrap valign="top"><b>Releases for Operating System:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td>
			    <SELECT id=selItemOSRelease name=selItemOSRelease disabled onchange='selItemOSRelease_onChanged()'>
				    <OPTION value="">--- Select ---</OPTION>
				    <% Call FillOSRelease() %>
			    </SELECT>
			    </td>
		    </tr>
            <tr style="display:none">
			    <td nowrap valign="top"><b>RTM Wave:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td>
                    <input type="text" id="RTMWave" value="" name="RTMWave" size="10" maxlength="5" onkeyup="return filterDigital(this,value)" />
			    </td>
		    </tr>
		    <tr>
			    <td nowrap valign="top"><b>Owner:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td>
			    <SELECT id=selItemOwner name=selItemOwner>
				    <OPTION value="">--- Select Owner ---</OPTION>
				    <% Call FillOwnerList() %>
			    </SELECT>
			    </td>
		    </tr>
	        <tr>
			    <td nowrap valign="top"><b>Task / Milestone:</b>&nbsp;<font color="red" size="1">*</font></td>
			    <td>
			        <INPUT type="radio" id=rbMilestone name=rbMilestone value="N">Task&nbsp;
			        <INPUT type="radio" id=Radio1 name=rbMilestone value="Y">Milestone
			    </td>
	        </tr>
            <tr>
                <td nowrap valign="top"><b>Show On Product Status Reports:</b></td>
                <td>
                    <input type="checkbox" id="cbxShowOnReports" name="cbxShowOnReports" checked />
                </td>
            </tr>
		    <tr>
			    <td nowrap valign="top"><b>Add to Other Product(s):</b></td>
			    <td>
			        <SELECT id=selProd name=selProd multiple size="6">
				        <% Call FillProductList() %>
			        </SELECT>
			    </td>
		    </tr>
		    <tr>
			    <td nowrap valign="top"><b>Product Groups:</b></td>
			    <td>
			        <SELECT id=selProdGroup name=selProdGroup multiple size="6">
				        <% Call FillProductGroupList() %>
			        </SELECT>
			    </td>
		    </tr>
	    </table>
        <br />
        <table width="100%" border=0>
          <tr>
              <TD align=right>
                <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
                <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
              </TD>
          </tr>
        </table>
        <INPUT type="hidden" id="PostBack" name="PostBack" value="True">
        <INPUT type="hidden" id="ScheduleID" name="ScheduleID" value="<%= m_ScheduleID%>">
        <INPUT type="hidden" id="ReleasesOSName" Name="ReleasesOSName">
        <INPUT type="hidden" id="SelectedCate" Name="SelectedCate">
    </form>
</BODY>
</HTML>

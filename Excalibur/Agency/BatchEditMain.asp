<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/EmailWrapper.asp" --> 
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<%

Response.AddHeader "Pragma", "No-Cache"

Dim m_IsSysAdmin
Dim m_IsDeliverableOwner
Dim m_EditModeOn
Dim m_FormSave
Dim m_FormClose
Dim m_FormDisplay
Dim m_CurrentUserEmail
Dim m_ProductID
Dim m_DcrID
Dim m_IsSystemTeamLead
Dim strProducts
Dim strCountries
Dim m_DeliverableRootID
Dim arDCRs

m_FormSave = Trim(Request.Form("hidSave"))
If Len(Trim(m_FormSave)) = 0 Then
	m_FormSave = False
End If

m_FormClose = Trim(Request.Form("hidClose"))
If Len(Trim(m_FormClose)) = 0 Then
	m_FormClose = False
End If

m_DeliverableRootID = Request("DeliverableRootID")

m_IsSysAdmin = False
m_IsDeliverableOwner = False
m_EditModeOn = False
m_FormDisplay = False
m_IsSystemTeamLead = False

Sub Main()
	If m_FormSave Then
		Call SaveData()
	'Else
	'	Call DisplayData()
	End If
End Sub

Sub SaveData()
    Dim dw
    Dim cn
    Dim cmd
    Dim RecordCount

    Dim Security, sUserFullName
	Set Security = New ExcaliburSecurity
	
	sUserFullName = Security.CurrentUser()
        
    Set dw = New DataWrapper
    Set cn = server.CreateObject("ADODB.Connection")
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    cn.BeginTrans()
		        
    Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAgencyStatus")
    dw.CreateParameter cmd, "@p_AgencyStatusID", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_SelectedProducts", adVarChar, adParamInput, 5000, Trim(Request("hidSelectedProducts"))
    dw.CreateParameter cmd, "@p_SelectedCountries", adVarChar, adParamInput, 5000, Trim(Request("hidSelectedCountries"))
    dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Trim(Request("hidDeliverableRootID"))
    dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
    dw.CreateParameter cmd, "@p_StatusCd", adChar, adParamInput, 5, Trim(Request("cboStatus"))
    dw.CreateParameter cmd, "@p_ProjectedDate", adDate, adParamInput, 8, Trim(Request("txtProjectedDate"))
    dw.CreateParameter cmd, "@p_ActualDate", adDate, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_CertificationNo", adVarChar, adParamInput, 50, ""
    dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_Notes", adVarChar, adParamInput, 5000, Trim(Request("txtNotes"))
    dw.CreateParameter cmd, "@p_TestOrganizer", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_TestBudget", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_POR_DCR", adChar, adParamInput, 3, Trim(Request("cboPorDcr"))
    dw.CreateParameter cmd, "@p_Dcr_Id", adInteger, adParamInput, 8, Trim(Request("cboDcr"))
    RecordCount = dw.ExecuteNonQuery(cmd)
        	
    Set cmd = Nothing

    If RecordCount = 0 Then
        Response.Write "<H3>Error Updating Record</H3>"
        m_FormDisplay = False
		m_FormClose = False
        cn.RollbackTrans()
        Exit Sub
    End If
    cn.CommitTrans()

	If Not m_FormClose Then
		m_FormDisplay = True
	End If

End Sub

Sub GenerateHiddenDcrCache()
    Response.Write "<input type=""hidden"" id=""hidDCRs"" name=""hidDCRs"" value=""" 
	FillDcrStatus()
    Response.Write """>"

End Sub 


Sub FillDcrStatus()
	Dim dw, cn, cmd, rs, i, strLastProduct

	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListApprovedDCRsByRoot")
	dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request("DeliverableRootID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
    i = 0
    arDCRs = ""

	Do until rs.eof
        Response.Write rs("ID") & ";;" & server.HTMLEncode(rs("Summary")) & ";;" & rs("DotsName") & ";;" & rs("PVID")				
        Response.Write "||"
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub



'Populate Available Products
Sub FillProducts(ProductID)
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spGetProductsForRoot")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("DeliverableRootID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(rs("productstatusid")) < 4 then
			Response.Write "<Option value= """ & rs("ID") & """>" & rs("Name") & " " & rs("Version")  & "</OPTION>"					
		end if
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

'Populate Available Countries
Sub FillCountries()
	Dim dw, cn, cmd, rs, region
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spGetCountries")
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	region = ""
	Do until rs.eof
        if rs("Active") then
            if (rs("Region2") <> region) then
                Response.Write "<Option value="""">------------" & rs("Region2") & "-------------</Option>"
                region = rs("Region2")
            end if
			Response.Write "<Option value= """ & rs("ID") & """>" & rs("Language") & "</OPTION>"					
		end if
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

%>

<HTML>
<HEAD>


<script language="JavaScript" src="../_ScriptLibrary/jsrsClient.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var cboStatus_LastIndex;

//Copied working date logic from AgencyMain.asp
function cmdDate_onclick(FieldID) {
    var strID;
    var oldValue = window.frmStatus.elements(FieldID).value;

    strID = window.showModalDialog("../mobilese/today/caldraw1.asp", FieldID, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) == "undefined")
        return

    window.frmStatus.elements(FieldID).value = strID;

    var newDate = Date.parse(strID);
    var oldDate = Date.parse(oldValue);

    var dcrID;
    var statusID = window.frmStatus.agency_status_id.value;

    if (newDate > oldDate) {
        dcrID = window.showModalDialog("ChooseDCR.asp?ID=" + programID + "&StatusID=" + statusID, FieldID, "dialogWidth:700px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    }
}

 function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight,datainput) {
            if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
            if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
            $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
            $("#modalDialog").attr("width", "98%");
            $("#modalDialog").attr("height", "98%");
            $("#modalDialog").attr("src", QueryString);
            $("#iframeDialog").dialog("option", "title", Title);
	    $("#iframeDialog").data('dataid', datainput);
            $("#iframeDialog").dialog('open');
        }

function ClosePropertiesDialog(strID) {

        $("#iframeDialog").dialog("close");
       // if (typeof (strID) != "undefined") window.location.reload(true);
    }

function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }

function window_onload() {
    if (window.frmStatus.hidClose.value.toLowerCase() == 'true') {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
            parent.window.parent.location.reload();
        } else {
            parent.window.parent.CloseIframeDialog();
            parent.window.parent.location.reload();
        }
    }
}

//Called from BatchEditButtons when saved
function GetSelectedProducts() {
    var i;
    var strProducts = "";

    //Product List
    for (i = 0; i < frmStatus.lstProducts.length; i++)
        if (frmStatus.lstProducts.options[i].selected)
            strProducts = strProducts + "," + frmStatus.lstProducts.options[i].value;
    if (strProducts.length > 0)
        strProducts = strProducts.substr(1);
    frmStatus.hidSelectedProducts.value = strProducts;
    //alert(frmStatus.hidSelectedProducts.value);
}

//Called from BatchEditButtons when saved
function GetSelectedCountries() {
    var i;
    var strCountries = "";

    //Country List
    for (i = 0; i < frmStatus.lstCountries.length; i++)
        if (frmStatus.lstCountries.options[i].selected)
            strCountries = strCountries + "," + frmStatus.lstCountries.options[i].value;
    if (strCountries.length > 0)
        strCountries = strCountries.substr(1);
    frmStatus.hidSelectedCountries.value = strCountries
    
}

function cboPorDcr_onchange() {
    if (window.frmStatus.cboPorDcr.value == 'DCR') {
        GetSelectedProducts();
        LoadDCRsForSelectedProducts();
        if (frmStatus.hidSelectedProducts.value == "") {
            alert("Please select a Product for the Batch Update")
            window.frmStatus.cboPorDcr.value = "";
        } else {
            document.getElementById("cboDcr").disabled = false;
        }
    } else {
        document.getElementById("cboDcr").disabled = true;
        var sel = document.getElementById("cboDcr");
        sel.options.length = 0;
    }        
}

function lstProducts_onchange() {
    if (window.frmStatus.cboPorDcr.value == 'DCR') {
        GetSelectedProducts();
        LoadDCRsForSelectedProducts();
    }
}

function LoadDCRsForSelectedProducts() {
    var i;
    var j;
    var k;
    var arDropdownValues;
    var strSelectedProducts;
    var strDCRs;
    var intLastProduct = "";
    var opt;

    var sel = document.getElementById("cboDcr");
    sel.options.length = 0;
    
    strSelectedProducts = frmStatus.hidSelectedProducts.value;
    arProducts = strSelectedProducts.split(",");

    strDCRs = frmStatus.hidDCRs.value;
    arDCRs = strDCRs.split("||");

    for (i = 0; i < arProducts.length; i++) {
        for (j = 0; j < arDCRs.length; j++) {
            arDropdownValues = arDCRs[j].split(";;");

            
            if (arProducts[i] == arDropdownValues[3]) {
                if (intLastProduct != arDropdownValues[3]) {
                    opt = document.createElement('option');
                    opt.innerHTML = "------------" + arDropdownValues[2] + "-------------";
                    opt.value = "";
                    sel.appendChild(opt);
                    intLastProduct = arDropdownValues[3];
                }
                opt = document.createElement('option');
                opt.innerHTML = arDropdownValues[0] + ": " + arDropdownValues[1];
                opt.value = arDropdownValues[0];
                sel.appendChild(opt);
            }
        }
    }
}
//-->
</SCRIPT>

    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />

</HEAD>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<form id="frmStatus" method="post" action=BatchEditMain.asp><% Call Main() %>

<font face=verdana size=2><b><span style="font:bold x-small verdana">Agency Batch Update</span></b></font>

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
	    <td>&nbsp;</td>
		<td valign=top width=140 nowrap><b>Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td valign=top width=140 nowrap><b>Availability Date:</b></td>
		<td valign=top width=140 nowrap><b>Added By:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td valign=top width=140 nowrap><b>Added By DCR:</b>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
			<SELECT id=cboStatus name=cboStatus>
			<option value=SU>Supported</option>
			<option value=P>Partial</option>
			<option value=C>Complete</option>
			<option value=NS>Not Supported</option>
			<option value=NR>Not Requested</option>  
			<option value=NC>No Cert Needed</option>
			</SELECT>
		</td>
		<td>
			<INPUT type="text" id=txtProjectedDate name=txtProjectedDate value="" style="width:90px" >&nbsp;
            <a href="javascript: cmdDate_onclick('txtProjectedDate')">
                <img ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21">
            </a>
		</td>
		<td>
			<SELECT id=cboPorDcr name=cboPorDcr onchange="cboPorDcr_onchange();">
            <option value="">------------Select One------------</option>
			<option value="POR">POR</option>
			<option value="DCR">DCR</option>
			</SELECT>
		</td>
		<td>
			<SELECT id=cboDcr name=cboDcr style="width:200px" disabled>
		
			</SELECT>
		</td>
	</tr>
		<tr>
	    <td style="text-align:right; vertical-align:top;"><b>Notes:</b>&nbsp;</td>
		<td valign=top colspan=4 nowrap>
            <textarea id="txtNotes" name="txtNotes" style="width:100%;" rows="2"></textarea></td>
	</tr>
    <tr>
        <td>&nbsp;</td>
	    <td valign=top width=140 nowrap><b>Products:</b></td>
        <td valign=top width=140 nowrap><b>Countries:</b></td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td valign=top width=140 nowrap>
            <select style="WIDTH: 180px; HEIGHT: 380px" multiple id="lstProducts" name="lstProducts" onchange="return lstProducts_onchange()">
				<%FillProducts(m_ProductID)%>
			</select>
		</td>
		<td valign=top width=140 nowrap>
            <select style="WIDTH: 180px; HEIGHT: 380px" multiple id="lstCountries" name="lstCountries">
				<%FillCountries()%>
			</select>
		</td>
	</tr>

	<% Call Main() %>
</table>
<input type="hidden" id="hidDisplay" name="hidDisplay" value="<%= m_FormDisplay%>">
<input type="hidden" id="hidSave" name="hidSave" value="<%= m_FormSave%>">
<input type="hidden" id="hidClose" name="hidClose" value="<%= m_FormClose%>">
<input type="hidden" id="hidEdit" name="hidEdit" value="<%= m_EditModeOn%>">
<input type="hidden" id="hidDeliverableRootID" name="hidDeliverableRootID" value="<%= m_DeliverableRootID%>">    
<input type="hidden" id="hidSelectedProducts" name="hidSelectedProducts">
<input type="hidden" id="hidSelectedCountries" name="hidSelectedCountries">
<% GenerateHiddenDcrCache() %>

<div style="display: none;">
    <div id="iframeDialog" title="ExtendTables" >
        <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
    </div>
</div>
</form>

</BODY>
</HTML>



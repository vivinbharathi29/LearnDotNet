<%@  language="VBScript" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!--#include file="../includes/Security.asp"-->
<!--#include file="../includes/DataWrapper.asp"-->
<!--#include file="../includes/no-cache.asp"-->
<%
Dim AppRoot

'##############################################################################	
'
' Create Security Object to get User Info
'
Dim regEx
Set regEx = New RegExp
regEx.Global = True

regEx.Pattern = "[^0-9]"
Dim PVID : PVID = regEx.Replace(Request.QueryString("PVID"), "")

Dim m_IsSysAdmin
Dim m_IsGplm
Dim m_IsBomAnalyst
Dim m_EditModeOn
Dim m_UserFullName
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
    m_IsGplm = Security.UserInRole(PVID, "GPLM")
    m_IsBomAnalyst = Security.UserInRole(PVID, "SBA")

	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsGplm Or m_IsBomAnalyst Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		sMode = "view"
	End If

	Set Security = Nothing
'##############################################################################	

AppRoot = Session("ApplicationRoot")





'*********************************************************************************************************************************************************************************************************************
' LIMIT OSSP USERS (PARTNERTYPEID=2) TO READ ONLY ACCESS
'
'*********************************************************************************************************************************************************************************************************************
Dim CurrentUser : CurrentUser = lcase(Session("LoggedInUser"))
Dim CurrentDomain
Dim CurrentPartnerTypeID
Dim blnIsOSSPUser: blnIsOSSPUser=False

If instr(CurrentUser,"\") > 0 Then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "usp_GetUserType")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 30, CurrentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then

	CurrentPartnerTypeID=trim(rs("PartnerTypeID"))

	If(CurrentPartnerTypeID = "2" OR CurrentPartnerTypeID = 2) Then
		blnIsOSSPUser=true
	End If
	
	rs.Close
End If

set rs=nothing

cn.Close
set cn=nothing
set cmd=nothing

'*********************************************************************************************************************************************************************************************************************

If Request.Form("Action") = "save" And Not blnIsOSSPUser Then
    Call SaveSpareKit()
Else 
    Call DisplaySpareKit(blnIsOSSPUser)
End If

Sub SaveSpareKit

    If Not m_EditModeOn Then
        Response.Write "<span style=""font:bold large verdana;color:darkred"">Insufficient Security Privileges</span>"
        Response.End
    End If
    
    Dim spareKitId
    Dim serviceFamilyPn
    Dim productVersionId
    Dim deliverableRootId
    Dim spareKitNo
    Dim spareKitCategoryId
    Dim spareKitDescription
    Dim spareKitNotes
    Dim spareKitComments
    Dim spareKitCsrLevel
    Dim spareKitWarrantyTier
    Dim spareKitDisposition
    Dim spareKitGeoNa
    Dim spareKitGeoLa
    Dim spareKitGeoApj
    Dim spareKitGeoEmea
    Dim spareKitStockAdvice
    Dim spareKitPpcProductLine
    Dim spareKitStatus
    Dim spareKitFirstServiceDt
    Dim spareKitMfgSubAssembly
    Dim spareKitSvcSubAssembly
    
    spareKitId = request.Form("SKID")
    serviceFamilyPn = request.Form("SFPN")
    productVersionId = request.Form("PVID")
    deliverableRootId = request.Form("DRID")
    spareKitNo = trim(request.Form("spsPartNo"))
    spareKitCategoryId = request.Form("spsCategory")
    spareKitDescription = request.Form("spsDescription")
    If(Len(Trim(spareKitDescription))>40) Then
        spareKitDescription=Mid(Trim(spareKitDescription),1,40)
    End If
    spareKitNotes = request.Form("spsNotes")
    spareKitComments = request.Form("spsComments")
    spareKitCsrLevel = request.Form("spsCsrLevel")
    spareKitWarrantyTier = request.Form("spsWarranty")
    spareKitDisposition = request.Form("spsDisposition")
    spareKitMfgSubAssembly = trim(request.Form("spsMfgSubAssy"))
    spareKitSvcSubAssembly = trim(request.Form("spsSvcSubAssy"))
    spareKitGeoNa = request.Form("spsGeosNa")
    spareKitGeoLa = request.Form("spsGeosLa")
    spareKitGeoApj = request.Form("spsGeosApj")
    spareKitGeoEmea = request.Form("spsGeosEmea")
    spareKitStockAdvice = request.Form("spsLocalStockAdvice")
    spareKitPpcProductLine = ""
    spareKitStatus = ""
    spareKitFirstServiceDt = request.Form("spsFirstServiceDt")
    If Not IsDate(spareKitFirstServiceDt) Then spareKitFirstServiceDt = null
    
    Response.Write spareKitId & "<br>"
    Response.Write serviceFamilyPn & "<br>"
    Response.Write productVersionId & "<br>"
    Response.Write spareKitNo & "<br>"
    Response.Write spareKitCategoryId & "<br>"
    Response.Write spareKitDescription & "<br>"
    Response.Write spareKitNotes & "<br>"
    Response.Write spareKitComments & "<br>"
    Response.Write spareKitCsrLevel & "<br>"
    Response.Write spareKitWarrantyTier & "<br>"
    Response.Write spareKitDisposition & "<br>"
    Response.Write spareKitGeoNa & "<br>"
    Response.Write spareKitGeoLa & "<br>"
    Response.Write spareKitGeoApj & "<br>"
    Response.Write spareKitGeoEmea & "<br>"
    Response.Write spareKitMfgSubAssembly & "<br>"
    Response.Write spareKitSvsSubAssembly & "<br>"

    Dim dw, cn, cmd
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cmd = dw.CreateCommandSp(cn, "usp_UpdateSpareKit")

    dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, CLng(productVersionId)
    dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 0, spareKitId
    dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, spareKitNo
    dw.CreateParameter cmd, "@p_Description", adVarChar, adParamInput, 100, spareKitDescription
    dw.CreateParameter cmd, "@p_SpareCategoryId", adInteger, adParamInput, 0, CLng(spareKitCategoryId)
    dw.CreateParameter cmd, "@p_CsrLevelId", adInteger, adParamInput, 0, CLng(spareKitCsrLevel)
    dw.CreateParameter cmd, "@p_Disposition", adInteger, adParamInput, 0, CLng(spareKitDisposition)
    dw.CreateParameter cmd, "@p_WarrantyTier", adChar, adParamInput, 1, spareKitWarrantyTier
    dw.CreateParameter cmd, "@p_LocalStockAdvice", adInteger, adparamInput, 0, CLng(spareKitStockAdvice)
    dw.CreateParameter cmd, "@p_PpcProductLine", adChar, adParamInput, 2, spareKitPpcProductLine
    dw.CreateParameter cmd, "@p_MfgSubAssembly", adChar, adParamInput, 10, spareKitMfgSubAssembly
    dw.CreateParameter cmd, "@p_SvcSubAssembly", adChar, adParamInput, 10, spareKitSvcSubAssembly
    dw.CreateParameter cmd, "@p_GeoNa", adBoolean, adParamInput, 0, spareKitGeoNa
    dw.CreateParameter cmd, "@p_GeoLa", adBoolean, adParamInput, 0, spareKitGeoLa
    dw.CreateParameter cmd, "@p_GeoApj", adBoolean, adParamInput, 0, spareKitGeoApj
    dw.CreateParameter cmd, "@p_GeoEmea", adBoolean, adParamInput, 0, spareKitGeoEmea
    dw.CreateParameter cmd, "@p_DeliverableRootId", adInteger, adParamInput, 0, deliverableRootId
    dw.CreateParameter cmd, "@p_Status", adTinyInt, adParamInput, 0, ""
    dw.CreateParameter cmd, "@p_Comments", adVarChar, adParamInput, 2000, spareKitComments
    dw.CreateParameter cmd, "@p_Notes", adVarChar, adParamInput, 2000, spareKitNotes
    dw.CreateParameter cmd, "@p_FirstServiceDt", adDate, adParamInput, 0, spareKitFirstServiceDt
    dw.CreateParameter cmd, "@p_LastUpdUser", adVarchar, adParamInput, 100, m_UserFullName
    
    Dim retVal
    retVal = dw.ExecuteNonQuery(cmd)
    
    Response.Write "retVal: " & retVal
        
    If retVal >= -1 Then response.write "<script type='text/javascript'>window.close();</script>"

End Sub



Sub DisplaySpareKit(blnReadOnly)

    Dim strElemDisabled: strElemDisabled=""
    Dim strEditLinkText: strEditLinkText="Edit"

    If (blnReadOnly) Then
	strElemDisabled="disabled='disabled'"
	strEditLinkText=""
    End If

    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim PVID : PVID = regEx.Replace(Request.QueryString("PVID"), "")
    Dim DRID : DRID = regEx.Replace(Request.QueryString("DRID"), "")
    Dim SKID : SKID = regEx.Replace(Request.QueryString("SKID"), "")
    Dim CID : CID = regEx.Replace(Request.QueryString("CID"), "")
    regEx.Pattern = "[^0-9-]"
    Dim SFPN : SFPN = trim(Request.QueryString("SFPN"))
    
    If CID = "" Then CID = 0
    
    Dim rs, dw, cn, cmd, strSql
    Dim categoryOptions
    Dim csrLevelOptions
    Dim dispositionOptions
    Dim warrantyOptions
    Dim existingKitOptions
    Dim stockAdviceOptions
    Dim currentKitCategoryId : currentKitCategoryId = CID
    Dim currentKitNo
    Dim currentKitDescription
    Dim currentKitNotes
    Dim currentKitComments
    Dim currentKitCsrLevel
    Dim currentKitFirstServiceDt
    Dim currentMfgSubAssy
    Dim currentSvcSubAssy
    Dim fsDate
    Dim ppc : ppc = false
    Dim ppcCsr : ppcCsr = false
    Dim gpg : gpg = false
    
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

'
' Get Current Kit Info
'
    If SKID <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "usp_SelectSpareKitDetails")
        dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 0, CLng(SKID)
        dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, ""
        dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, SFPN
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        If Not rs.Eof Then
            currentKitNo = rs("SpareKitNo") & ""
            currentKitDescription = rs("Description") & ""
            If(Len(Trim(currentKitDescription))>40) Then 
                currentKitDescription=mid(Trim(currentKitDescription),1,40)
            End If
            currentKitCategoryId = rs("SpareCategoryId") & ""
            currentKitNotes = rs("Notes") & ""
            currentKitComments = rs("Comments") & ""
            currentKitFirstServiceDt = rs("FirstServiceDt") & ""
            currentMfgSubAssy = rs("MfgSubAssembly") & ""
            currentSvcSubAssy = rs("SvcSubAssembly") & ""
            ppc = LCase(rs("ppc") & "")
            ppcCsr = LCase(rs("PpcCsrLevelSet") & "")
            gpg = LCase(rs("gpg") & "")
        End If
        rs.close
    End If

'
' Get Category List
'
    Set cmd = dw.CreateCommandSP(cn, "usp_ListServiceSpareCategories")
    Set rs = dw.ExecuteCommandReturnRS(cmd)
    Do Until rs.Eof

        If CLng(currentKitCategoryId) = CLng(rs("ID")) Then
            categoryOptions = categoryOptions & "<option value=""" & rs("ID") & """ selected>" & rs("CategoryName") & "</option>"
        Else
            categoryOptions = categoryOptions & "<option value=""" & rs("ID") & """>" & rs("CategoryName") & "</option>"
        End If

        rs.MoveNext

    Loop
    rs.close
    
'
' Get CSR Level List
'

    Set cmd = dw.CreateCommandSP(cn, "usp_ListServiceCsrLevels")
    Set rs = dw.ExecuteCommandReturnRS(cmd)
    Do Until rs.Eof
        If CLng(currentKitCsrLevel) = CLng(rs("ID")) Then
	    csrLevelOptions = csrLevelOptions & "<option value=""" & rs("ID") & """ selected>" & rs("CsrLevel") & " - " & rs("CsrDescription") & "</option>"
        Else
            csrLevelOptions = csrLevelOptions & "<option value=""" & rs("ID") & """>" & rs("CsrLevel") & " - " & rs("CsrDescription") & "</option>"
        End If
        rs.MoveNext
    Loop
    rs.close    
    
'
' Get Disposition List
' 
    dispositionOptions = ""
    dispositionOptions = dispositionOptions & "<option value=""1"">1 - Disposable</option>"
    dispositionOptions = dispositionOptions & "<option value=""2"">2 - Repairable</option>"
    dispositionOptions = dispositionOptions & "<option value=""3"">3 - Return to Vendor</option>"
    dispositionOptions = dispositionOptions & "<option value=""4"">4 - Repair/Exhcange Only</option>"

'
' Get Warranty Tier List
' 
    warrantyOptions = ""
    warrantyOptions = warrantyOptions & "<option value=""A"">A - 0 - 5 Minutes</option>"
    warrantyOptions = warrantyOptions & "<option value=""B"">B - 5 - 10 Minutes</option>"
    warrantyOptions = warrantyOptions & "<option value=""C"">C - 15 - 30 Minutes</option>"
    warrantyOptions = warrantyOptions & "<option value=""D"">D - No Reimbursment</option>"

'
' Get Stock Advice List
' 
    stockAdviceOptions = ""
    stockAdviceOptions = stockAdviceOptions & "<option value=""1"">1 - Don't stock local (non SPOF and not likely to fail)</option>"
    stockAdviceOptions = stockAdviceOptions & "<option value=""2"">2 - Stock Strategically (non SPOF and likely to fail)</option>"
    stockAdviceOptions = stockAdviceOptions & "<option value=""3"">3 - Stock Local (SPOF and not likely to fail)</option>"
    stockAdviceOptions = stockAdviceOptions & "<option value=""4"">4 - Stock Local Critical (SPOF and likely to fail)</option>"


'
' Get Existing Spare Kits
'
    existingKitOptions = ""
    If SKID = "" And PVID <> "" And DRID <> "" Then
        Set cmd = dw.CreateCommandSP(cn, "usp_SelectSpareKitsForRoot")
        dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, CLng(PVID)
        dw.CreateParameter cmd, "@p_DeliverableRootId", adInteger, adParamInput, 0, CLng(DRID)
        Set rs = dw.ExecuteCommandReturnRS(cmd)

        Do Until rs.Eof
            existingKitOptions = existingKitOptions & "<input " & strElemDisabled & " type=""radio"" name=""rbSpareKits"" value=""" & rs("PartNumber") & "|" & rs("GpgDescription") & """ />" & rs("PartNumber") & "&nbsp;-&nbsp;" & rs("GpgDescription") & "<br />"
            rs.MoveNext
        Loop
        rs.close
    End If
    
    '
    ' Get First Service Date
    '
    fsDate = ""
    If PVID <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "usp_GetRSLFSDDefault")
        dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, CLng(PVID)
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        
        If Not rs.Eof Then
            fsDate = rs("Summary_Dt")
            If IsDate(fsDate) Then
                'fsDate = DateAdd("d", 15, fsDate)
		fsDate = DateAdd("d", 14, fsDate)
            End If
	    rs.close
        End If
        
        If currentKitFirstServiceDt = "" Then currentKitFirstServiceDt = fsDate
        
    End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Untitled Page</title>
    <style type="text/css">
        body
        {
            font: xx-small verdana;
        }
        legend
        {
            font: bold x-small verdana;
            color: #004874;
        }
        .floatRight
        {
        	position: absolute;
            left: 250px;
            float: left;
            text-align: left;
        }
        .floatLeft
        {
            float: left;
            font: bold x-small verdana;
        }
        .floatText
        {
            position: absolute;
            left: 250px;
        }
        .display
        {
            text-align: left;
        }
        .link
        {
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .link:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .linkEditFloatRight
        {
            float: right;
            text-align: right;
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .linkEditFloatRight:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .inputBox
        {
            width: 350px;
        }
        .BomTable
        {
            width: 100%;
            border-bottom: solid 1px black;
            border-collapse: collapse;
        }
        .BomTable th
        {
            background: #004874;
            color: #ffffff;
        }
        .BomTable td
        {
            border-bottom: solid 1px black;
        }
    </style>

    <script type="text/javascript" src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js"></script>

    <script type="text/javascript">
        String.prototype.trim = function() {
            return this.replace(/^\s+|\s+$/g, "");
        }

        String.prototype.ltrim = function() {
            return this.replace(/^\s+/, "");
        }

        String.prototype.rtrim = function() {
            return this.replace(/\s+$/, "");
        }


        function body_onload() {
            var skid = document.getElementById("skid");
            var spsPartNo = document.getElementById("spsPartNo");
            var existingKits = document.getElementById("ExistingKits");
            var radioButtons = document.getElementsByName("rbSpareKits");


            var i = 0;
            ShowForDisplay(false);
            if (skid.value != '') {
                ShowForDisplay(true);
                existingKits.style.display = "none";
                GetKitDetails();
                GetBomDetails(spsPartNo.value);
            }

	    <% IF Not blnIsOSSPUser THEN %>
            //
            // Setup Event Handlers
            //
            document.getElementById("spsPartNo").onblur = spsPartNo_onblur;
            document.getElementById("spsPartNo").onfocus = spsPartNo_onfocus;
            document.getElementById("spsPartNoEditLink").onclick = EditLink_onclick;
            document.getElementById("spsDescriptionEditLink").onclick = EditLink_onclick;
            document.getElementById("spsCategoryEditLink").onclick = EditLink_onclick;
            document.getElementById("spsCsrLevelEditLink").onclick = EditLink_onclick;
            document.getElementById("spsDispositionEditLink").onclick = EditLink_onclick;
            document.getElementById("spsWarrantyEditLink").onclick = EditLink_onclick;
            document.getElementById("spsLocalStockAdviceEditLink").onclick = EditLink_onclick;
            document.getElementById("spsFirstServiceDtEditLink").onclick = EditLink_onclick;
            document.getElementById("spsMfgSubAssyEditLink").onclick = EditLink_onclick;
            document.getElementById("spsSvcSubAssyEditLink").onclick = EditLink_onclick;
	    <% END IF %>
            
            for (i = 0; i < radioButtons.length; i++) {
                radioButtons[i].onclick = rbSpareKits_onclick;

                if (radioButtons[i].value == "new" && skid.value == "")
                    radioButtons[i].checked = true;
            }

        }


        function UpdateDisplayValues() {
            var spsPartNo = document.getElementById("spsPartNo");
            var spsPartNoDisplay = document.getElementById("spsPartNoDisplay");
            var spsCategory = document.getElementById("spsCategory");
            var spsCategoryDisplay = document.getElementById("spsCategoryDisplay");
            var spsDescription = document.getElementById("spsDescription");
            var spsDescriptionDisplay = document.getElementById("spsDescriptionDisplay");
            var spsCsrLevel = document.getElementById("spsCsrLevel");
            var spsCsrLevelDisplay = document.getElementById("spsCsrLevelDisplay");
            var spsDisposition = document.getElementById("spsDisposition");
            var spsDispositionDisplay = document.getElementById("spsDispositionDisplay");
            var spsWarranty = document.getElementById("spsWarranty");
            var spsWarrantyDisplay = document.getElementById("spsWarrantyDisplay");
            var spsLocalStockAdvice = document.getElementById("spsLocalStockAdvice");
            var spsLocalStockAdviceDisplay = document.getElementById("spsLocalStockAdviceDisplay");
            var spsFirstServiceDt = document.getElementById("spsFirstServiceDt");
            var spsFirstServiceDtDisplay = document.getElementById("spsFirstServiceDtDisplay");
            var spsWarranty = document.getElementById("spsWarranty");
            var spsWarrantyDisplay = document.getElementById("spsWarrantyDisplay");
            var spsDescriptionLabel = document.getElementById("spsDescriptionLabel");

            spsPartNoDisplay.innerHTML = spsPartNo.value;

            spsDescriptionDisplay.innerHTML = spsDescription.value;
            spsDescriptionLabel.innerHTML = "Description:";

            spsCategoryDisplay.innerHTML = spsCategory.options[spsCategory.selectedIndex].text;
            spsCsrLevelDisplay.innerHTML = spsCsrLevel.options[spsCsrLevel.selectedIndex].text;
            spsDispositionDisplay.innerHTML = spsDisposition.options[spsDisposition.selectedIndex].text;
            spsWarrantyDisplay.innerHTML = spsWarranty.options[spsWarranty.selectedIndex].text;
            spsLocalStockAdviceDisplay.innerHTML = spsLocalStockAdvice.options[spsLocalStockAdvice.selectedIndex].text;
            spsFirstServiceDtDisplay.innerHTML = spsFirstServiceDt.value;
        }


        function EditLink_onclick() {
            var sender = window.event.srcElement.id.replace("EditLink", "");

            if (sender == "spsDescription") {
                var elementLabel = document.getElementById("spsDescriptionLabel");
                elementLabel.innerHTML = "Description (40 characters max.):";
            }

            var elementDisplay = document.getElementById(sender + "Display");
            var elementEditLink = document.getElementById(sender + "EditLink");
            var elementEdit = document.getElementById(sender + "Edit");

            elementDisplay.style.display = "none";
            elementEditLink.style.display = "none";
            elementEdit.style.display = "";
        }


        function ShowDescriptionForDisplay(value) {
            var spsDescriptionLabel = document.getElementById("spsDescriptionLabel");
            var spsDescriptionEditLink = document.getElementById("spsDescriptionEditLink");
            var spsDescriptionDisplay = document.getElementById("spsDescriptionDisplay");
            var spsDescriptionEdit = document.getElementById("spsDescriptionEdit");
            var bGpg = document.getElementById("gpg").value == "false";

            if (value) {
                spsDescriptionDisplay.style.display = "";
                spsDescriptionEdit.style.display = "none";
                if (bGpg)
                    spsDescriptionEditLink.style.display = "";
            }
            else {
                spsDescriptionLabel.innerHTML = "Description (40 characters max.):";
                spsDescriptionDisplay.style.display = "none";
                spsDescriptionEdit.style.display = "";
                spsDescriptionEditLink.style.display = "none";
            }
        }

        function ShowMfgSubAssyForDisplay(value) {
            var spsMfgSubAssyEditLink = document.getElementById("spsMfgSubAssyEditLink");
            var spsMfgSubAssyDisplay = document.getElementById("spsMfgSubAssyDisplay");
            var spsMfgSubAssyEdit = document.getElementById("spsMfgSubAssyEdit");
            var spsMfgSubAssy = document.getElementById("spsMfgSubAssy");
            var bGpg = document.getElementById("gpg").value == "false";

            if ((value) && (spsMfgSubAssy.value != "")) {
                spsMfgSubAssyDisplay.style.display = "";
                spsMfgSubAssyEdit.style.display = "none";
                spsMfgSubAssyEditLink.style.display = "";
            }
            else {
                spsMfgSubAssyDisplay.style.display = "none";
                spsMfgSubAssyEdit.style.display = "";
                spsMfgSubAssyEditLink.style.display = "none";
            }
        }

        function ShowSvcSubAssyForDisplay(value) {
            var spsSvcSubAssyEditLink = document.getElementById("spsSvcSubAssyEditLink");
            var spsSvcSubAssyDisplay = document.getElementById("spsSvcSubAssyDisplay");
            var spsSvcSubAssyEdit = document.getElementById("spsSvcSubAssyEdit");
            var spsSvcSubAssy = document.getElementById("spsSvcSubAssy");

            if ((value) && (spsSvcSubAssy.value != "")) {
                spsSvcSubAssyDisplay.style.display = "";
                spsSvcSubAssyEdit.style.display = "none";
                spsSvcSubAssyEditLink.style.display = "";
            }
            else {
                spsSvcSubAssyDisplay.style.display = "none";
                spsSvcSubAssyEdit.style.display = "";
                spsSvcSubAssyEditLink.style.display = "none";
            }
        }
        
        function ShowForDisplay(value) {
            var spsPartNoEditLink = document.getElementById("spsPartNoEditLink");
            var spsPartNoDisplay = document.getElementById("spsPartNoDisplay");
            var spsPartNoEdit = document.getElementById("spsPartNoEdit");
            var spsCategoryEditLink = document.getElementById("spsCategoryEditLink");
            var spsCategoryDisplay = document.getElementById("spsCategoryDisplay");
            var spsCategoryEdit = document.getElementById("spsCategoryEdit");
            var spsCsrLevelEditLink = document.getElementById("spsCsrLevelEditLink");
            var spsCsrLevelDisplay = document.getElementById("spsCsrLevelDisplay");
            var spsCsrLevelEdit = document.getElementById("spsCsrLevelEdit");
            var spsDispositionEditLink = document.getElementById("spsDispositionEditLink");
            var spsDispositionDisplay = document.getElementById("spsDispositionDisplay");
            var spsDispositionEdit = document.getElementById("spsDispositionEdit");
            var spsWarrantyEditLink = document.getElementById("spsWarrantyEditLink");
            var spsWarrantyDisplay = document.getElementById("spsWarrantyDisplay");
            var spsWarrantyEdit = document.getElementById("spsWarrantyEdit");
            var spsLocalStockAdviceEditLink = document.getElementById("spsLocalStockAdviceEditLink");
            var spsLocalStockAdviceDisplay = document.getElementById("spsLocalStockAdviceDisplay");
            var spsLocalStockAdviceEdit = document.getElementById("spsLocalStockAdviceEdit");
            var spsFirstServiceDtEditLink = document.getElementById("spsFirstServiceDtEditLink");
            var spsFirstServiceDtDisplay = document.getElementById("spsFirstServiceDtDisplay");
            var spsFirstServiceDtEdit = document.getElementById("spsFirstServiceDtEdit");
            var spsWarrantyEditLink = document.getElementById("spsWarrantyEditLink");
            var spsWarrantyDisplay = document.getElementById("spsWarrantyDisplay");
            var spsWarrantyEdit = document.getElementById("spsWarrantyEdit");
            var bGpg = document.getElementById("gpg").value == "false";
            var bPpc = document.getElementById("ppc").value == "false";
            var bPpcCsr = document.getElementById("ppcCsr").value == "false";
            var bSpsPartNoEmpty = document.getElementById("spsPartNo").value == "";

            ShowDescriptionForDisplay(value);
            ShowMfgSubAssyForDisplay(value);
            ShowSvcSubAssyForDisplay(value);
            if (value) {
                spsCategoryDisplay.style.display = "";
                spsCategoryEdit.style.display = "none";
                spsCategoryEditLink.style.display = "";

                spsCsrLevelDisplay.style.display = "";
                spsCsrLevelEdit.style.display = "none";
                if (bPpcCsr)
                    spsCsrLevelEditLink.style.display = "";

                if (!bSpsPartNoEmpty) {
                    spsPartNoDisplay.style.display = "";
                    spsPartNoEdit.style.display = "none";
                    //spsPartNoEditLink.style.display = "";
                }

                spsDispositionDisplay.style.display = "";
                spsDispositionEdit.style.display = "none";
                spsDispositionEditLink.style.display = "";

                spsWarrantyDisplay.style.display = "";
                spsWarrantyEdit.style.display = "none";
                spsWarrantyEditLink.style.display = "";

                spsLocalStockAdviceEdit.style.display = "none";
                spsLocalStockAdviceDisplay.style.display = "";
                if (bPpc)
                    spsLocalStockAdviceEditLink.style.display = "";

                spsFirstServiceDtEdit.style.display = "none";
                spsFirstServiceDtDisplay.style.display = "";
                spsFirstServiceDtEditLink.style.display = "";
                
                UpdateDisplayValues();
            } else {
                spsCategoryDisplay.style.display = "none";
                spsCategoryEdit.style.display = "";
                spsCategoryEditLink.style.display = "none";
                
                spsCsrLevelDisplay.style.display = "none";
                spsCsrLevelEdit.style.display = "";
                spsCsrLevelEditLink.style.display = "none";
                
                spsPartNoDisplay.style.display = "none";
                spsPartNoEdit.style.display = "";
                spsPartNoEditLink.style.display = "none";
                
                spsDispositionDisplay.style.display = "none";
                spsDispositionEdit.style.display = "";
                spsDispositionEditLink.style.display = "none";
                
                spsWarrantyDisplay.style.display = "none";
                spsWarrantyEdit.style.display = "";
                spsWarrantyEditLink.style.display = "none";
                
                spsLocalStockAdviceEdit.style.display = "";
                spsLocalStockAdviceDisplay.style.display = "none";
                spsLocalStockAdviceEditLink.style.display = "none";
                
                spsFirstServiceDtEdit.style.display = "";
                spsFirstServiceDtDisplay.style.display = "none";
                spsFirstServiceDtEditLink.style.display = "none";
                
            }
        }

        var spsPartNoValue = "";
        function spsPartNo_onfocus() {
            var spsPartNo = document.getElementById("spsPartNo");
            spsPartNoValue = spsPartNo.value;
        }

        function spsPartNo_onblur() {
            var spsPartNo = document.getElementById("spsPartNo");
            if (spsPartNoValue != spsPartNo.value) {
                GetKitDetails();
                GetBomDetails();
            }
        }

        function rbSpareKits_onclick(sender) {
            var spsPartNo = document.getElementById("spsPartNo");
            var spsDescription = document.getElementById("spsDescription");
            var bomDetails = document.getElementById("BomDetails");
            var existingSearch = document.getElementById("existingSearch");
            var rbSender = window.event.srcElement;
            if (rbSender.value != "new" && rbSender.value != "bridge" && rbSender.value != "existing") {
                var arrKitNfo = rbSender.value.split("|")
                spsPartNo.value = arrKitNfo[0];
                spsDescription.value = arrKitNfo[1];
                GetKitDetails()
                GetBomDetails();
            }

            if (rbSender.value == "new") {
                spsPartNo.value = "";
                spsDescription.value = "";
                ShowDescriptionForDisplay(false);
                bomDetails.innerHTML = "&nbsp;";

            }

            if (rbSender.value == "existing") {
                existingSearch.style.display = "";
            }
            else {
                existingSearch.style.display = "none";
            }
        }

        function GetKitDetails() {
            var spsPartNo = document.getElementById("spsPartNo");
            var spsPartId = document.getElementById("skid");
            var serviceFamilyPn = document.getElementById("sfpn");
            if (spsPartNo != "")
                jsrsExecute("<%=AppRoot %>/Service/rsService.asp", GetKitDetailsCallBack, "GetKitDetails", Array(String(spsPartNo.value.trim()), String(spsPartId.value.trim()), String(serviceFamilyPn.value.trim())));
        }

        function GetKitDetailsCallBack(result) {
            var spsGeosNa = document.getElementById("spsGeosNa");
            var spsGeosLa = document.getElementById("spsGeosLa");
            var spsGeosApj = document.getElementById("spsGeosApj");
            var spsGeosEmea = document.getElementById("spsGeosEmea");
            var spsNotes = document.getElementById("spsNotes");
            var spsComments = document.getElementById("spsComments");
            var spsPartNo = document.getElementById("spsPartNo");
            var spsCategory = document.getElementById("spsCategory");
            var spsDescription = document.getElementById("spsDescription");
            var spsCsrLevel = document.getElementById("spsCsrLevel");
            var spsDisposition = document.getElementById("spsDisposition");
            var spsWarranty = document.getElementById("spsWarranty");
            var spsLocalStockAdvice = document.getElementById("spsLocalStockAdvice");
            var spsFirstServiceDt = document.getElementById("spsFirstServiceDt");
            var spsWarranty = document.getElementById("spsWarranty");
            var spsMfgSubAssy = document.getElementById("spsMfgSubAssy");
            var spsSvcSubAssy = document.getElementById("spsSvcSubAssy");

            //alert(result);

            if ((result != "") && (result != "No Rows")) {
                var arrResult = result.split("|");
                spsDescription.value = arrResult[2];
                ShowDescriptionForDisplay(spsDescription.value.trim() != "");

                if (arrResult[3] != '')
                    spsCategory.value = arrResult[3];
                if (arrResult[4] != '')
                    spsCsrLevel.value = arrResult[4];
                if (arrResult[5] != '')
                    spsDisposition.value = arrResult[5];
                if (arrResult[6] != '')
                    spsWarranty.value = arrResult[6];
                if (arrResult[10] != '')
                    spsGeosNa.checked = (arrResult[10] == 'True') ? true : false;
                if (arrResult[11] != '')
                    spsGeosLa.checked = (arrResult[11] == 'True') ? true : false;
                if (arrResult[12] != '')
                    spsGeosApj.checked = (arrResult[12] == 'True') ? true : false;
                if (arrResult[13] != '')
                    spsGeosEmea.checked = (arrResult[13] == 'True') ? true : false;
                spsNotes.value = arrResult[7];
                spsComments.value = arrResult[8];
                if (arrResult[14] != '')
                    spsLocalStockAdvice.value = arrResult[14];
                if (arrResult[9] != '')
                    spsFirstServiceDt.value = arrResult[9];
                spsMfgSubAssy.value = arrResult[15];
                spsSvcSubAssy.value = arrResult[16];

                UpdateDisplayValues();
            }
            else {
                spsDescription.value = "";
                ShowDescriptionForDisplay(false);

                //spsCategory.value = 0;
                spsCsrLevel.selectedIndex = 0;
                spsDisposition.selectedIndex = 0;
                spsWarranty.selectedIndex = 0;
                spsGeosNa.checked = false;
                spsGeosLa.checked = false;
                spsGeosApj.checked = false;
                spsGeosEmea.checked = false;
                spsNotes.value = "";
                spsComments.value = "";
                spsLocalStockAdvice.selectedIndex = 0;
                //spsFirstServiceDt.value = "";

                UpdateDisplayValues();
            }
        }

        function GetBomDetails() {
            var spsPartNo = document.getElementById("spsPartNo");
            var bomDetails = document.getElementById("BomDetails");
            bomDetails.innerHTML = '<span style="color:#004874;font:bold x-small;"><img src="/images/loading24.gif" /> Loading ...</span>';

            jsrsExecute("<%=AppRoot %>/Service/rsService.asp", GetBomDetailsCallBack, "GetBomDetails", String(spsPartNo.value));
        }

        function GetBomDetailsCallBack(result) {
            var bomDetails = document.getElementById("BomDetails");
            bomDetails.innerHTML = result;
        }

        function FindSpareKit_onclick() {
            var spsPartNo = document.getElementById("spsPartNo");
            var spsDescription = document.getElementById("spsDescription");
            var spsCategory = document.getElementById("spsCategory");
            var pvid = document.getElementById("pvid");
            var strResult;
            strResult = window.showModalDialog("FindSpareKit.asp?category=" + spsCategory.value + "&PVID=" + pvid.value, "", "dialogWidth:600px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");

            if (typeof (strResult) != "undefined") {
                spsPartNo.value = strResult;
                if (strResult.indexOf("|") > 0) {
                    var strResultArray = strResult.split("|");
                    spsPartNo.value = strResultArray[0];
                    spsDescription.value = strResultArray[1];
                }
                GetBomDetails();
                GetKitDetails();
            }
        }


        function cmdDate_onclick(target) {
            var strID;
            var txtDateField = document.getElementById(target);
            strID = window.showModalDialog("../MobileSE/Today/calDraw1.asp", txtDateField.value, "dialogWidth:300px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                txtDateField.value = strID;
            }
        }


    </script>

</head>

<body onload="body_onload()">
<form id="frmMain" method="post">
    <div style="width: 600px">
        <div id="KitDetails">
            <fieldset>
                <legend>Spare Kit - <%= m_UserFullName %></legend>
                <div style="height: 25px">
                    <span class="floatLeft">Part No:</span><span id="spsPartNoEdit" class="floatRight">
                        <input class="inputBox" style="width: 100px;" type="text" id="spsPartNo" name="spsPartNo" value="<%= currentKitNo %>" <%=strElemDisabled%>/></span>
                    <span id="spsPartNoDisplay" class="floatText">
                        <%= currentKitNo %></span><span id="spsPartNoEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height: 25px">
                    <span id="spsDescriptionLabel" class="floatLeft">Description:</span><span id="spsDescriptionEdit" class="floatRight">
                        <input class="inputBox" type="text" id="spsDescription" name="spsDescription" value="<%= currentKitDescription %>" <%=strElemDisabled%>/></span>
                    <span id="spsDescriptionDisplay" class="floatText">
                        <%= currentKitDescription %></span><span id="spsDescriptionEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height: 25px">
                    <span class="floatLeft">Part Type:</span><span id="spsCategoryEdit" class="floatRight">
                        <select class="inputBox" id="spsCategory" name="spsCategory" <%=strElemDisabled%>>
                            <option value="0">-- Select Category --</option>
                            <%= categoryOptions %>
                        </select></span> <span id="spsCategoryDisplay" class="floatText"></span><span id="spsCategoryEditLink"
                            class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height: 25px">
                    <span class="floatLeft">CSR Level:</span>
                    <span id="spsCsrLevelEdit" class="floatRight"><select class="inputBox" id="spsCsrLevel" name="spsCsrLevel" <%=strElemDisabled%>>
                            <option value="0">-- Select CSR Level --</option>
                            <%= csrLevelOptions %>
                        </select></span> 
                    <span id="spsCsrLevelDisplay" class="floatText"></span>
                    <span id="spsCsrLevelEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height:25px">
                    <span class="floatLeft">Disposition:</span>
                    <span id="spsDispositionEdit" class="floatRight"><select class="inputBox" id="spsDisposition" name="spsDisposition" <%=strElemDisabled%>>
                        <option value="0">-- Select Disposition --</option>
                        <%= dispositionOptions %>
                        </select></span>
                    <span id="spsDispositionDisplay" class="floatText"></span>
                    <span id="spsDispositionEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height:25px">
                    <span class="floatLeft">Warranty Labour Tier:</span>
                    <span id="spsWarrantyEdit" class="floatRight"><select class="inputBox" id="spsWarranty" name="spsWarranty" <%=strElemDisabled%>>
                        <option value="0">-- Select Warranty Labour Tier --</option>
                        <%= warrantyOptions %>
                        </select></span>
                    <span id="spsWarrantyDisplay" class="floatText"></span>
                    <span id="spsWarrantyEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height:25px">
                    <span class="floatLeft">Local Stock Advice:</span>
                    <span id="spsLocalStockAdviceEdit" class="floatRight"><select class="inputBox" id="spsLocalStockAdvice" name="spsLocalStockAdvice" <%=strElemDisabled%>>
                        <option value="0">-- Select Stock Advice --</option>
                        <%= stockAdviceOptions %>
                        </select></span>
                    <span id="spsLocalStockAdviceDisplay" class="floatText"></span>
                    <span id="spsLocalStockAdviceEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height:25px">
                    <span class="floatLeft">GEOS:</span>
                    <span id="spsGeosEdit" class="floatText">
                        <input type="checkbox" id="spsGeosNa" name="spsGeosNa" <%=strElemDisabled%>/> NA
                        <input type="checkbox" id="spsGeosLa" name="spsGeosLa" <%=strElemDisabled%>/> LA
                        <input type="checkbox" id="spsGeosApj" name="spsGeosApj" <%=strElemDisabled%>/> APJ
                        <input type="checkbox" id="spsGeosEmea" name="spsGeosEmea" <%=strElemDisabled%>/> EMEA
                    </span>
		</div>
                <div style="height: 25px">
                    <span class="floatLeft">First Service Dt.:</span><span id="spsFirstServiceDtEdit" class="floatRight">
                        <input class="inputBox" style="width:300px" type="text" id="spsFirstServiceDt" name="spsFirstServiceDt" value="<%= currentKitFirstServiceDt %>" <%=strElemDisabled%>/>
			<% If Not blnIsOSSPUser Then %>
                        <img src="../images/calendar.gif" alt="Calendar" onclick="cmdDate_onclick('spsFirstServiceDt')" />
			<% End If %>
			</span>
                    <span id="spsFirstServiceDtDisplay" class="floatText">
                        <%= currentKitNo %></span><span id="spsFirstServiceDtEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>
                <div style="height: 25px">
                    <span class="floatLeft">Manufacturing Sub Assembly:</span><span id="spsMfgSubAssyEdit" class="floatRight">
                        <input class="inputBox" style="width: 100px;" type="text" id="spsMfgSubAssy" name="spsMfgSubAssy" value="<%= currentMfgSubAssy %>" <%=strElemDisabled%>/></span>
                    <span id="spsMfgSubAssyDisplay" class="floatText">
                        <%= currentMfgSubAssy %></span><span id="spsMfgSubAssyEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>

                <div style="height: 25px;display:;">
                    <span class="floatLeft">Service Sub Assembly:</span><span id="spsSvcSubAssyEdit" class="floatRight">
                        <input class="inputBox" style="width: 100px;" type="text" id="spsSvcSubAssy" name="spsSvcSubAssy" value="<%= currentSvcSubAssy %>" <%=strElemDisabled%>/></span>
                    <span id="spsSvcSubAssyDisplay" class="floatText">
                        <%= currentSvcSubAssy %></span><span id="spsSvcSubAssyEditLink" class="linkEditFloatRight"><%=strEditLinkText%></span></div>

                <div style="height: 75px">
                    <span class="floatLeft">Internal Notes:</span><span class="floatRight"><textarea class="inputBox"
                        rows="4" id="spsNotes" name="spsNotes" <%=strElemDisabled%>><%= currentKitNotes %></textarea></span></div>
                <div style="height: 75px">
                    <span class="floatLeft">RSL Comments:</span><span class="floatRight"><textarea class="inputBox"
                        rows="4" id="spsComments" name="spsComments" <%=strElemDisabled%>><%= currentKitComments %></textarea></span></div>
            </fieldset>
            <br />
        </div>
        <div id="ExistingKits">
            <fieldset>
                <legend>Existing Spare Kits</legend>
                <%= existingKitOptions %>
                <input type="radio" name="rbSpareKits" value="new" <%=strElemDisabled%>/>Create New Spare Kit<br />
                <input type="radio" name="rbSpareKits" value="existing" <%=strElemDisabled%>/>Use Existing Kit <span
                    id="existingSearch" class="link" style="display:none;" onclick="FindSpareKit_onclick()">Find Kit</span><br />
                <!--<input type="radio" name="rbSpareKits" value="bridge" />Bridge To Another Kit <span id="bridgeSearch" style="display: none; font: bold x-small verdana;
                    color: #004874; text-decoration: underline;" onclick="FindSpareKit_onclick()" onmouseover="Link_mouseover()" onmouseout="Link_mouseout()">Find Kit</span>-->

            </fieldset>
            <br />
        </div>
        <div id="BomInfo">
            <fieldset>
                <legend>BOM Info</legend><span id="BomDetails">&nbsp;</span>
            </fieldset>
        </div>
    </div>



    <input type="hidden" id="sfpn" name="sfpn" value="<%= SFPN %>" />
    <input type="hidden" id="pvid" name="pvid" value="<%= PVID %>" />
    <input type="hidden" id="drid" name="drid" value="<%= DRID %>" />
    <input type="hidden" id="skid" name="skid" value="<%= SKID %>" />
    <input type="hidden" id="ppc" name="ppc" value="<%= ppc %>" />
    <input type="hidden" id="ppcCsr" name="ppcCsr" value="<%= ppcCsr %>" />
    <input type="hidden" id="gpg" name="gpg" value="<%= gpg %>" />
    <input type="hidden" id="action" name="action" />



</form>
</body>
</html>
<%
End Sub 'DisplaySpareKit
%>

<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->
<%

Dim rs, dw, cn, cmd
Dim BrandID, BrandDesc, BrandRadioButtons, _
    BUID, BUDesc, BUCheckBoxes, _
    CPUID, CPUDesc, CPUCheckBoxes, _
    FamilyID, FamilyName, OsFamilyRadioButtons

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sMode				: sMode = Request.QueryString("Mode")
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim sFeatureCat			: sFeatureCat = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sConfigRules		: sConfigRules = ""
Dim iBrandID			: iBrandID = ""
Dim sAvList				: sAvList = ""
Dim sErrors             : sErrors = ""
Dim sPreview            : sPreview = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_EditModeOn
Dim m_UserFullName


'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

'	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
'	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
'	If m_IsSysAdmin Then
		m_EditModeOn = True
'	End If
	
'	If Not m_EditModeOn Then
'		Response.Write "<H3>Insufficient User Privilidges</H3><H4>Access Denied</H4>"
'		Response.End
'	End If

	Set Security = Nothing
'##############################################################################	
'
'PC Can do any thing
'Marketing can change description

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Sub Main()
'
' TODO: Select Base Units based on PVID
'

    Dim lastBrandID, firstDiv

	Set cmd = dw.CreateCommandSP(cn, "usp_ListProductBrands")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	BrandRadioButtons = ""
	If Not rs.EOF Then
	    Do Until rs.EOF
		    BrandID = rs("ProductBrandID")&""
		    BrandDesc = rs("BrandFullName")&""
            BrandRadioButtons = BrandRadioButtons & "<input id=""rdoBrandID"" type=""radio"" name=""BrandID"" value=""" & BrandID & """ onclick=""showBrand(" & BrandID & ")""/><label>" & BrandDesc & "</label><br />"
            rs.MoveNext
        Loop
    Else
        BrandRadioButtons = "<label>No Brands Found</label>"
	End If
		
	rs.Close
	
	set cmd = dw.CreateCommandSP(cn, "usp_ListOsFamilies")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	OsFamilyRadioButtons = ""
	If Not rs.EOF Then
	    Do Until rs.EOF
	        FamilyID = rs("ID") & ""
	        FamilyName = rs("FamilyName") & ""
	        OsFamilyRadioButtons = OsFamilyRadioButtons & "<input id=""rdoOsFamilyID"" type = ""radio"" name=""OsFamilyID"" value=""" & FamilyID & """ onclick=""setOSID(" & FamilyID & ")""/><label>" & FamilyName & "</label><br />" 
	        rs.MoveNext
	    Loop
	Else
	    OsFamilyRadioButtons = "<label>No Operating Systems Found</label>"
	End If
		
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
   	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, "1"
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	BUCheckBoxes = ""
	lastBrandID = ""
	firstDiv = true
	If Not rs.EOF Then
	    rs.Sort = "ProductBrandID, GPGDescription"
	    Do Until rs.EOF
	        If lastBrandID <> rs("ProductBrandID") Then
	            lastBrandID = rs("ProductBrandID")
	            If NOT firstDiv Then
	                BUCheckBoxes = BUCheckBoxes & "</DIV>"
	            End If

                firstDiv = false
                BUCheckBoxes = BUCheckBoxes & "<DIV ID=""BU" & rs("ProductBrandID") & """ name=""BaseUnits"">"
	            
	        End If
		    BUID = rs("AvDetailID")&""
		    BUDesc = rs("GPGDescription")&""
            BUCheckBoxes = BUCheckBoxes & "<input id=""cbxBu" & BUID & """ value=""" & BUID & """ type=""checkbox"" /><label>" & BUDesc & "</label><br />"
            rs.MoveNext
        Loop
        
        BUCheckBoxes = BUCheckBoxes & "</DIV>"
    
    Else
        BUCheckBoxes = "<label>No Base Units Found</label>"
	End If
		
	rs.Close
		
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
   	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, "4"
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	CPUCheckBoxes = ""
	lastBrandID = ""
	firstDiv = true
	If Not rs.EOF Then
	    rs.Sort = "ProductBrandID, GPGDescription"
	    Do Until rs.EOF
	        If lastBrandID <> rs("ProductBrandID") Then
	            lastBrandID = rs("ProductBrandID")
	            If NOT firstDiv Then
	                CPUCheckBoxes = CPUCheckBoxes & "</DIV>"
	            End If
                firstDiv = false
                CPUCheckBoxes = CPUCheckBoxes & "<DIV ID=""CPU" & rs("ProductBrandID") & """ name=""CPU"">"
	            
	        End If
	    
		    CPUID = rs("AvDetailID")&""
		    CPUDesc = rs("GPGDescription")&""
            CPUCheckBoxes = CPUCheckBoxes & "<input id=""cbxCpu" & CPUID & """ type=""checkbox"" /><label>" & CPUDesc & "</label><br />"
            rs.MoveNext
        Loop
        
        CPUCheckBoxes = CPUCheckBoxes & "</DIV>"
    
    Else
        CPUCheckBoxes = "<label>No CPUs Found</label>"
	End If
		
	rs.Close


End Sub

Sub Preview()

    Dim sSubmissionInfo, sBuCpuOkay, sBuCpuError
    'Write out submission information
    
    'Check base unit cpu combos for validity
    sPreview = sPreview & _
        "<table class='FormTable' cellpadding='1' cellspacing='0'>" & _
        "<tr><th colspan='3'>The following records will be added.</th></tr>" & _
        "<tr><th>WHQL ID</th><th>OS Family</th><th>Base Unit</th><th>CPU</th></tr>"

    sPreview = sPreview & sBuCpuOkay
   
    sPreview = sPreview & _
        "</table>"
    
    sPreview = sPreview & _
        "<table class='FormTable' cellpadding='1' cellspacing='0'>" & _
        "<tr><th colspan='3'>The following records will not be added because they are duplicates.</th></tr>" & _
        "<tr><th>WHQL ID</th><th>OS Family</th><th>Base Unit</th><th>CPU</th></tr>"

    sPreview = sPreview & sBuCpuError
   
    sPreview = sPreview & _
        "</table>" 
        
    'Repopulate selections

    Dim lastBrandID, firstDiv

	Set cmd = dw.CreateCommandSP(cn, "usp_ListProductBrands")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	BrandRadioButtons = ""
	If Not rs.EOF Then
	    Do Until rs.EOF
		    BrandID = rs("ProductBrandID")&""
		    BrandDesc = rs("BrandFullName")&""
            BrandRadioButtons = BrandRadioButtons & "<input id=""rdoBrandID"" type=""radio"" name=""BrandID"" value=""" & BrandID & """ onclick=""showBrand(" & BrandID & ")""" 
            If BrandID = Request.Form("BrandID") Then
                BrandRadioButtons = BrandRadioButtons & " CHECKED "
            End If
            BrandRadioButtons = BrandRadioButtons & "/><label>" & BrandDesc & "</label><br />"
            rs.MoveNext
        Loop
        BrandRadioButtons = BrandRadioButtons & "<script type='text/javascript'>showBrand(" & Request.Form("BrandID") & ");</script>"
    Else
        BrandRadioButtons = "<label>No Brands Found</label>"
	End If
		
	rs.Close
	
	set cmd = dw.CreateCommandSP(cn, "usp_ListOsFamilies")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	OsFamilyRadioButtons = ""
	If Not rs.EOF Then
	    Do Until rs.EOF
	        FamilyID = rs("ID") & ""
	        FamilyName = rs("FamilyName") & ""
	        OsFamilyRadioButtons = OsFamilyRadioButtons & "<input id=""rdoOsFamilyID"" type = ""radio"" name=""OsFamilyID"" value=""" & FamilyID & """ onclick=""setOSID(" & FamilyID & ")""/><label>" & FamilyName & "</label><br />" 
	        rs.MoveNext
	    Loop
	Else
	    OsFamilyRadioButtons = "<label>No Operating Systems Found</label>"
	End If
		
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
   	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, "1"
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	BUCheckBoxes = ""
	lastBrandID = ""
	firstDiv = true
	If Not rs.EOF Then
	    rs.Sort = "ProductBrandID, GPGDescription"
	    Do Until rs.EOF
	        If lastBrandID <> rs("ProductBrandID") Then
	            lastBrandID = rs("ProductBrandID")
	            If NOT firstDiv Then
	                BUCheckBoxes = BUCheckBoxes & "</DIV>"
	            End If

                firstDiv = false
                BUCheckBoxes = BUCheckBoxes & "<DIV ID=""BU" & rs("ProductBrandID") & """ name=""BaseUnits"">"
	            
	        End If
		    BUID = rs("AvDetailID")&""
		    BUDesc = rs("GPGDescription")&""
            BUCheckBoxes = BUCheckBoxes & "<input id=""cbxBu" & BUID & """ value=""" & BUID & """ type=""checkbox"" /><label>" & BUDesc & "</label><br />"
            rs.MoveNext
        Loop
        
        BUCheckBoxes = BUCheckBoxes & "</DIV>"
    
    Else
        BUCheckBoxes = "<label>No Base Units Found</label>"
	End If
		
	rs.Close
		
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
   	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, "4"
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	CPUCheckBoxes = ""
	lastBrandID = ""
	firstDiv = true
	If Not rs.EOF Then
	    rs.Sort = "ProductBrandID, GPGDescription"
	    Do Until rs.EOF
	        If lastBrandID <> rs("ProductBrandID") Then
	            lastBrandID = rs("ProductBrandID")
	            If NOT firstDiv Then
	                CPUCheckBoxes = CPUCheckBoxes & "</DIV>"
	            End If
                firstDiv = false
                CPUCheckBoxes = CPUCheckBoxes & "<DIV ID=""CPU" & rs("ProductBrandID") & """ name=""CPU"">"
	            
	        End If
	    
		    CPUID = rs("AvDetailID")&""
		    CPUDesc = rs("GPGDescription")&""
            CPUCheckBoxes = CPUCheckBoxes & "<input id=""cbxCpu" & CPUID & """ type=""checkbox"" /><label>" & CPUDesc & "</label><br />"
            rs.MoveNext
        Loop
        
        CPUCheckBoxes = CPUCheckBoxes & "</DIV>"
    
    Else
        CPUCheckBoxes = "<label>No CPUs Found</label>"
	End If
		
	rs.Close



End Sub

Sub Save()
On Error Goto 0

	Dim returnValue
	Dim whqlID
	
	cn.BeginTrans

'
' Save WHQL Submission and get the WHQLID
'
	Set cmd = dw.CreateCommandSP(cn, "usp_InsertProductWHQL")
	cmd.NamedParameters = True
	dw.CreateParameter cmd, "@p_SubmissionID", adVarchar, adParamInput, 50, Request.Form("submissionID")
	dw.CreateParameter cmd, "@p_SubmissionDt", adDate, adParamInput, 8, Request.Form("submissionDt")
	dw.CreateParameter cmd, "@p_WhqlDt", adDate, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_Status", adInteger, adParamInput, 8, "1"
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_Location", adVarchar, adParamInput, 200, Request.Form("txtLabelLocation")
	dw.CreateParameter cmd, "@p_DateReleased", adDate, adParamInput, 8, Request.Form("txtDateReleased")
	dw.CreateParameter cmd, "@p_LogoDisplayed", adBoolean, adParamInput, 1, Request.Form("chkLogoDisplayed")
	dw.CreateParameter cmd, "@p_Milestone3", adBoolean, adParamInput, 1, Request.Form("chkMilestone3")
	dw.CreateParameter cmd, "@p_ProductWhqlID", adInteger, adParamOutput, 8, ""
	returnValue = dw.ExecuteNonQuery(cmd)
		
	whqlID = cmd("@p_ProductWhqlID")

	If returnValue = 0 Then
		' Abort Transaction
		Response.Write "Error while saving WHQL Details"
		cn.RollbackTrans()
		Exit Sub
	End If		
'
' Build the combined bu & cpu list and save to the WHQL_AvDetail table.
'
    Dim saBaseUnits, sBaseUnits
    Dim saCPUs, sCpus
    Dim iBU
    Dim iCPU
    Dim iOsFamilyID : iOsFamilyID = Request.Form("OSID")
    Dim iRecordsSaved : iRecordsSaved = 0
    
    sBaseUnits = Request.Form("hidBaseUnits")                
    sCpus = Request.Form("hidCpus")
    sBaseUnits = Mid(sBaseUnits,2,Len(sBaseUnits)-1)
    sCpus = Mid(sCpus,2,Len(sCpus)-1)
    saBaseUnits = split(sBaseUnits,",")
    saCPUs = split(sCpus,",")
    
    'response.Write request.Form
    'response.Write ubound(saBaseUnits) & "<br>"
    
    For iBU = 0 to ubound(saBaseUnits)
        'Response.Write "BU:" & saBaseUnits(iBU) & "<BR>"
        For iCPU = 0 to ubound(saCPUs)
            'Response.Write "CPU:" & saCPUs(iCPU) & "<BR>"
            'Insert WHQL_AvDetail Record
	        Set cmd = dw.CreateCommandSP(cn, "usp_InsertWhqlAvDetail")
	        cmd.NamedParameters = True
	        dw.CreateParameter cmd, "@p_WHQLID", adInteger, adParamInput, 8, whqlID
	        dw.CreateParameter cmd, "@p_BUID", adInteger, adParamInput, 8, saBaseUnits(iBU)
	        dw.CreateParameter cmd, "@p_CPUID", adInteger, adParamInput, 8, saCPUs(iCPU)
	        dw.CreateParameter cmd, "@p_OSFamilyID", adInteger, adParamInput, 8, iOsFamilyID
        	Set rs = dw.ExecuteCommandReturnRS(cmd)
        	
	        If rs.State > 0 Then
	            If Not (rs.EOF and rs.BOF) Then
                    sErrors = sErrors & "<TR><TD>" & rs("WHQL") & "</TD><TD>" & rs("OS") & "</TD><TD>" & rs("BU") & "</TD><TD>" & rs("CPU") & "</TD></TR>"
                End IF
	        Else
	            iRecordsSaved = iRecordsSaved + 1
	        End If
        Next
    Next
    
    response.Write iRecordsSaved
    If iRecordsSaved = 0 Then
        cn.RollbackTrans
    Else
        cn.CommitTrans
    End If
    
    set cmd = nothing
    set cn = nothing
    set dw = nothing
    
    sFunction = "close"

End Sub

Select Case LCase(sFunction)
	Case "save"
		Call Save()
	Case "preview"
	    Call Preview()
	Case Else
		Call Main()
End Select
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css">

    <script language="JavaScript" src="../includes/client/Common.js"></script>

    <script type="text/javascript">
function Body_OnLoad()
{
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		window.parent.frames["LowerWindow"].frmButtons.cmdFinish.disabled =true;
	}
	
	//hideAllBrands();
	//hideAllSteps();
	
    if (frmMain.hidFunction.value != "preview"){
        window.document.all["step1"].style.display = "";    
        }

	if (frmMain.hidFunction.value == "close"){
	    if (frmMain.hidErrors.value == ""){
            window.parent.close();
	    }	
	    else{
	        hideAllSteps()
	        window.document.all["errors"].style.display = "";
    	}
    }		


}

function hideAllSteps()
{
    for (element in window.document.getElementsByTagName("DIV"))
	    if (element.substr(0,4) == "step")
	        window.document.all[element].style.display = "none";
}

function hideAllBrands()
{
	for (element in window.document.getElementsByTagName("DIV"))
	{

	    if (element.substr(0,2) == "BU")
	        document.all[element].style.display = "none";
	        
	    if (element.substr(0,3) == "CPU")
	        document.all[element].style.display = "none";
	}
}

function showBrand(brandID)
{
    hideAllBrands();
    document.all['BU'+brandID].style.display = "";
    document.all['CPU'+brandID].style.display = "";
    frmMain.BID.value = brandID;
    
    for (element in document.getElementsByTagName("INPUT"))
        if (element.substr(0,3) == "cbx")
            document.all[element].checked = false;

}

function setOSID(OsFamilyID) {
    frmMain.OSID.value = OsFamilyID;
}

function cmdDate_onclick(FieldID) {
	var strID;
	var oldValue = window.document.all(FieldID).value;
		
	strID = window.showModalDialog("../mobilese/today/caldraw1.asp",FieldID,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) == "undefined")
		return
	
	window.document.all(FieldID).value = strID;
}

</script>

</head>
<body onload="Body_OnLoad()">
    <form method="post" id="frmMain" action="whqlIdDetail.asp">
    <input id="hidMode" name="hidMode" type="text" value="<%= LCase(sMode)%>" />
    <input id="hidFunction" name="hidFunction" type="text" value="<%= LCase(sFunction)%>" />
    <input id="BID" name="BID" type="text" value="<%= Request("BID") %>"/>
    <input id="OSID" name="OSID" type="text" value="<%= Request("OSID") %>" />
    <input id="PVID" name="PVID" type="text" value="<%= Request("PVID") %>" />
    <input id="hidErrors" name="hidErrors" type="text" value="<%= sErrors %>" />
    <input id="hidCurrentStep" type="text" value="1" />
    <input id="hidBaseUnits" name="hidBaseUnits" type="text" value="<%= Request.Form("hidBaseUnits") %>"/>
    <input id="hidCpus" name="hidCpus" type="text" value="<%= Request.Form("hidCpus") %>"/>
    <div id="preview">
    </div>
    <div id="errors" style="display: none;">
        <table class="FormTable" cellpadding="1" cellspacing="0">
            <tr>
                <th colspan="3">
                    <span style="color: Red; font-size: x-small">The following records were not added because
                        they are duplicates.</span></th>
            </tr>
            <tr>
                <th>
                    WHQL ID</th>
                <th>
                    OS Family</th>
                <th>
                    Base Unit</th>
                <th>
                    CPU</th>
            </tr>
            <%= sErrors %>
        </table>
    </div>
    <!-- Enter Submission ID & Submission Date -->
    <div id="step1">
        <fieldset>
            <legend>WHQL Submission</legend>
            <br />
            <table class="FormTable" cellpadding="1" cellspacing="0">
                <tr>
                    <th>
                        Submission ID:</th>
                    <td>
                        <input id="submissionID" name="submissionID" type="text" /></td>
                </tr>
                <tr>
                    <th>
                        Submission Date:</th>
                    <td>
                        <input id="submissionDt" name="submissionDt" type="text" />
                        <a href="javascript: cmdDate_onclick('submissionDt')"><img ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a></td>
                </tr>
                <tr>
                    <th>
                        Date Released:</th>
                    <td>
                        <input id="txtDateReleased" name="txtDateReleased" type="text" />
                        <a href="javascript: cmdDate_onclick('txtDateReleased')"><img ID="Img1" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a></td>
                </tr>
                <tr>
                    <th>
                        Label Location:</th>
                    <td>
                        <textarea id="txtLabelLocation" name="txtLabelLocation" cols="40" rows="2">Monitor & Printed on Service Tag located on the bottom of unit.</textarea></td>
                </tr>
                <tr>
                    <th>
                        Logo Displayed:</th>
                    <td>
                        <input id="chkLogoDisplayed" name="chkLogoDisplayed" type="checkbox" checked /></td>
                </tr>
                <tr>
                    <th>
                        Milestone 3:</th>
                    <td>
                        <input id="chkMilestone3" name="chkMilestone3" type="checkbox" checked /></td>
                </tr>
            </table>
            <br />
        </fieldset>
    </div>
    <!-- Select the product brand -->
    <div id="step2">
        <fieldset>
            <legend>Product Brand</legend>
            <br />
            <table class="FormTable" cellpadding="1" cellspacing="0">
                <tr>
                    <th>
                        Brand:</th>
                    <td>
                        <%= BrandRadioButtons %>
                        &nbsp;</td>
                </tr>
            </table>
            <br />
        </fieldset>
    </div>
    <!-- Select the OS Family -->
    <div id="step3">
        <fieldset>
            <legend>OS Family</legend>
            <br />
            <table class="FormTable" cellpadding="1" cellspacing="0">
                <tr>
                    <th>
                        OS Family:</th>
                    <td>
                        <%= OsFamilyRadioButtons %>
                        &nbsp;</td>
                </tr>
            </table>
            <br />
        </fieldset>
    </div>
    <!-- Select Base Units that apply -->
    <div id="step4">
        <fieldset>
            <legend>Base Units Covered</legend>
            <br />
            <table class="FormTable" cellpadding="1" cellspacing="0">
                <tr>
                    <th>
                        Base Units:</th>
                    <td>
                        <%= BUCheckBoxes %>
                        &nbsp;</td>
                </tr>
            </table>
            <br />
        </fieldset>
    </div>
    <!-- Select Processors that apply -->
    <div id="step5">
        <fieldset>
            <legend>Processors Covered</legend>
            <br />
            <table class="FormTable" cellpadding="1" cellspacing="0">
                <tr>
                    <th>
                        Processors:</th>
                    <td>
                        <%= CPUCheckBoxes %>
                    </td>
                </tr>
            </table>
            <br />
        </fieldset>
    </div>
    <div id="step6">
        <fieldset>
            <legend>Submission Preview</legend>
            <br />
        </fieldset>
    </div>
    </form>
</body>
</html>

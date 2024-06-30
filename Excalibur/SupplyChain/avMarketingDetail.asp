<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->
<%
'Response.Write Request.QueryString
'Response.End
response.Buffer = true

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim i
Dim sMode				: sMode = Request.QueryString("Mode")
Dim sAvNo				: sAvNo = ""
Dim sCategoryOpt		: sCategoryOpt = ""
Dim	iCategoryOpt		: iCategoryOpt = ""
Dim sSCMCat			    : sSCMCat = ""
Dim sCategoryAbbr       : sCategoryAbbr = ""
Dim sPlatformName       : sPlatformName = ""
Dim sGpgDesc			: sGpgDesc = ""
Dim sMarketingDesc		: sMarketingDesc = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sProgramVersion		: sProgramVersion = GetProductVersion(Request("PVID"))
Dim sConfigRules		: sConfigRules = ""
Dim bIdsSkus			: bIdsSkus = false
Dim bIdsCto				: bIdsCto = false
Dim bRctoSkus			: bRctoSkus = false
Dim bRctoCto			: bRctoCto = false
Dim sUpc				: sUpc = "&nbsp;"
Dim saBrands			: saBrands= Split(Request.Form("chkBrand"), ",")
Dim sCbxBrand			: sCbxBrand = ""
Dim sStatus				: sStatus = "A"
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim iBrandID			: iBrandID = ""
Dim sCplBlindDt			: sCplBlindDt = ""
Dim sRasDiscDt			: sRasDiscDt = ""
Dim sPAADDate			: sPAADDate = ""    
Dim sWeight				: sWeight = ""
Dim sRTPDt		        : sRTPDt = ""
Dim sPhWebInstruction	: sPhWebInstruction = ""
Dim sSortOrder			: sSortOrder = ""
Dim sSDFFlag            : sSDFFlag = 0
Dim sIsDesktop         
Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_IsMarketingUser
Dim m_EditModeOn
Dim m_UserFullName
Dim strProductLineID	: strProductLineID = ""
Dim strProductLineName	: strProductLineName = ""
Dim bSharedAV           : bSharedAV = false
Dim sFeatureID		    : sFeatureID = ""

Dim sCreated 
Dim sCreatedBy
Dim sUpdated
Dim sUpdatedBy
Dim scmlist : scmlist = ""
Dim sGeneralAvailDt  : sGeneralAvailDt = ""
Dim bAlgorithm           : bAlgorithm = ""
Dim sEOM            : sEOM = "" 
Dim sRTP            : sRTP = "" 
Dim sChangeNote         : sChangeNote= ""
Dim m_UserID
Dim sProductReleaseIDs : sProductReleaseIDs = Request.QueryString("Release")
'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

'
' Debug Section
'
'	If Security.CurrentUserID = 1396 Then
'		m_IsSysAdmin = False
'		Security.CurrentUserID = 1288
'		Response.Write Security.CurrentUserID
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSysEngProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSystemTeamLead(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramCoordinator(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write m_ProductVersionID
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If


	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
    Dim MarketingProductCount : MarketingProductCount = 0
    Dim CurrentUser : CurrentUser = lcase(Session("LoggedInUser"))
    Dim CurrentDomain
    If instr(CurrentUser,"\") > 0 Then
        CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
        CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
    End If
    
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Set cmd = dw.CreateCommAndSP(cn, "spGetUserInfo")
    dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, CurrentUser
    dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
    Set rs = dw.ExecuteCommandReturnRS(cmd)

    If not (rs.EOF And rs.BOF) Then
	    'add the permission from the Users and Roles to the Pulsar products
        If Not m_IsMarketingUser Then
		    MarketingProductCount = rs("MarketingProductCount")
            if MarketingProductCount > 0 then
                m_IsMarketingUser = True
            end if
	    End If
    End If
    rs.Close

	m_UserFullName = Security.CurrentUserFullName()
    m_UserID = Security.CurrentUserId()
	
	If m_IsMarketingUser Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges for Marketing</H3><H4>Access Denied</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	
'
'PC Can do any thing
'Marketing can change description


Function GetProductVersion( ProductVersionID )
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	GetProductVersion = rs("version") & ""
	
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Function

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Function GetBoolValue( value )
    if(len(trim(value)) > 0) then
	    Select Case UCase( value )
		    Case "Y"
			    GetBoolValue = true
		    Case "N"
			    GetBoolValue = false
		    Case 1
			    GetBoolValue = true
		    Case 0
			    GetBoolValue = false
		    Case "T"
			    GetBoolValue = true
		    Case "F"
			    GetBoolValue = false
		    Case Else
			    GetBoolValue = false
	    End Select
    else
        GetBoolValue = false
    end if
End Function

Sub Main()
'
'TODO: Get AvDetail Data
'
    Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sIsDesktop = rs("IsDesktop") 
	
	rs.Close

	Set cmd = dw.CreateCommAndSP(cn, "spListBrAnds4Product")
	dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID")
	dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
	
	Do Until rs.EOF
		sCbxBrand = sCbxBrand & "<input type=checkbox id=chkBrand name=chkBrand value=" & rs("ProductBrandID") 
		If Trim(rs("ProductBrandID")) = Trim(Request("BID")) Then
			sCbxBrand = sCbxBrand & " CHECKED "
		End If
		sCbxBrand = sCbxBrand & ">" & rs("Name") & "<BR>"
		rs.MoveNext
	Loop
	
	sCbxBrand = Left(sCbxBrand, Len(sCbxBrand) - 4)

	If Request("AVID") <> "" Then	'Get the values for the request AV

		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail_Pulsar")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)

        sFeatureID = rs("FeatureID")        
		sAvNo = rs("AvNo")
		iCategoryOpt = rs("SCMCategoryID")
        sCategoryAbbr = rs("SCMCategoryAbbr")
        sPlatformName = rs("PlatformName")
		sSCMCat = rs("Name")
		sGpgDesc = rs("GPGDescription")
		sMarketingDesc = rs("MarketingDescription")
		sConfigRules = rs("ConfigRules")
		sManufacturingNotes = rs("ManufacturingNotes")
		bIdsSkus = GetBoolValue(rs("IdsSkus_YN"))
		bIdsCto = GetBoolValue(rs("IdsCto_YN"))
		bRctoSkus = GetBoolValue(rs("RctoSkus_YN"))
		bRctoCto = GetBoolValue(rs("RctoCto_YN"))
		sUpc = rs("UPC")
		sStatus = rs("Status")
		iBrandID = rs("ProductBrandID")
		sCplBlindDt = rs("CplBlindDt")
		sRasDiscDt = rs("RasDiscontinueDt")
		sWeight = rs("Weight")
		sRTPDt = rs("RTPDate")
		sPhWebInstruction = rs("PhWebInstruction")
		sSDFFlag = rs("SDFFlag")
		sSortOrder = rs("SortOrder")
        strProductLineID = rs("ProductLineID")
        strProductLineName = rs("ProductLine")
        if LCase(rs("bSharedAV")) = "true" then
            bSharedAV = true
        end if

        sPAADDate=rs("PHWebDate")
        'if sIsDesktop=True Then 'Desktops
        '    sPAADDate=rs("GeneralAvailDt")	
        'else 'notebooks
        '    sPAADDate=rs("PHWebDate")
        'end if
    
        sCreated = rs("Created")
        sCreatedBy = rs("CreatedBy")
        sUpdated = rs("Updated")
        sUpdatedBy = rs("UpdatedBy")
        sGeneralAvailDt = rs("GeneralAvailDt")
        sChangeNote = rs("ChangeNote")
    	rs.Close

        'Get the Releases for the AV
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_ProductRelease_OnAV")
	    dw.CreateParameter cmd, "@PBID", adInteger, adParamInput, 8, Trim(Request("BID")) 
        dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
	    Set rs = dw.ExecuteCommandReturnRS(cmd)
        sProductReleaseIDs = rs("PRIDs")
        rs.Close

	End If
	
    if sProductReleaseIDs <> "" then
        'Get End of manufaturing date
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_EOMDate")
		    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
            dw.CreateParameter cmd, "@p_ReleaseID", adVarchar, adParamInput, 200, sProductReleaseIDs
		    Set rs = dw.ExecuteCommandReturnRS(cmd)  
            if not rs.BOF and not rs.EOF then
                if not IsNull(rs("EOM")) then
                    sEOM = rs("EOM")
                end if
            end if
            rs.Close 
        
        'Get RTP date
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_RTPDate")
		    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
            dw.CreateParameter cmd, "@p_ReleaseID", adVarchar, adParamInput, 200, sProductReleaseIDs
		    Set rs = dw.ExecuteCommandReturnRS(cmd)
            if not rs.BOF and not rs.EOF then
                if not IsNull(rs("RTP")) then
                    sRTP = rs("RTP")
                end if
            end if
            rs.Close 
    end if

	Set cmd = dw.CreateCommandSP(cn, "usp_SCM_GetSCMCategories")
	'dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.EOF
		sCategoryOpt = sCategoryOpt & "<OPTION Value='" & rs("SCMCategoryID") & "'"
		If iCategoryOpt = rs("SCMCategoryID") Then
			sCategoryOpt = sCategoryOpt & " SELECTED "
		End If
		sCategoryOpt = sCategoryOpt & ">" & rs("Name") & "</OPTION>" & VbCrLf
		rs.MoveNext
	Loop

	rs.Close 


    If bSharedAV = true Then
        'Response.Write("<script language=VBScript>MsgBox """ + scmlist + """</script>") 
         Dim rsSCM
         Set cmd = dw.CreateCommAndSP(cn, "usp_GetSCMNames")
         dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, sAvNo
         Set rsSCM = dw.ExecuteCommAndReturnRS(cmd)

         If rsSCM.EOF Then
              scmlist = ""
        Else
             Do While Not rsSCM.EOF
             'Response.Write("<script language=VBScript>MsgBox """ + rsSCM.Fields("DOTSName").value + """</script>") 
             scmlist = rsSCM.Fields("DOTSName").value + ","  + scmlist   
             rsSCM.MoveNext()
             Loop
                 'Response.Write("<script language=VBScript>MsgBox """ + scmlist + """</script>") 
         End If
        rsSCM.Close  

   End If

	
End Sub

Function GetSkuCtoValue( value )
	If lcase(Value) = "y" Then
		GetSkuCtoValue = "X"
	Else
		GetSkuCtoValue = "&nbsp;"
	End If
End Function

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	Dim iAvId
	
	cn.BeginTrans
	iAvId = Request("AVID")
		
	'Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand_RTPDate")
	'dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
    'dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request("BID")
	'dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("txtRTPDate")
	'dw.ExecuteNonQuery(cmd)
	
    Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sIsDesktop = rs("IsDesktop") 
	
	
    'Save AvDetail data
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_Pulsar")
	dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
	dw.CreateParameter cmd, "@p_SCMCategoryID", adInteger, adParamInput, 8, Request.Form("hidSCMCategoryID")
	dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtGPGDescription")
	dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
	dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 40, ""
	
  
    '----------------check if SA value is from edit or from calculation-----------------------------------
    If Request.Form("hdnflag1") = "E" Then
       'Response.Write("<script language=VBScript>MsgBox """ + "if loop cplblinddate" + """</script>")
        dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 50, Request.Form("txtBlindDate1")
     else 
       'Response.Write("<script language=VBScript>MsgBox """ + "else loop cplblinddate" + """</script>")
        dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 50, Request.Form("hidCplBlindDt")
    end IF

   
    dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 50, Request.Form("txtMarketingDiscDate")
	dw.CreateParameter cmd, "@p_UPC", adVarchar, adParamInput, 12, Request.Form("hidUpc")
	dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, Request.Form("txtChangeNote")
	dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
	dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("txtRTPDate")
    
    'Response.Write("<script language=VBScript>MsgBox """ + Request.Form("txtRTPDate") + """</script>")
	
    'Dien: For both Notebook and Desktop: RTP/MR/FCS date = General Availability (GA) Date PBI 6352
    'dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50,  Request.Form("txtRTPDate")   
    
     '----------------check if GA value is from edit or from calculation-----------------------------------   
    If Request.Form("hdnflag2") = "E" Then
       'Response.Write("<script language=VBScript>MsgBox """ + "if loop GA date" + """</script>")
        dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50, Request.Form("txtGeneralAvailDt1")
     else 
       'Response.Write("<script language=VBScript>MsgBox """ + "else loop GA date" + """</script>")
       'dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50,  Request.Form("txtRTPDate") 
       dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50,  Request.Form("hidGeneralAvailDt") 
      
    end IF
     

    dw.CreateParameter cmd, "@p_NameElements", adVarchar, adParamInput, 500, ""
    dw.CreateParameter cmd, "@p_weight", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, ""
    dw.CreateParameter cmd, "@p_ProductLineID", adInteger, adParamInput, 8,  Request.Form("hidProductLineID")
    

    '----------------check if PAAD value is from edit or from calculation-----------------------------------   
  
    If Request.Form("hdnflag") = "E" Then
       'Response.Write("<script language=VBScript>MsgBox """ + "if loop" + """</script>")
        dw.CreateParameter cmd, "@PhwebDate", adDate, adParamInput, 50,  Request.Form("txtPAADDate1")
     else 
       'Response.Write("<script language=VBScript>MsgBox """ + "else loop" + """</script>")
        dw.CreateParameter cmd, "@PhwebDate", adDate, adParamInput, 50,  Request.Form("txtPAADDate")
    end IF

    dw.CreateParameter cmd, "@p_FeatureID", adInteger, adParamInput, 50, sFeatureID
    dw.CreateParameter cmd, "@p_BUAvailList", adVarchar, adParamInput, 50, null
    dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request.Form("BID")
    dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, m_UserID

    returnValue = dw.ExecuteNonQuery(cmd)
      
	If Not (returnValue = 1 Or returnValue = -1) Then
   		' Abort Transaction
		Response.Write returnValue
		cn.RollbackTrans()
		Exit Sub
	End If

	sFunction = "close"
	cn.CommitTrans()

	If Request.Form("hidAVID") <> Request.QueryString("AVID") And sMode <> "add" Then
		Response.Redirect "avMarketingDetail.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request("hidAVID")
	End If

End Sub

Function GetCbxValue( value )
	If lcase(value) = "on" Or lcase(value) = "yes" Then
		GetCbxValue = 1
	Else
		GetCbxValue = 0
	End If
End Function

If LCase(sFunction) = "save" Then
	Call Save()
    Call Main()
Else
	Call Main()
End If
%>
<html>
<head>
    <title>Marketing Detail</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css" />
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {
            //$("#txtRTPDate").keyup(function () {
            //    calcDates();
            //});     


            $("#txtRTPDate").focusout(function () {
                if (document.getElementById('hdnAlgorithm').value == 'Y') {
                    calcDates();
                }
            });

            if ($("#txtPAADDate1").val() == "" && $("#txtBlindDate1").val() == "" && $("#txtGeneralAvailDt1").val() == "") {
                if ($("#hdnSharedValue1").val().toLowerCase() != 'true') {
                    $("#txtRTPDate").val($("#hidRTP").val());
                    $("#txtMarketingDiscDate").val($("#hidEOM").val());
                    document.getElementById('hdnAlgorithm_blankdates').value = 'Y';
                    calcDates_whendatesareblank();

                }
            }
        });

        function calcDates() {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (isDate($("#txtRTPDate").val())) {

                var RTPDate = new Date($("#txtRTPDate").val());
                // remove FCS from all areas - task 20243
                var GeneralAvailDt = RTPDate;

                var monday = getMonday(RTPDate);
                var firstDay = new Date(GeneralAvailDt.getFullYear(), GeneralAvailDt.getMonth(), 1);
                var PAADDate = new Date((monday.getMonth() + 1) + '/' + monday.getDate() + '/' + monday.getFullYear());

                var BlindDate;                    
                BlindDate = new Date(PAADDate.getFullYear(), (PAADDate.getMonth() - 1), 1);
                if (BlindDate < Today)
                    BlindDate = new Date(Today.getFullYear(), Today.getMonth(), Today.getDate() + 7);

                if ($("#hidCplBlindDt").val() != "") {
                    var existingsadate = new Date($("#hidCplBlindDt").val());
                    if (existingsadate < Today)
                        BlindDate = existingsadate;
                }

                $("#mktgBlindDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td></tr></table>');
                $("#hidCplBlindDt").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtPAADDate").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                $("#mktgPAADDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear() + '</span></td><td align=Right style="border:none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td></tr></table>');
                $("#mktgGeneralAvailDtText").html('<table width="100%" border="0"><tr><td style="border:none">' + (GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td></tr></table>');
                $("#hidGeneralAvailDt").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());

                if (document.getElementById('hdnAlgorithm').value == 'Y') {
                    $("#txtBlindDate1").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                    $("#txtGeneralAvailDt1").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                    $("#txtPAADDate1").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                }                
            }
        }


        function calcDates_whendatesareblank() {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (isDate($("#txtRTPDate").val())) {

                var RTPDate = new Date($("#txtRTPDate").val());

                var GeneralAvailDt = RTPDate;

                var monday = getMonday(RTPDate);
                var firstDay = new Date(GeneralAvailDt.getFullYear(), GeneralAvailDt.getMonth(), 1);
                var PAADDate = new Date((monday.getMonth() + 1) + '/' + monday.getDate() + '/' + monday.getFullYear());

                var BlindDate;
                BlindDate = new Date(PAADDate.getFullYear(), (PAADDate.getMonth() - 1), 1);
                if (BlindDate < Today)
                    BlindDate = new Date(Today.getFullYear(), Today.getMonth(), Today.getDate() + 7);

                if ($("#hidCplBlindDt").val() != "") {
                    var existingsadate = new Date($("#hidCplBlindDt").val());
                    if (existingsadate < Today)
                        BlindDate = existingsadate;
                }

                $("#mktgBlindDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td></tr></table>');
                $("#hidCplBlindDt").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtPAADDate").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                $("#mktgPAADDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear() + '</span></td><td align=Right style="border:none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td></tr></table>');
                $("#mktgGeneralAvailDtText").html('<table width="100%" border="0"><tr><td style="border:none">' + (GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td></tr></table>');
                $("#hidGeneralAvailDt").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());

                $("#txtBlindDate1").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtGeneralAvailDt1").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                $("#txtPAADDate1").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());


                if (document.getElementById('hdnAlgorithm_blankdates').value == 'Y') {
                    $("#RTP1").html($("#hidRTP").val());
                    $("#SA1").html((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                    $("#GA1").html((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                    $("#PAAD1").html((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                    $("#EOM1").html($("#hidEOM").val());
                }
            }
        }

        function getMonday(d) {
            d = new Date(d);
            var day = d.getDay(),
                diff = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday
            return new Date(d.setDate(diff));
        }

        function isDate(txtDate) {
            var currVal = txtDate;
            if (currVal == '')
                return false;

            var rxDatePattern = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
            var dtArray = currVal.match(rxDatePattern); // is format OK?

            if (dtArray == null)
                return false;

            dtMonth = dtArray[1];
            dtDay = dtArray[3];
            dtYear = dtArray[5];

            if (dtMonth < 1 || dtMonth > 12)
                return false;
            else if (dtDay < 1 || dtDay > 31)
                return false;
            else if ((dtMonth == 4 || dtMonth == 6 || dtMonth == 9 || dtMonth == 11) && dtDay == 31)
                return false;
            else if (dtMonth == 2) {
                var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
                if (dtDay > 29 || (dtDay == 29 && !isleap))
                    return false;
            }

            return true;
        }


        function Body_OnLoad() {

            var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
            switch (frmMain.hidFunction.value) {
                case "close":
                    if (window.frmMain.hidMode.value.toLowerCase() == 'edit') {
                        if (IsFromPulsarPlus()) {
                            ClosePulsarPlusPopup();
                        } else {
                            var AvID = '', GPGDescription = '', MarketingDesc = '', RTPDate = '', RASDisDate = '', PAADDate = '', SADate = ''; GADate = '';
                            AvID = frmMain.hidAVID.value;
                            GPGDescription = frmMain.txtGPGDescription.value;
                            MarketingDesc = frmMain.txtMarketingDesc.value;
                            RTPDate = frmMain.txtRTPDate.value;
                            RASDisDate = frmMain.txtMarketingDiscDate.value;
                            PAADDate = frmMain.txtPAADDate1.value;
                            SADate = frmMain.hidCplBlindDt.value;
                            GADate = frmMain.hidGeneralAvailDt.value;
                            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                                // For Closing current popup if Called from pulsarplus
                                parent.window.parent.closeExternalPopup();
                            }
                            else {
                                parent.window.parent.ReloadAVDataFromMkt(AvID, GPGDescription, MarketingDesc, RTPDate, RASDisDate, PAADDate, SADate, GADate);
                                parent.window.parent.ClosePropertiesDialog();
                            }
                        }
                    }
                    else
                        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                            parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                            // For Closing current popup if Called from pulsarplus
                            parent.window.parent.closeExternalPopup();
                        }
                        else {
                            if (IsFromPulsarPlus()) {
                                ClosePulsarPlusPopup();
                            } else {
                                parent.window.parent.ClosePropertiesDialog(1);
                            }
                        }

                    break;
            }


            if (typeof (window.parent.frames["LowerWindow"].frmButtons) == 'object') {
                if (window.frmMain.hidMode.value.toLowerCase() == 'edit' || window.frmMain.hidMode.value.toLowerCase() == 'add' || window.frmMain.hidMode.value.toLowerCase() == 'editdates')
                    window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = false;
            }

        }

        function EditMktgBlindDate() {
            mktgBlindDate.style.display = "";
            mktgBlindDateText.style.display = "none";
        }

        function EditMktgBlindDate1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgBlindDate.style.display = "";
                    document.getElementById('hdnflag1').value = 'E';
                    mktgBlindDateText.style.display = "none";
                }
                else {
                    mktgBlindDate.style.display = "none";
                    mktgBlindDateText.style.display = "";
                }
            }
            else {
                mktgBlindDate.style.display = "";
                document.getElementById('hdnflag1').value = 'E';
                mktgBlindDateText.style.display = "none";
            }
        }


        function EditMktgGeneralAvailDate() {
            mktgGeneralAvailDt.style.display = "";
            mktgGeneralAvailDtText.style.display = "none";
        }

        function EditMktgGeneralAvailDt1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgGeneralAvailDt.style.display = "";
                    document.getElementById('hdnflag2').value = 'E';
                    mktgGeneralAvailDtText.style.display = "none";
                }
                else {
                    mktgGeneralAvailDt.style.display = "none";
                    mktgGeneralAvailDtText.style.display = "";
                }
            }
            else {
                mktgGeneralAvailDt.style.display = "";
                document.getElementById('hdnflag2').value = 'E';
                mktgGeneralAvailDtText.style.display = "none";
            }
        }


        function EditMktgDiscDate() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgDiscDate.style.display = "";
                    mktgDiscDateText.style.display = "none";
                }
                else {
                    mktgDiscDate.style.display = "none";
                    mktgDiscDateText.style.display = "";
                }
            }
            else {
                mktgDiscDate.style.display = "";
                mktgDiscDateText.style.display = "none";
            }
        }


        function EditMktgRTPDate() {

            if ($('#hdnSharedValue1').val() == "True") {

                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    // remove FCS from all areas - task 20243
                    $("#dialog").attr("title", "Please Confirm");
                    $("#dialog").html("<p>Apply this new RTP/MR Date to reset the Select Availability (SA) Date, General Availability (GA) Date, and PA:AD (Intro Date) for this AV in this Product? </p> <p><ul><li> General Availability (GA) Date = RTP/MR Date</li> <li> PA:AD (Intro) Date = Monday of the week of RTP/MR Date </li> <li> Select Availability (SA) Date = one month prior to PA:AD and always the first of the month </li></p>");
                    $("#dialog").dialog({
                        resizable: false,
                        width: 600,
                        height: 400,
                        modal: true,
                        buttons: {
                            "Yes": function () {
                                $(this).dialog("close");
                                mktgRTPDate.style.display = "";
                                mktgRTPDateText.style.display = "none";
                            },

                            "No": function () {
                                $(this).dialog("close");
                                $("#txtRTPDate").unbind('keyup');
                                mktgRTPDate.style.display = "";
                                mktgRTPDateText.style.display = "none";
                            }
                        }
                    });
                }

                else {
                    mktgRTPDate.style.display = "none";
                    mktgRTPDateText.style.display = "";
                }

            }
            else {
                $("#dialog").attr("title", "Please Confirm");
                $("#dialog").html("<p>Apply this new RTP/MR Date to reset the Select Availability (SA) Date, General Availability (GA) Date, and PA:AD (Intro Date) for this AV in this Product? </p> <p><ul><li> General Availability (GA) Date = RTP/MR Date</li> <li> PA:AD (Intro) Date = Monday of the week of RTP/MR Date </li> <li> Select Availability (SA) Date = one month prior to PA:AD and always the first of the month </li></p>");
                $("#dialog").dialog({
                    resizable: false,
                    width: 600,
                    height: 400,
                    modal: true,
                    buttons: {
                        "Yes": function () {
                            $(this).dialog("close");
                            mktgRTPDate.style.display = "";
                            mktgRTPDateText.style.display = "none";
                            document.getElementById('hdnAlgorithm').value = 'Y';
                        },

                        "No": function () {
                            $(this).dialog("close");
                            $("#txtRTPDate").unbind('keyup');
                            mktgRTPDate.style.display = "";
                            mktgRTPDateText.style.display = "none";
                            document.getElementById('hdnAlgorithm').value = '';
                        }
                    }
                });
            }
        }


        function EditMktgPAADDate() {
            mktgPAADDate.style.display = "";
            mktgPAADDateText.style.display = "none";
        }


        function EditMktgPAADDate1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgPAADDate.style.display = "";
                    document.getElementById('hdnflag').value = 'E';
                    mktgPAADDateText.style.display = "none";
                }
                else {
                    mktgPAADDate.style.display = "none";
                    mktgPAADDateText.style.display = "";
                }
            }
            else {
                mktgPAADDate.style.display = "";
                document.getElementById('hdnflag').value = 'E';
                mktgPAADDateText.style.display = "none";
            }
        }


        function OpenFeatureProperties(FeatureID, DeliveryType) {
            var url = "../../IPulsar/Features/FeatureProperties.aspx?FeatureID=" + FeatureID + "&DeliveryType=" + DeliveryType + "&ViewFrom=AvDetail" + "&IsDesktop=" + "<%=LCase(sIsDesktop) %>";
            if (IsFromPulsarPlus() || $("#pulsarplusDivId") != "") {
                url = "../../IPulsar/Features/FeatureProperties.aspx?FeatureID=" + FeatureID + "&DeliveryType=" + DeliveryType + "&app=PulsarPlus&ViewFrom=AvDetail&IsDesktop=" + "<%=LCase(sIsDesktop) %>";
                strID = window.showModalDialog(url, "Feature Properties", "dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
		if (typeof (strID) != "undefined") {
		    var retValues = strID;
                    ResetAvDescriptions(retValues.refreshgrid, retValues.FeatureID, retValues.FeatureName, retValues.GPGDescription, retValues.MarketingDescription, retValues.MarketingDescriptionPMG, retValues.RequiresRoot, retValues.ComponentLinkage, retValues.ComponentRootID, 0, 0, "", "");
		}
            } else {
            parent.window.parent.ShowFeatureSelectDialog(url, "Feature Properties", 980, 800);
	    }
        }
        function ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform) {
            if (Refresh) {
                document.getElementById('txtGPGDescription').value = GPGDescription;
                $("#gpgText").html(GPGDescription);
                document.getElementById('txtMarketingDesc').value = MarketingDescription;
                $("#mktgText").html(MarketingDescription);
            }
        }

    </script>
</head>
<body onload="Body_OnLoad()">
    <form method="post" id="frmMain">
        <input id="hidMode" name="hidMode" type="hidden" value="<%= LCase(sMode)%>" />
        <input id="hidFunction" name="hidFunction" type="hidden" value="<%= LCase(sFunction)%>" />
        <input id="hidStatus" name="hidStatus" type="hidden" value="<%= UCase(sStatus)%>" />
        <input id="hidAVID" name="hidAVID" type="hidden" value="<%= Request("AVID")%>" />
        <input id="hidSCMCategoryID" name="hidSCMCategoryID" type="hidden" value="<%= iCategoryOpt%>" />
        <input id="hidProductLineID" name="hidProductLineID" type="hidden" value="<%= strProductLineID%>" />
        <input id="txtGPGDescription" name="txtGPGDescription" type="hidden" value="<%= sGpgDesc%>" />
        <input id="AvNo" name="AvNo" type="hidden" value="<%=sAvNo%>" />
        <input type="hidden" id="txtMarketingDesc" name="txtMarketingDesc" value="<%= sMarketingDesc%>" />
        <input type="hidden" id="txtPAADDate" name="txtPAADDate" value="<%= sPAADDate %>" />
        <input id="hidCplBlindDt" name="hidCplBlindDt" type="hidden" value="<%= sCplBlindDt%>" />
        <input id="hidRasDidcontinueDt" name="hidRasDidcontinueDt" type="hidden" value="<%= sRasDiscDt%>" />
        <input id="hidRTPDt" name="hidRTPDt" type="hidden" value="<%= sRTPDt%>" />
        <!--<input id="hidGeneralAvailDt" name="hidGeneralAvailDt" type="hidden" value="<%= sRTPDt%>" />-->
        <input id="hidGeneralAvailDt" name="hidGeneralAvailDt" type="hidden" value="<%= sGeneralAvailDt%>" />
        <input id="hidPhWebInstruction" name="hidPhWebInstruction" type="hidden" value="<%= sPhWebInstruction%>" />
        <input id="hidUpc" name="hidUpc" type="hidden" value="<%= sUpc%>" />
        <input id="BID" name="BID" type="hidden" value="<%= iBrandID%>" />
        <input id="PVID" name="PVID" type="hidden" value="<%= m_ProductVersionID %>" />
        <input id="hidIsDesktop" name="hidIsDesktop" type="hidden" value="<%=sIsDesktop%>" />
        <input style="display: none" type="text" id="txtIsMarketingScreen" name="txtIsMarketingScreen" value="1">

        <input type="hidden" id="hdnSharedValue1" name="hdnSharedValue1" value="<%= bSharedAV %>" />
        <input type="hidden" id="hdnscmlist" name="hdnscmlist" value="<%= scmlist %>" />
        <input type="hidden" id="hdnflag" name="hdnflag" value="" />
        <input type="hidden" id="hdnflag1" name="hdnflag1" value="" />
        <input type="hidden" id="hdnflag2" name="hdnflag2" value="" />
        <input type="hidden" id="hdnAlgorithm" name="hdnAlgorithm" value="<%=bAlgorithm%>" />
        <input type="hidden" id="hdnAlgorithm_blankdates" name="hdnAlgorithm_blankdates" value="" />
        <input id="hidEOM" name="hidEOM" type="HIDDEN" value="<%= sEOM%>" />
        <input id="hidRTP" name="hidRTP" type="HIDDEN" value="<%= sRTP%>" />


        <table class="FormTable" width="100%" border="1" cellspacing="0" cellpadding="1" style="background-color: cornsilk; border-color: tan;">
            <tr>
                <th>Feature&nbsp;ID:</th>
                <td>
                    <table width="100%" border="0">
                        <tr>
                            <td style="border: none">
                                <div id="assignfeatureIDText"><%= PrepForWeb(sFeatureID)%></div>
                            </td>
                            <td style="border: none; text-align: right; font-size: x-small">
                                <%If sFeatureID = "" or sFeatureID = "0"  or ISNULL(sFeatureID) Then%>
                                <%Else%>
                                <div id="divfeatureIDDesc"><a href="javascript:OpenFeatureProperties(<%=sFeatureID%>, 'SRP')" id="lnkShowFeature1">View Feature</a></div>
                                <%End If%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <th>Shared AV?</th>
                <td>
                    <%If bSharedAV Then %>
                        Yes (SCM Category and product Line for Shared AVs may only be changed in Shared AV Admin)
		            <%else%>
                        No            
                    <%	End If %>               
                </td>
            </tr>
            <tr>
                <th>SCM Category</th>
                <td><%= PrepForWeb(sSCMCat)%></td>
            </tr>
            <tr>
                <th>Product&nbsp;Line:</th>
                <td><%= PrepForWeb(strProductLineName)%></td>
            </tr>
            <tr>
                <th>AV#</th>
                <td><%= PrepForWeb(sAvNo)%></td>
            </tr>
            <tr>
                <th>Status</th>
                <td><%= PrepForWeb(sStatus)%></td>
            </tr>

            <% if sCategoryAbbr = "BUNIT" then %>
            <tr>
                <th>Platform</th>
                <td>
                    <%=PrepForWeb(sPlatformName)%>
                </td>
            </tr>
            <% end if %>
            <tr>
                <th>GPG Description</th>
                <td>
                    <div id="gpgText"><%= PrepForWeb(sGpgDesc)%></div>
                </td>
            </tr>
            <tr>
                <th>Marketing Description</th>
                <td>
                    <div id="mktgText"><%= PrepForWeb(sMarketingDesc)%></div>
                </td>
            </tr>
            <tr>
                <th>RTP/MR Date</th>
                <td>
                    <div id="mktgRTPDate" style="display: none">
                        <input type="text" id="txtRTPDate" name="txtRTPDate" value="<%= sRTPDt%>" style="width: 300px" autocomplete='off' /></div>
                    <div id="mktgRTPDateText">
                        <table width="100%" border="0">
                            <tr>
                                <td id="RTP1" style="border: none"><%= PrepForWeb(sRTPDt)%></td>
                                <%If not bSharedAV Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditMktgRTPDate();">Edit</a></td>
                                <% end if %>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <th>PA:AD (Intro Date)<sup style="color: green;"> [1]</sup></th>
                <td>
                    <div id="mktgPAADDate" style="display: none">
                        <input type="text" id="txtPAADDate1" name="txtPAADDate1" value="<%= sPAADDate%>" style="width: 300px" autocomplete='off' /></div>
                    <div id="mktgPAADDateText">
                        <table width="100%" border="0">
                            <tr>
                                <td id="PAAD1" style="border: none"><%= PrepForWeb(sPAADDate)%></td>
                                <%If not bSharedAV Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td>
                                <% end if %>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <tr>
                <th>Select Availability (SA) Date<sup style="color: green;"> [1]</sup></th>
                <td>
                    <div id="mktgBlindDate" style="display: none">
                        <input type="text" id="txtBlindDate1" name="txtBlindDate1" value="<%= sCplBlindDt%>" style="width: 300px" autocomplete='off' /></div>
                    <div id="mktgBlindDateText">
                        <table width="100%" border="0">
                            <tr>
                                <td id="SA1" style="border: none"><%= PrepForWeb(sCplBlindDt)%></td>
                                <%If not bSharedAV Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td>
                                <% end if %>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <tr>
                <th>General Availability (GA) Date<sup style="color: green;"> [1]</sup></th>
                <td>
                    <div id="mktgGeneralAvailDt" style="display: none">
                        <input type="text" id="txtGeneralAvailDt1" name="txtGeneralAvailDt1" value="<%= sGeneralAvailDt%>" style="width: 300px" autocomplete='off' /></div>
                    <div id="mktgGeneralAvailDtText">
                        <table width="100%" border="0">
                            <tr>
                                <td id="GA1" style="border: none"><%= PrepForWeb(sGeneralAvailDt)%></td>
                                <%If not bSharedAV Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td>
                                <% end if %>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <tr>
                <th>End of Manufacturing (EM) Date</th>
                <td>
                    <div id="mktgDiscDate" style="display: none">
                        <input type="text" id="txtMarketingDiscDate" name="txtMarketingDiscDate" value="<%= sRasDiscDt%>" style="width: 300px" autocomplete='off' /></div>
                    <div id="mktgDiscDateText">
                        <table width="100%" border="0">
                            <tr>
                                <td id="EOM1" style="border: none"><%= PrepForWeb(sRasDiscDt)%></td>
                                <%If not bSharedAV Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditMktgDiscDate();">Edit</a></td>
                                <% end if %>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <th>Reason for Change</th>
                <td>
                    <textarea rows="2" id="txtChangeNote" name="txtChangeNote" style="width: 600px" maxlength="600"><%= sChangeNote%></textarea></td>
            </tr>
            <!-- Do not need this PHweb Instructions any more for pulsar
            -->

            <tr>
                <th>Created</th>
                <td><%= PrepForWeb(sCreated)%></td>
            </tr>
            <tr>
                <th>Created By</th>
                <td><%= PrepForWeb(sCreatedBy)%></td>
            </tr>
            <tr>
                <th>Updated</th>
                <td><%= PrepForWeb(sUpdated)%></td>
            </tr>
            <tr>
                <th>Updated By</th>
                <td><%= PrepForWeb(sUpdatedBy)%></td>
            </tr>

        </table>
        <div style="font-size: xx-small; color: green; font-style: italic;">
            <p></p>
            1. PA:AD (Intro Date), Selected Availability (SA) Date, and General Availability (GA) Date are calculated based on the RTP/MR date and automatically updated</div>
        <div id="dialog" title="Confirmation"></div>
    </form>    
    <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>
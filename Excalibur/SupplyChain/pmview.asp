<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file="../includes/lib_debug.inc" -->
<%
Server.ScriptTimeout = 480
Response.Clear
'response.Buffer = false
' --- GLOBAL & OPTIONAL INCLUDES: ---%>
<!--#INCLUDE FILE="../includes/oConnect.asp"-->
<!--#INCLUDE FILE="../includes/orsProgramMatrix.asp"-->
<%
'--- INSTANTIATE OBJECTS: ---
Call OpenDBConnection(PULSARDB(), True)

Dim regEx
Set regEx = New RegExp
regEx.Global = True

regEx.Pattern = "[^0-9]"
Dim PVID : PVID = regEx.Replace(Request.QueryString("ID"), "")
regEx.Pattern = "[^0-9a-zA-Z_ ]"
Dim sList : sList = regEx.Replace(Request("List"), "")
If sList = "" Then
	on error resume next
    sList = regEx.Replace(Request.Cookies("PMTab"), "")
    on error goto 0
End If

regEx.Pattern = "[^0-9a-zA-Z_# ]"
Dim sClass : sClass = regEx.Replace(Request("Class"), "")
Dim sSeries : sSeries = regEx.Replace(Request("Series"), "")
Dim sGroupBy : sGroupBy = regEx.Replace(Request("GroupBy"), "")
Dim sStatus : sStatus = regEx.Replace(Request("Status"), "")
Dim sInterval : sInterval = regEx.Replace(Request("Interval"), "")
Dim sRegion : sRegion = regEx.Replace(Request("Region"), "")

Dim strNoticeTable

Dim rs, dw, cn, cmd
Dim iSCMCategoryID : iSCMCategoryID = -1
Dim sSCMCategory, chAvValue
Dim iAvCount
Dim bFirstWrite
Dim bIsPc : bIsPc = False
Dim bIsPcCm : bIsPcCm = False
Dim bIsPDM : bIsPDM = False
dim bIsRpdm : bIsRpdm = False
Dim bUnpublished : bUnpublished = False
Dim bShowPublishRollback : bShowPublishRollback = False
Dim m_BrandID : m_BrandID = ""
Dim m_IsDesktop : m_IsDesktop = False
Dim m_BrandName : m_BrandName = ""
Dim m_ShortName : m_ShortName = ""
Dim m_LastPublishDt : m_LastPublishDt = ""
Dim m_BusinessID : m_BusinessID = ""
Dim m_BusinessSegmentID : m_BusinessSegmentID = ""
on error resume next
Dim strTitleColor : strTitleColor = regex.Replace(Request.Cookies("TitleColor"),"")
if strTitleColor = "" then 
	strTitleColor = "#0000cd"
end if
on error goto 0
Dim bAdministrator : bAdministrator = false
Dim CurrentUser : CurrentUser = lcase(Session("LoggedInUser"))
Dim CurrentDomain
Dim CurrentUserPartner
Dim CurrentUserName
Dim CurrentUserID
Dim CurrentUserSysAdmin
Dim CurrentWorkgroupID
Dim bPreinstallPM
Dim bCommodityPM
Dim sFavs
Dim sFavCount
Dim sProductName
Dim sDisplayedProductName
Dim SEPMID
Dim PMID
Dim PCID
Dim sProgramVersion
Dim strSCMPath
Dim sKmat
Dim strDisplayedList
Dim CurrentUserDefaultTab
Dim strProdType
Dim orphanCount
Dim UnReleasedAVCount
Dim UnpublishedAvsCount
Dim AVsWithMissingDataCount
Dim AVsWithMissingDelRootCount
Dim ShowReport
Dim sProdVersionBSAMFlag
Dim sProgramName
Dim ShowPhWebActionItems
Dim PendingAvActions : PendingAvActions = "True"
Dim sSCMCategories : sSCMCategories = ""
Dim sGADateTo : sGADateTo = ""
Dim sGADateFrom : sGADateFrom = ""
Dim sSADateTo : sSADateTo = ""
Dim sSADateFrom : sSADateFrom = ""
Dim sEMDateTo : sEMDateTo = ""
Dim sEMDateFrom : sEMDateFrom = ""
Dim sReleaseIDs : sReleaseIDs = ""
Dim sFilterApplied: sFilterApplied=""
Dim intNoLocalization : intNoLocalization = 0
Dim sPMCategories : sPMCategories = ""
Dim strSeriesSummary : strSeriesSummary = ""
Dim bMultipleSeriesNumbers : bMultipleSeriesNumbers = False
Dim sFusionRequirements : sFusionRequirements = false
Dim sBrandDisplayed : sBrandDisplayed = ""
Dim sDisplayMinMax : sDisplayMinMax = ""
Dim PulsarSystemAdmin
Dim bPulsarSystemAdmin: bPulsarSystemAdmin = false
Dim featurelist : featurelist = ""
Dim sEOM            : sEOM = "" 
Dim sRTP            : sRTP = ""   


Dim AppRoot
AppRoot = Session("ApplicationRoot")

on error resume next
ShowReport = RegEx.Replace(Request.Cookies("ShowReport"), "")
regEx.Pattern = "^scm|pm$"
If Not regEx.Test(ShowReport) Then ShowReport = ""
If ShowReport = "" Then ShowReport = "scm"

on error goto 0
If instr(CurrentUser,"\") > 0 Then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If

Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsPDM_MarketingOps
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim m_ProductVersionID : m_ProductVersionID = PVID
Dim m_CurrentUserId
Dim m_IsTestLead
Dim m_IsToolPm
Dim m_IsActivePm
Dim m_IsCommodityPM
Dim m_IsSCFactoryEngineer
Dim m_IsAccessoryPM
Dim m_IsHardwarePm
dim m_IsRpdm
Dim m_IsSCMPublishCoordinator
Dim m_IsPinPm
dim strMyBrowser
dim bIe10_11_RegularMode  
DIM m_IsMarketingUser
dim blnCanEditProduct : blnCanEditProduct = false
dim MarketingProductCount : MarketingProductCount = 0
Dim blnCanClearChangeLog : blnCanClearChangeLog = false
Dim EditABTInformation 

strMyBrowser = Request.ServerVariables("HTTP_User_Agent")

if instr(strMyBrowser,"MSIE") > 0 then
	strMyBrowser = mid(strMyBrowser,instr(strMyBrowser,"MSIE")+5)
	if left(strMyBrowser,1) < 7 and left(strMyBrowser,2) <> "10" then
		bIe10_11_RegularMode="0"
	elseif left(strMyBrowser,1) = 7 then
		bIe10_11_RegularMode="0"
	else
		bIe10_11_RegularMode="1"
	end if
else
	bIe10_11_RegularMode="1"
end if

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	bIsPc = False
    bIsPcCm = False
	bIsPDM = False
	bIsRpdm = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
    m_CurrentUserId = Security.CurrentUserId()
    m_IsTestLead = Security.IsTestLead()
    m_IsToolPm = Security.IsToolsPm()
    m_IsActivePm = Security.IsActivePm()
    m_IsCommodityPM = Security.IsCommodityPM()
    m_IsSCFactoryEngineer = Security.IsSCFactoryEngineer()
    m_IsAccessoryPM = Security.IsAccessoryPM()
    m_IsHardwarePm = Security.IsHardwarePm(m_ProductVersionID)
    m_IsPinPm = Security.UserInRole("", "PINPM")
    
    m_IsSCMPublishCoordinator = Security.UserInRole("", "SCMPUBC")
    
    m_IsPDM_MarketingOps = Security.UserInRole(m_ProductVersionID,"MARKETINGOPS")
    m_IsRpdm = Security.UserInRole("", "RPDM")
            	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		bIsPc = True
	End If
	
    If m_IsProgramCoordinator Or m_IsConfigurationManager Then
    bIsPcCm = True
	End If

	If m_IsPDM_MarketingOps Then
	    bIsPDM = True
	End If
	
	if m_IsRpdm Or m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager then
	    bIsRpdm = true
	end if
	
    m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
	
	
	Set Security = Nothing
'##############################################################################	

'
' Setup the data connections
'
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommAndSP(cn, "spGetUserInfo")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, CurrentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then
	CurrentUserName = rs("Name") & ""
	CurrentUserID = rs("ID") & ""
	CurrentUserSysAdmin = rs("SystemAdmin")
	CurrentWorkgroupID = rs("WorkgroupID") & ""
	CurrentUserPartner = trim(rs("PartnerID") & "")
	bPreinstallPM = rs("PreinstallPM") & ""
	bCommodityPM = rs("CommodityPM") & ""
	CurrentUserDefaultTab = rs("DefaultProductTab") & ""

	sFavs = trim(rs("Favorites") & "")
	sFavCount = trim(rs("FavCount") & "")
	PulsarSystemAdmin = rs("PulsarSystemAdmin")
	if PulsarSystemAdmin = 1 then
		bPulsarSystemAdmin = true
	else
		bPulsarSystemAdmin = false
	end if

    'permission needed for Edit Product link on Pulsar products
    if rs("CanEditProduct") = 1 then
        blnCanEditProduct = true
    end if

    'add the permission from the Users and Roles to the Pulsar products
    If Not m_IsMarketingUser Then
		MarketingProductCount = rs("MarketingProductCount")
        if MarketingProductCount > 0 then
            m_IsMarketingUser = True
        end if
	End If
	
    if rs("PCProductCount") >= 1 then
        blnCanClearChangeLog = true
    end if 

    EditABTInformation = rs("PCProductCount") & ""

End If
rs.Close
on error resume next
    if trim(sList) <> "" then
	    strDisplayedList = sList
    else
        strDisplayedList = "General"
    end if
		
	Response.Cookies("LastProductDisplayed") = PVID
on error goto 0

dim ShowItem
If CurrentUserPartner = "1" Then
	ShowItem = ""
Else
	ShowItem = "none"
End If

strProdType = "1"

Set cmd = dw.CreateCommAndSP(cn, "spGetProductVersion_Pulsar")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, PVID
Set rs = dw.ExecuteCommAndReturnRS(cmd)

If (rs.EOF And rs.BOF) And PVID <> "-1" Then
	Response.Write "Unable to find the selected program.<br><font size=1>ID=" & PVID & "</font>"
	Response.Write "<BR><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & PVID & ")""><font face=verdana size=1>Remove From Favorites</font></a>"
	Response.Write "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
	Response.Write "<span id=EditLink style=""Display:none""></span><span id=StatusLink style=""Display:none""></span><span id=menubar style=""Display:none""></span><span ID=Wait style=""Display:none""></span>"
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""1"">"
Else
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""0"">"
	sProductName = rs("Name") & " " & rs("Version") 
	sDisplayedProductName = rs("Name") & " " & rs("Version")
	sProgramVersion = rs("Version") & ""
	SEPMID = rs("sepmid")
	PMID = rs("PMID")
    m_IsDesktop = rs("IsDesktop")
    if rs("SMID") & "" <> "" then
        PMID = PMID & "_" & rs("SMID")
    end if
    PMID = "_" & PMID & "_"
	strSCMPath = rs("SCMPath") & ""
    strProdType = rs("TypeID") & ""
    sProdVersionBSAMFlag = rs("BSAMFlag") & ""
	sFusionRequirements = rs("FusionRequirements")
End If				
rs.Close

Function PrepForWeb( value )   
	Dim myString
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = ""
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
	    myString = Server.HTMLEncode( value )
	    myString = Replace(myString, vbCrLf, "<BR>")
	    myString = Replace(myString, vbCr , "<BR>")
	    myString = Replace(myString, vbLf, "<BR>")
	    myString = Replace(myString, Chr(10), "<BR>")
		PrepForWeb = myString
	End If

End Function

If CurrentUserSysAdmin or SEPMID = CurrentUSerID or instr(trim(PMID),"_" & trim(CurrentUSerID) & "_") > 0  Then
	bAdministrator = true
End If

'
' Get Service Family Pn
'
Set cmd = dw.CreateCommAndSP(cn, "usp_GetServiceFamilyPn")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, m_ProductVersionID
Set rs = dw.ExecuteCommAndReturnRS(cmd)

Dim sServiceFamilyPn
If Not rs.EOF Then 
    sServiceFamilyPn = rs("ServiceFamilyPn")
End If
rs.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectPendingAvActions")
dw.CreateParameter cmd, "@p_PVID", adInteger, adParamInput, 8, m_ProductVersionID
Set rs = dw.ExecuteCommAndReturnRS(cmd)

If rs.eof and rs.bof Then 
    PendingAvActions = "False"
End If
rs.Close


'Get End of manufaturing date from Product Schedule
Set cmd = dw.CreateCommandSP(cn, "usp_Get_EOMDate")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, PVID
	Set rs = dw.ExecuteCommandReturnRS(cmd)
    if not rs.BOF and not rs.EOF then
        sEOM = rs("EOM")
    end if
    rs.Close 
          
'Get RTP date from Product Schedule
Set cmd = dw.CreateCommandSP(cn, "usp_Get_RTPDate")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, PVID
	Set rs = dw.ExecuteCommandReturnRS(cmd)
    if not rs.BOF and not rs.EOF then
        sRTP = rs("RTP")
    end if
    rs.Close 


%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!--meta http-equiv="X-UA-Compatible" content="IE=8" /
this page will automatically run in IE5 mode, IE=8 will break it
-->

    <title>PM View</title>
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/SupplyChain/style.css" rel="stylesheet" type="text/css" />
    <link href="<%= AppRoot %>/SupplyChain/sample.css" rel="stylesheet" type="text/css" />
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script src="<%= AppRoot %>/includes/client/json2.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/popup.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/spin.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/jquery.spin.js" type="text/javascript"></script>
    <script src="/pulsar2/js/userfavorite.js" type="text/javascript"></script>

    <script type="text/javascript">
<!--
    $(function() {

        
        $("#scmLoading").hide();

        var strFavorites = "," + $("#txtFavs").val();
        var found = strFavorites.indexOf(",P" + $("#txtID").val() + ",");

        var expireDate = new Date();
        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "LastProductDisplayed=" + $("#txtID").val() + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";

        if ($("#txtClass").val() == "") {
            $("#EditLink").hide();
            $("#RFLink").hide();
            $("#AFLink").hide();
            $("#StatusLink").show();
        }
        else if (found == -1) {
            $("#EditLink").show();
            $("#RFLink").hide();
            $("#AFLink").show();
        }
        else {
            $("#EditLink").show();
            $("#RFLink").show();
            $("#AFLink").hide();
        }

        $("#EditLink").show();

        $("#lblAvCount").html("( " + $("#hidAvCount").val() + " AVs Displayed )");

        $("tr").on("contextmenu", function(){
            return false;
        });
            	
        $('#iframeDialog').bind('dialogclose', function(event) {
            $("#iframeDialog").dialog('destroy'); 
            ReloadIframe("#modalDialog");
        });
    });

    function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {     
       
            $("#iframeDialog").dialog({ 
                width: DlgWidth, 
                height: DlgHeight
            });
            $("#modalDialog").attr("width", "100%");
            $("#modalDialog").attr("height", "100%");
            $("#modalDialog").attr("src", QueryString);
            $("#iframeDialog").dialog("option", "title", Title);
            $("#iframeDialog").dialog("open");

            $('#iframeDialog').bind('dialogclose', function(event) {
                $("#modalDialog").attr("src", "");
            });
        }

    function ShowBUDialog(QueryString, Title, DlgWidth, DlgHeight, bIsPc) {       
        
        $("#iframeDialog").dialog({ 
            width: DlgWidth, 
            height: DlgHeight,
            buttons: {
                "Cancel": function() { 
                    $(this).dialog('destroy'); 
                    ReloadIframe("#modalDialog");
                },
                "Save" : function() { 
                    $("#modalDialog").get(0).contentWindow.Save();
                }
            }
        });
        $("#modalDialog").attr("width", "100%");
        $("#modalDialog").attr("height", "100%");
        $("#modalDialog").attr("src", QueryString);
        $("#iframeDialog").dialog("option", "title", Title);
        $("#iframeDialog").dialog("open");

        if(bIsPc)
            $(".ui-dialog-buttonset button:contains('OK')").removeAttr("disabled").removeClass("ui-state-disabled").addClass('ui-state-default');
        else 
            $(".ui-dialog-buttonpane button:contains('Save')").attr("disabled", true).addClass("ui-state-disabled");

        $('#iframeDialog').bind('dialogclose', function(event) {
            $("#iframeDialog").dialog('destroy'); 
            ReloadIframe("#modalDialog");
        });

    }

    function ReloadIframe(selector)
    {
        $(selector).contents().find("body").html('');
    }

    function ShowAVhistoryDialog(QueryString, Title) {
       
        DlgWidth = $(window).width()-5;
        DlgHeight = Math.max(document.body.scrollHeight, document.documentElement.offsetHeight, document.body.clientHeight, document.documentElement.clientHeight);
        DlgHeight = DlgHeight-300;      
        
        if ($("#txtSecondAVhistoryOpen").val() == "")
        {
            $("#txtSecondAVhistoryOpen").val("1")
        }
        else //to solve the isue that second time the av history list opens, 
            //the pop-up opens higher than told in ie 10 and 11 regilar mode
        {   if ($("#txtIe10_11_RegularMode").val() == "1")
            DlgHeight = DlgHeight -150
        }

        $("#divAVChangeHistory").dialog({ width: DlgWidth, height: DlgHeight });
        $("#ifAVChangeHistory").attr("src", QueryString);
     
        $("#divAVChangeHistory").dialog("option", "title", Title);
        $("#divAVChangeHistory").dialog("open");
    }

    function ShowSCMPublishDialog(QueryString, Title, DlgWidth, DlgHeight) {
       
        $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
       
        $("#modalDialog").attr("width", "98%");
        $("#modalDialog").attr("height", "98%");
        $("#modalDialog").attr("src", QueryString);
        $("#iframeDialog").dialog("option", "title", Title);
        $("#iframeDialog").dialog("open");
        $("#iframeDialog").attr("height", DlgHeight + "px");
    }
    function ShowSCMCatDetails()
    {
        //divSCMCatDetails
        $("#divSCMCatDetails").dialog({ width: DlgWidth, height: DlgHeight });
       
        $("#ifSCMCatDetails").attr("width", "98%");
        $("#ifSCMCatDetails").attr("height", "98%");
        $("#ifSCMCatDetails").attr("src", QueryString);
        $("#divSCMCatDetails").dialog("option", "title", Title);
        $("#divSCMCatDetails").dialog("open");
        $("#divSCMCatDetails").attr("height", DlgHeight + "px");
    }
     function ShowProperties_Product(DisplayedID, Clone, FusionRequirements) {
        var shouldClone;

        if (Clone == 1) {
            shouldClone = "&Clone=1";
        } else {
            shouldClone = "";
        }

        var DlgWidth = 1200;
        if(adjustWidth(98) < DlgWidth)
            DlgWidth = adjustWidth(98);

        var DlgHeight = adjustHeight(95);
        ShowPropertiesDialog("<%=AppRoot %>/mobilese/today/programs.asp?HWPM=0&ID=" + DisplayedID + shouldClone + "&Pulsar=" + FusionRequirements, "Product Properties", DlgWidth, DlgHeight); //980, 800);
     }
    
     function UploadScm(BID,IsDesktop,BSAMFlag,CurrentUserName,IsPc, IsMarketing) {      
         ShowSCMUploadDialog("/ipulsar/SCM/SCMUploadWorkSheet.aspx?PBID=" + BID + "&IsDesktop=" + IsDesktop + "&BSAMFlag=" + BSAMFlag + "&UserName=" + CurrentUserName + "&IsPC=" + IsPc + "&IsMarketing=" + IsMarketing, "Upload SCM Worksheet",1200,700); //580, 350 //use the same ShowRPNDialog - (changed from 350 to 400 to accomadate error message while uploading - PBI 22199)
     }
     function ShowSCMUploadDialog(QueryString, Title, DlgWidth, DlgHeight) {    
         $("#iframeDialog").dialog({ 
             width: DlgWidth, 
             height: DlgHeight
         });
         $("#modalDialog").attr("width", "100%");
         $("#modalDialog").attr("height", "100%");
         $("#modalDialog").attr("src", QueryString);
         $("#iframeDialog").dialog("option", "title", Title);
         $("#iframeDialog").dialog("open");
     }
     function ExportRPN(ProductBrandID) {
         //--- PBI 16442 - When AVs are selected on Supply Chain tab in SCM, only pull selected AVs into Export RPN -----
         var id = '';
         var selectedids = '';
         var chkAv = document.getElementsByTagName("input");
         for (i = 0; i < chkAv.length; i++) {
             if (chkAv(i).name != "")
                 if ((chkAv(i).name.indexOf("chkAv")> -1 ) && chkAv(i).checked == true)
                 {
                     ids = chkAv(i).name.substring(5);
                     selectedids = selectedids + "," + ids;
                 }
         }
         
         selectedids = selectedids.substring(1,selectedids.length);
         //alert(selectedids);
         ShowRPNDialog("/ipulsar/SCM/HPRPN.aspx?ProductBrandID=" + ProductBrandID + "&FromSCM=1" + "&AVDetailIDs=" + selectedids, "Export RPN", 580, 450); 
             }

     function ExportRPNToS4(ProductBrandID) {
         var selectedids = '';
         var chkAv = document.getElementsByTagName("input");
         for (i = 0; i < chkAv.length; i++) {
             if (chkAv(i).name != "")
                 if ((chkAv(i).name.indexOf("chkAv") > -1) && chkAv(i).checked == true) {
                     ids = chkAv(i).name.substring(5);
                     selectedids = selectedids + "," + ids;
                 }
         }
         selectedids = selectedids.substring(1, selectedids.length);
         ShowRPNDialog("/Pulsar/Report/ExportRPNToS4?requesterFrom=0&productBrandId=" + ProductBrandID + "&selectedAvIds=" + selectedids, "Export RPN to S4 for Supply Chain", 580, 450); 
     }

     function ShowRPNDialog(QueryString, Title, DlgWidth, DlgHeight) {    
         $("#iframeDialog").dialog({ 
             width: DlgWidth, 
             height: DlgHeight
         });
         $("#iframeDialog").dialog("widget")            // get the dialog widget element
            .find(".ui-dialog-titlebar-close") // find the close button for this dialog
            .hide();
         $("#modalDialog").attr("width", "98%");
         $("#modalDialog").attr("height", "98%");
         $("#modalDialog").attr("src", QueryString);
         $("#iframeDialog").dialog("option", "title", Title);
         $("#iframeDialog").dialog("open");
     }

     function ClosePropertiesDialog(strID) {
        $("#modalDialog").attr("src", "");
        $("#iframeDialog").dialog("close");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
        $("#iframeDialog").dialog('destroy'); 
        ReloadIframe("#modalDialog");
     }

    function CloseIframeDialog() {
        $("#modalDialog").attr("src", "");
        $("#iframeDialog").dialog("close");
        $("#iframeDialog").dialog('destroy'); 
        ReloadIframe("#modalDialog");
    }

    function CloseAVHistoryDialog() {
        $("#divAVChangeHistory").dialog("close");
        $("#divAVChangeHistory").dialog('destroy');         
    }

    function ShowSCMBUInformation(ProductVersionID, ProductBrandID, bIsPc)
    {
        var DlgWidth = adjustWidth(90);
        var DlgHeight = adjustHeight(95);       

        ShowBUDialog("/IPulsar/SCM/SCM_BaseUnitInformation.aspx?PVID=" + ProductVersionID + "&ProductBrandID=" + ProductBrandID, "SCM Base Unit Information", DlgWidth, DlgHeight, bIsPc);
    }


    function SCMWorkSheetShowHide(ProductBrandID, isDesktop, IsPC, IsMarketing)
    {
        var width = adjustWidth(65);
        var height = adjustHeight(85);       
        var url = "/IPulsar/SCM/SCMWorkSheet.aspx?PBID=" + ProductBrandID + "&isDesktop=" +isDesktop + "&IsPC=" + IsPC + "&IsMarketing=" + IsMarketing ;
        OpenPopUp(url, height, width, "SCM Worksheet", true, false, true, "divSCMWorksheetShowHide", "ifSCMWorksheetShowHide")
        
        var millisecondsToWait = 2000; //Run every 2 seconds to check for cookie
        var intrvl = setInterval(function () {
            if (getCookieValue("SCMWorkSheetexportcookie") != "") // if cookie exists
            {  
                document.cookie = "SCMWorkSheetexportcookie=; expires=Thu, 01 Jan 1970 00:00:00 UTC" + "; path=/"; // delete cookie
                CloseSCMWorkSheetShowHide(); // close SCMWorkSheet popup        
                clearInterval(intrvl); //Clear timer 
            }
        }, millisecondsToWait);
    }

    function CloseSCMWorkSheetShowHide() {
        $("#divSCMWorksheetShowHide").dialog("close");
    }

    function CloseABTInfoDialog() {
        $("#divABTInfo").dialog("close");
        
    }

    function CloseSCMManufacturingSitesDialog() {
        $("#divSCMManufacturingSites").dialog("close");
        
    }

    function CloseAddMultipleAVsDialog(refresh) {
        $("#divMultipleAVs").dialog("close");
        if (refresh == "true") {
            window.location.reload();
        }
    }

    // ------------------------- COMBINED AV FUNCTIONALITY--------(santodip)--------------------

    function CloseAddMultipleAVsDialogforSingleFeature(refresh,FID, FName, FCategoryID,FRequiresRoot,FComponentLinkage,FComponentRootID,FGPGDescription,FMarketingDescriptionPMG,FMarketingDescription,SCMCategoryID_singlefeature, AliasID, Abbreviation, Platform, Releases,ReleasesName, RTPDate, EMDate) 
    {
        $("#divMultipleAVs").dialog("close");
        if (refresh == "true") {
            window.location.reload();
        }
        
        if(refresh=="single")
        {
            var productversionID = $("#txtID").val();
            var brandID = $("#hidProductBrandID").val();
            var currentuserID = $("#txtUser").val();
            AddAVFunctionality(productversionID, brandID, currentuserID, FID, FName, FCategoryID,FRequiresRoot,FComponentLinkage,FComponentRootID,FGPGDescription,FMarketingDescriptionPMG,FMarketingDescription,SCMCategoryID_singlefeature, AliasID, Abbreviation, Platform, Releases,ReleasesName, RTPDate, EMDate);
        }
    }


    function AddAVFunctionality(PVID, BID, CurrentUserId, FID, FName, FCategoryID,FRequiresRoot,FComponentLinkage,FComponentRootID,FGPGDescription,FMarketingDescriptionPMG,FMarketingDescription,SCMCategoryID_singlefeature, AliasID, Abbreviation, Platform, Releases, ReleasesName, RTPDate, EMDate) {
        //alert Task 9773
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95); 

        var IsPC = $("#hidIsPC").val();
               
        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/avFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID + "&FID=" + FID + "&FName=" + encodeURIComponent(FName) + "&FRequiresRoot=" + FRequiresRoot + "&FComponentLinkage=" + FComponentLinkage + "&FComponentRootID=" + FComponentRootID + "&FGPGDescription=" + encodeURIComponent(FGPGDescription) + "&FMarketingDescriptionPMG=" + encodeURIComponent(FMarketingDescriptionPMG) + "&FMarketingDescription=" + encodeURIComponent(FMarketingDescription) + "&FSCMCategoryID_singlefeature=" + SCMCategoryID_singlefeature + "&AliasID=" + AliasID + "&Abbreviation=" + Abbreviation + "&Platform=" + Platform + "&IsPC=" + IsPC + "&Release=" + Releases + "&ReleaseName=" + ReleasesName + "&RTPDate=" + RTPDate + "&EMDate=" + EMDate, "SCM AV Details", DlgWidth, DlgHeight);
    }


    function CloseAddMultipleAVsDialogforFeatureSearch(refresh,url)
    {
        $("#divMultipleAVs").dialog("close");
        OpenFeatureCreatePopUp(refresh,url);

    }

    function OpenFeatureCreatePopUp(refresh,url) {
        var strFromCreateAV = "FromCreateAV";

        $("#divFeatureCreateDialog").dialog({
            width: 625, height:380, modal: true, resizable: false, position: ['top', 100]
        });
        $("#ifFeatureCreateDialog").attr("width", "600px");
        $("#ifFeatureCreateDialog").attr("height", "180px");
        $("#ifFeatureCreateDialog").attr("src", "/IPulsar/Features/FeatureCreate.aspx?FromCreateAVPage=FromCreateAV");
        $("#divFeatureCreateDialog").dialog("option", "title", "Create Feature");
        $("#divFeatureCreateDialog").dialog("open");
        //Create Feature
        /*window.showModalDialog(url, "","dialogWidth:600px;dialogHeight:180px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
     }
    function CloseFeatureCreatePopUp() {
        $("#divFeatureCreateDialog").dialog("close");
        return false;
    }


    function OpenFeatureProperties(url) {        
        //OpenPopUp(url, "800", "1250", "Feature Properties", true, false);
        var strFromCreateAV = "FromCreateAV";
        var title = "";
        var propertypageurl = "/IPulsar/Features/" + url + "&FromCreateAVPage=" + strFromCreateAV;
        //if properties page is accessed through grid
        if (url.indexOf("&DeliveryType=") > 0) {
            propertypageurl = url.substring(0, url.indexOf("&DeliveryType"));
            if (url.indexOf("AMO") > -1) {
                propertypageurl = "AMOFeatureProperties.aspx?" + propertypageurl;
                title = "AMO Feature Properties";
            }
            else {
                propertypageurl = "FeatureProperties.aspx?" + propertypageurl;
                title = "Feature Properties";
            }
        }
        else
        {
            title = "Feature Properties";
        }
        //OpenPopUp(propertypageurl, "700", "1100", title, true, false, true, "divFeaturePropertiesDialog", "ifFeaturePropertiesDialog");

        $("#divFeaturePropertiesDialog").dialog({width: adjustWidth(90), height: adjustHeight(95), modal: true, resizable: true, position: ['top', 100]});
        $("#ifFeaturePropertiesDialog").attr("width", "100%");
        $("#ifFeaturePropertiesDialog").attr("height", "100%");
        $("#ifFeaturePropertiesDialog").attr("src", propertypageurl);
        $("#divFeaturePropertiesDialog").dialog("option", "title", "Feature Properties");
        $("#divFeaturePropertiesDialog").dialog("open");

        //Feature Properties
        /*window.showModalDialog(propertypageurl, "","dialogWidth:700px;dialogHeight:1100px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
    }

    function adjustWidth(percent) {
        return document.documentElement.offsetWidth * (percent/100);
    }

    function adjustHeight(percent) {
        return (document.documentElement.offsetHeight * (percent/100)); 
    }

    function CloseFeaturePropertiesPopUp(refreshgrid,FeatureID) {
        $("#divFeaturePropertiesDialog").dialog("close");
        var productversionID = $("#txtID").val();
        var brandID = $("#hidProductBrandID").val();
        var currentuserID = $("#txtUser").val();
        var IsDesktop = $("#hidIsDesktop").val();

        ShowAddMultipleAVsDialog(brandID, IsDesktop, currentuserID, productversionID)

        return false;
    }

    // --------------------------END COMBINED AV FUNCTIONALITY-------------------------

    function CloseBasePartInformationDialog(refresh) {
        $("#divBasePartInformation").dialog("close");
        if (refresh == "true") {
            window.location.reload();
        }
    }

    String.prototype.trim = function() {
        return this.replace(/^\s+|\s+$/g, "");
    }

    String.prototype.ltrim = function() {
        return this.replace(/^\s+/, "");
    }

    String.prototype.rtrim = function() {
        return this.replace(/\s+$/, "");
    }

    function AddAV(PVID, BID, CurrentUserId) {
        //alert Task 9773
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);

        var IsPC = $("#hidIsPC").val();

        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/avFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID + "&IsPC=" + IsPC, "SCM AV Details", DlgWidth, DlgHeight);
    }

    function LocalizeAVs(PVID, strSeriesSummary, bMultipleSeriesNumbers, BID, UserName, sFusionRequirements) {
        var strID;
        var url;
        
        if(sFusionRequirements){
            url = "<%=AppRoot %>/SupplyChain/LocalizeAVsSeriesNum_Pulsar.aspx?Mode=add&PVID=" + PVID + "&BID=" + BID + "&strSeriesSummary=" + strSeriesSummary + "&UserName=" + UserName;
        }else{
            url = "<%=AppRoot %>/SupplyChain/LocalizeAVsSeriesNumFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID + "&strSeriesSummary=" + strSeriesSummary + "&UserName=" + UserName;
        }
        modalDialog.open({dialogTitle:'Add Localized AVs', dialogURL:''+url+'', dialogHeight:300, dialogWidth:600, dialogResizable:true, dialogDraggable:true}); 
    }

    function OpenLocalizeAVs(strPath){        
        if(strPath == null){
            return;
        }else{
            //close LocalizeAVs open modal dialog: ---
            modalDialog.cancel();

            //configure width and height and then open modal dialog with selected AV
            var DlgWidth = 1100;
            if(adjustWidth(95) < DlgWidth){
                DlgWidth = adjustWidth(95);
            }
            var DlgHeight = adjustHeight(95);
            //var retVal = window.parent.showModalDialog(strPath, "", "dialogWidth:" + DlgWidth + "px;dialogHeight:" + DlgHeight + "px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
            modalDialog.open({dialogTitle:'Add Localized AVs', dialogURL:''+strPath+'', dialogHeight:DlgHeight, dialogWidth:DlgWidth, dialogResizable:true, dialogDraggable:true}); 
        }
    }

    function ReloadLocalizeAVs(retVal){
        if(retVal == "YES"){
            //close LocalizeAVs open modal dialog: ---
            modalDialog.cancel();
            //reload parent page
            window.location.reload();
        }
    }

    function StructureBOM(User, BID, KMAT, BusinessID, PVID) {
        var strID;
        var url = "<%=AppRoot %>/SupplyChain/StructureBOMFrame.asp?Mode=add&User=" + User + "&BID=" + BID + "&KMAT=" + KMAT + "&BusinessID=" + BusinessID + "&PVID=" + PVID;
     modalDialog.open({dialogTitle:'Structure BOM', dialogURL:''+url+'', dialogHeight:250, dialogWidth:410, dialogResizable:true, dialogDraggable:true}); 
    }

    function AvActionScorecard(PVID, UserId) {
        var strID;
        var url = "/ipulsar/Reports/SCM/AvActionScoreCard.aspx?PVID=" + PVID + "&UserID=" + UserId + "&isPopup=true";
        mywindow = window.open(url, "AV_Action_Score_card", "location=0,width=1000,height=600,resizable=1", true);

    }

    function UploadAvData(BID, PVID, UserName) {
        var url = "/ipulsar/SCM/HPRPN_Import.aspx?PVID=" + PVID + "&ProductBrandID=" + BID;
        ShowRPNDialog(url, "Import AV Data", 780, 350); //780, 350
        
    }

    function ImportPN(PVID) {      
        var url = "/Pulsar/Import/ImportPN?isASCM=false&PVID=" + PVID;
        mywindow = window.open(url, "Import_PN_from_S4", "location=0,width=1000,height=600,resizable=1", true);
    }

    function AddSAs() {
        var strID;
        strID = window.open("<%=AppRoot %>/MobileSE/SubAssembly.asp?SAType=223&Business=100&Family=100", "", "width=1200,height=1000,toolbar=0,resizable=1,scrollbars=1")
    }
    
    function QuickSearch() {
        var url = "<%=AppRoot %>/MobileSE/QuickSearch_Pulsar.asp"
        modalDialog.open({dialogTitle:'Quick Search', dialogURL:''+url+'', dialogHeight:175, dialogWidth:350, dialogResizable:true, dialogDraggable:true});

    }

    function EditKMAT(PVID, BID) {
        modalDialog.open({dialogTitle:'Edit Program Data', dialogURL:'<%=AppRoot %>/SupplyChain/kmatFrame.asp?Mode=add&PVID=' + PVID + '&BID=' + BID+'', dialogHeight:700, dialogWidth:650, dialogResizable:true, dialogDraggable:true});

    }

    function showOrphanedAvReport(BID) {
        var strID
        strID = window.open("<%=AppRoot %>/SupplyChain/avOrphansFrame.asp?BID=" + BID, "OrphanedAVReport", "width=450,height=500,toolbar=0,resizable=1")
    }
     function showUnreleasedAvReport(BID) {
            var strID
            strID = window.open("/pulsar/Scm/SCMUnreleasedAV?ProductBrandID=" + BID, "UnReleasedAVs", "width=650,height=500,toolbar=0,resizable=1,scrollbars=1")
    }

    function showUnpublishedAVs(BID) {
        var url;

        url = "<%= AppRoot %>/SupplyChain/UnpublishedAVsFrame.asp?BID=" + BID, "";
       modalDialog.open({dialogTitle:'Hidden AVs', dialogURL:''+url+'', dialogHeight:560, dialogWidth:700, dialogResizable:true, dialogDraggable:true}); 

    }

    function showAVsWithMissingData(BID, IsPC) {

      modalDialog.open({dialogTitle:'AVs Missing Corporate Data', dialogURL:'<%= AppRoot %>/SupplyChain/AvsMissingDataFrame.asp?BID=' + BID + '&IsPC=' + IsPC + '', dialogHeight:600, dialogWidth:(GetWindowSize('width')), dialogResizable:true, dialogDraggable:true});

    }

    function showAVsWithMissingDelRoot(BID, PVID) {
       modalDialog.open({dialogTitle:'AVs Missing Deliverable Root', dialogURL:'<%= AppRoot %>/SupplyChain/AvsMissingDelRootFrame.asp?BID=' + BID + '&PVID=' + PVID + '',  dialogHeight:560, dialogWidth:700, dialogResizable:true, dialogDraggable:true});

    }

    function getCookieValue(cookieName) {
        var cookieValue = document.cookie;

        var cookieStartsAt = cookieValue.indexOf(" " + cookieName + "=");
        if (cookieStartsAt == -1) {
            cookieStartsAt = cookieValue.indexOf(cookieName + "=");
        }
        if (cookieStartsAt == -1) {
            cookieValue = "";
        }
        else {
            cookieStartsAt = cookieValue.indexOf("=", cookieStartsAt) + 1;
            var cookieEndsAt = cookieValue.indexOf(";", cookieStartsAt);
            if (cookieEndsAt == -1) {
                cookieEndsAt = cookieValue.length;
            }
            cookieValue = unescape(cookieValue.substring(cookieStartsAt, cookieEndsAt));
        }
        return cookieValue;
    }

    function SelectTab(strStep, blnLoad) {
        var i;
        var expireDate = new Date();

        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "PMTab=" + strStep + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";

        CurrentState = strStep;
        //add List to query string so when two browsers are open the same time with different production will not show "not display"
        window.location.replace("<%=AppRoot %>/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&List=" + strStep);
    }

    function getCookieValue(cookieName) {
        var cookieValue = document.cookie;
        var cookieStartsAt = cookieValue.indexOf(" " + cookieName + "=");
        if (cookieStartsAt == -1) {
            cookieStartsAt = cookieValue.indexOf(cookieName + "=");
        }
        if (cookieStartsAt == -1) {
            cookieValue = "";
        }
        else {
            cookieStartsAt = cookieValue.indexOf("=", cookieStartsAt) + 1;
            var cookieEndsAt = cookieValue.indexOf(";", cookieStartsAt);
            if (cookieEndsAt == -1) {
                cookieEndsAt = cookieValue.length;
            }
            cookieValue = unescape(cookieValue.substring(cookieStartsAt, cookieEndsAt));
        }
        return cookieValue;
    }

    function AddFavorites(strID) {
        var strFavorites;
        var FoundAt;
        var FavCount;

        AddingID = "P" + strID;
        strID = "P" + strID;

        strFavorites = txtFavs.value;
        FavCount = txtFavCount.value;
        if (FavCount == "" || FavCount == "NaN")
            FavCount = 0;
        FavCount = Number(FavCount) + 1;
        FoundAt = strFavorites.indexOf(strID + ",");
        if (!FoundAt > -1) {
            strFavorites = strFavorites + strID + ","
            txtFavCount.value = String(FavCount);
            txtFavs.value = strFavorites;
            var favoriteItemId = strID.replace("P", "");
            var favoriteItemName = $.trim($("#productNameTitle").text().replace('Information (Pulsar)', '')
                .replace('Information (Legacy)', '')
                .replace('SERVICE Information (Pulsar)', '')
                .replace('SERVICE Information (Legacy)', ''));
            var userFavorite = {
                FavoriteType: favoriteTypeProduct,
                ItemId: favoriteItemId,
                ItemName: favoriteItemName,
                ItemLink: "/Excalibur/Excalibur.asp?path=%2FExcalibur%2Fpmview.asp%3FClass%3D1%26ID%3D" + favoriteItemId,
                MenuItemDisplayMethod: menuItemDisplayMethodNone
            };

            addItemtoUserFavorites("/Pulsar2/api/PulsarUser/AddUserFavorite",
                                    userFavorite,
                                    function (){
                                        RFLink.style.display = "";
                                        AFLink.style.display = "none";
                                    },
                                    favoritesOperationFailed);
        }
    }

    function myCallback(returnstring) {
        if (returnstring == "1") {
            window.parent.frames("LeftWindow").location.replace("<%=AppRoot %>/tree.asp?Prog=1&ID=" + AddingID);
            RFLink.style.display = "";
            AFLink.style.display = "none";
        }
    }

	function ShowProperties(DisplayedID, FusionRequirements) {
        var strID
		strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?Commodity=0&ID=" + DisplayedID + "&Pulsar=" + FusionRequirements, "", "dialogWidth:800px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    function ShowCommodityProperties(DisplayedID, Type) {
        var strID
        if (Type == 1)
            strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?Commodity=1&ID=" + DisplayedID, "", "dialogWidth:675px;dialogHeight:350px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        else if (Type == 2)
            strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?FactoryEngineer=1&ID=" + DisplayedID, "", "dialogWidth:475px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        else if (Type == 3)
            strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?Accessory=1&ID=" + DisplayedID, "", "dialogWidth:475px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        else if (Type == 4)
            strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?HWPM=1&ID=" + DisplayedID, "", "dialogWidth:475px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")

        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }


    function RemoveFavorites(strID) {
        var strFavorites;
        var FoundAt;
        var FavCount;

        AddingID = "P" + strID;
        strID = "P" + strID;

        strFavorites = txtFavs.value;
        FavCount = txtFavCount.value;
        if (FavCount == "" || FavCount == "NaN")
            FavCount = 0;
        FavCount = Number(FavCount) - 1;
        if (FavCount < 0)
            FavCount = 0;

        FoundAt = strFavorites.indexOf(strID + ",");
        if (!FoundAt > -1) {
            strFavorites = strFavorites.replace(strID + ",", "")
            txtFavCount.value = String(FavCount);
            txtFavs.value = strFavorites;
            //jsrsExecute("<%=AppRoot %>/FavoritesRSupdate.asp", myCallback2, "UpdateFavs", Array(strFavorites, String(FavCount), txtUser.value));
            var favoriteItemId = strID.replace("P", "");
            var favoriteItemName = $.trim($("#productNameTitle").text().replace('Information (Pulsar)', '')
                .replace('Information (Legacy)', '')
                .replace('SERVICE Information (Pulsar)', '')
                .replace('SERVICE Information (Legacy)', ''));
            var userFavorite = {
                FavoriteType: favoriteTypeProduct,
                ItemId: favoriteItemId,
                ItemName: favoriteItemName
            };

            removeItemFromUserFavorites("/Pulsar2/api/PulsarUser/DeleteUserFavorite",
                                        userFavorite,
                                        function (){
                                            RFLink.style.display = "none";
                                            AFLink.style.display = "";
                                        },
                                        favoritesOperationFailed);
        }
    }

    function myCallback2(returnstring) {
        if (returnstring == "1") {
            window.parent.frames("LeftWindow").location.replace("<%=AppRoot %>/tree.asp?Prog=1");
            RFLink.style.display = "none";
            AFLink.style.display = "";
        }
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

    function BrandLink_onClick(ProductBrandID) {
        window.location.replace("<%=AppRoot %>/SupplyChain/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&BID=" + ProductBrandID);
    }

    function FRMO() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }

        node.style.backgroundColor = "lightseagreen";
        node.style.cursor = "hand";
        window.status = node.pvid + ":" + node.bid;
    }

    function FRMOut() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }

        node.style.backgroundColor = "lightsteelblue";
    }

    function FROC() {
        var node = window.event.srcElement;
        if (node.nodeName.toLowerCase() !== 'td')
            return;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }

        if (node.fcid == "0") {
        	alert("You cannot edit the TBD category");
        } else {
            var pvid = $(node).attr('pvid'), scmid = $(node).attr('fcid'),bid = $(node).attr('bid');
            ShowSCMCategory(pvid, scmid, bid);
        }
    }
    

    function ShowSCMCategory(ProductVersionID, SCMCategoryID, ProductBrandID) {
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);

        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/SCMCatFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&SCMID=" + SCMCategoryID + "&BID=" + ProductBrandID, "SCM Category Details", DlgWidth, DlgHeight);               
    }


    function CloseSCMCategoryPropertiesDialog(strID) {
        $("#modalDialog").attr("src", "");
        $("#iframeDialog").dialog("close");

        if (strID == null)
            return;
        if (strID.Refresh == "1")
        { 
            document.getElementById("tdCategoryMD"+ strID.SCMCategoryID).innerHTML = strID.CategoryMD;
            document.getElementById("tdCategoryRules" + strID.SCMCategoryID).innerHTML = strID.CategoryRules;
            document.getElementById("tdCategoryRuleSyntax" + strID.SCMCategoryID).innerHTML = strID.CategoryRuleSyntax;
            document.getElementById("tdCatMin" + strID.SCMCategoryID).innerHTML = (strID.CatMin.length > 0)  ? 'MIN=' + strID.CatMin : strID.CatMin;
            document.getElementById("tdCatMax" + strID.SCMCategoryID).innerHTML = (strID.CatMax.length > 0)  ? 'MAX=' + strID.CatMax : strID.CatMax;
        }
        $("#iframeDialog").dialog('destroy'); 
    }
    
    function AVMO() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }
        node.style.color = "red";
        node.style.cursor = "hand";
        window.status = "PVID:" + node.pvid + "BID:" + node.bid + "AVID:" + node.avid;
    }


    function AVMOut() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }

        node.style.color = "black";
    }

    function AVOC() {
        var node = window.event.srcElement;
        if (node.nodeName.toLowerCase() !== 'td')
            return;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }
        var pvid = $(node).attr('pvid'), avid = $(node).attr('avid'), bid = $(node).attr('bid'), mktDesc = $(node).attr('mktDesc'), mktDescPMG = $(node).attr('mktDescPMG');
        ShowAvDetails(pvid, avid, bid, mktDesc, mktDescPMG);
    }

    function AVMD(ProductVersionID, AvDetailID) {
        if (event.button == 2) {
            RtClickMenu();
            return;
        }
    }

    function ShowAvDetails(ProductVersionID, AvDetailID, ProductBrandID, mktDesc, mktDescPMG) {
        var objReturn;
        var SCMCat = 0;
        
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);

        SCMCat = document.getElementById("SCMCat" + AvDetailID).innerHTML.trim();

        var IsPC = $("#hidIsPC").val();

        var title = "SCM AV Details - " + mktDesc + " (" + mktDescPMG + ")";

        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/avFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&UserID=" + txtUser.value + "&SCMCat=" + SCMCat + "&IsPC=" + IsPC, title, DlgWidth, DlgHeight);               
    }
    function ShowAvDetailsForClone(ProductVersionID, AvDetailID, ProductBrandID) {
        var objReturn;
        var SCMCat = 0;
        
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);

        var IsPC = $("#hidIsPC").val();

        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/avFrame.asp?Mode=clone&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&UserID=" + txtUser.value + "&SCMCat=" + SCMCat + "&IsPC=" + IsPC, "SCM AV Details", DlgWidth, DlgHeight);               
    }


    function ReloadAVData(avid,AvNo,GpgDescription,MarketingDesc,MarketingDescPMG,ConfigRules,RuleSyntax, TextAvId,Group1,Group2,Group3,Group4,Group5,Group6, Group7, Ids_Skus,Ids_Cto,Rcto_Skus,Rcto_Cto,Weight,GSEndDt,ProductLine,PBID, RTP,PAAD, SA, GA, EOM, BSAMB, Releases)
    {
        document.getElementById("divAVNo" + avid).innerHTML = AvNo;
        if (document.getElementById("Gdsc" + avid) != null)
            document.getElementById("Gdsc" + avid).innerHTML = GpgDescription;
        else
            document.getElementById("Gdsc1" + avid).innerHTML = GpgDescription;
        if (document.getElementById("Mdsc" + avid) != null)
            document.getElementById("Mdsc" + avid).innerHTML = MarketingDesc;
        else
            document.getElementById("Mdsc1" + avid).innerHTML = MarketingDesc;
        if (document.getElementById("MdscPMG" + avid) != null)
            document.getElementById("MdscPMG" + avid).innerHTML = MarketingDescPMG;
        else
            document.getElementById("MdscPMG1" + avid).innerHTML = MarketingDescPMG;
        document.getElementById("rules" + avid).innerHTML = ConfigRules;
        document.getElementById("rulesX" + avid).innerHTML = RuleSyntax;
        document.getElementById("tdAvId" + avid).innerHTML = TextAvId;
        document.getElementById("g1" + avid).innerHTML = Group1;
        document.getElementById("g2" + avid).innerHTML = Group2;
        document.getElementById("g3" + avid).innerHTML = Group3;
        document.getElementById("g4" + avid).innerHTML = Group4; 
        document.getElementById("g5" + avid).innerHTML = Group5;
        document.getElementById("g6" + avid).innerHTML = Group6;
        document.getElementById("g7" + avid).innerHTML = Group7;
        document.getElementById("ids" + avid).innerHTML = Ids_Skus;
		document.getElementById("idsc" + avid).innerHTML = Ids_Cto;
		document.getElementById("rcto" + avid).innerHTML = Rcto_Skus;
        document.getElementById("rctoc" + avid).innerHTML = Rcto_Cto;
        
        <% If (sProdVersionBSAMFlag = "True") Then %>
            document.getElementById("tdBSAMB" + avid).innerHTML = BSAMB;
        <% end if %>
        document.getElementById("wt" + avid).innerHTML = Weight;
        document.getElementById("GSDt" + avid).innerHTML = GSEndDt;        

        if (document.getElementById("PL" + avid) != null)
            document.getElementById("PL" + avid).innerHTML = ProductLine;

        if(document.getElementById("CPLdt"+ avid) != null)
            document.getElementById("CPLdt" + avid).innerHTML = SA;

        if (document.getElementById("RTPdt" + avid) != null)
            document.getElementById("RTPdt" + avid).innerHTML = RTP;

        if (document.getElementById("RASdt" + avid) != null)
            document.getElementById("RASdt" + avid).innerHTML = EOM;

        if (document.getElementById("tdPAADDate" + avid) != null)
            document.getElementById("tdPAADDate" + avid).innerHTML = PAAD;

        if(document.getElementById("tdGADate"+ avid) != null)
            document.getElementById("tdGADate" + avid).innerHTML = GA;

        document.getElementById("rel" + avid).innerHTML = Releases;
    }
    function ReloadAVDataFromMkt(AvID,GPGDescription,MarketingDesc,RTPDate,RASDisDate,PAADDate,SADate,GADate)
    {
        if (document.getElementById("Gdsc" + AvID) != null)
            document.getElementById("Gdsc" + AvID).innerHTML = GPGDescription;
        else
            document.getElementById("Gdsc1" + AvID).innerHTML = GPGDescription;
        if (document.getElementById("Mdsc" + AvID) != null)
            document.getElementById("Mdsc" + AvID).innerHTML = MarketingDesc;
        else
            document.getElementById("Mdsc1" + AvID).innerHTML = MarketingDesc;
        if(document.getElementById("CPLdt"+ AvID) != null)
            document.getElementById("CPLdt" + AvID).innerHTML = SADate;
        if (document.getElementById("RTPdt" + AvID) != null)
            document.getElementById("RTPdt" + AvID).innerHTML = RTPDate;
        if (document.getElementById("RASdt" + AvID) != null)
            document.getElementById("RASdt" + AvID).innerHTML = RASDisDate;
        if (document.getElementById("tdPAADDate" + AvID) != null)
            document.getElementById("tdPAADDate" + AvID).innerHTML = PAADDate;        
        if(document.getElementById("tdGADate"+ AvID) != null)
            document.getElementById("tdGADate" + AvID).innerHTML = GADate;
    }
       
    var oPopup = window.createPopup();
    oPopup.document.createStyleSheet("<%=AppRoot %>/style/menu.css");
    var _avDetailID;
    var _productVersionID;
    var _productBrandID;
    var _ParentId;
    var _ParentStatus;
    function RtClickMenu() {
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;

        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }

        _avDetailID = node.avid;
        _productVersionID = node.pvid;
        _productBrandID = node.bid;
        _ParentId =node.parentID;
        _ParentStatus ="";
        _bSCMPublished =node.bSCMPublished;
        oPopup.document.body.innerHTML = window.PopUpMenu.innerHTML;

        oPopup.show(lefter, topper, 150, 85, document.body);

        //for a localization AV, find it's base part status, if base part is 'O', do not allow re-activate			
        if (node.status == "O" && _ParentId != 0)
        {
            $("#GridViewContainer tr").each(function (i, avrow) {                  
            if (avrow.avid == _ParentId)
            {_ParentStatus=avrow.status;}
        });
        }
        if (node.status == "A")
        {
            oPopup.document.body.all["activate"].style.display = "none";    
            oPopup.document.body.all["unhide"].style.display = "none";
        }

        if  (node.status == "O")
        {  
            if  (_ParentStatus=="O")
            {  oPopup.document.body.all["activate"].style.display = "none";
            oPopup.document.body.all["Li2"].style.display = "none";
            oPopup.document.body.all["Li3"].style.display = "none";
            }
            oPopup.document.body.all["obsolete"].style.display = "none";
            oPopup.document.body.all["Li2"].style.display = "none";
            oPopup.document.body.all["hide"].style.display = "none";
            oPopup.document.body.all["unhide"].style.display = "none";
        }
        if (node.status == "H")  
        {
            oPopup.document.body.all["hide"].style.display = "none";   
            oPopup.document.body.all["activate"].style.display = "none";
        }
        
        //if av is published, do not show the hide option:
        if (_bSCMPublished =="1" || _bSCMPublished=="True")
        {   oPopup.document.body.all["hide"].style.display = "none";
            oPopup.document.body.all["Li2"].style.display = "none";
        }

        //Adjust window size
        if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
            NewHeight = oPopup.document.body.scrollHeight;
            NewWidth = oPopup.document.body.scrollWidth;
            oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
        }

    }

    function MenuProperties() {
        ShowAvDetails(_productVersionID, _avDetailID, _productBrandID);
    }

    function MenuSetObsolete(bIsPc) {
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'O',bIsPc);
    }

    function MenuSetActive(bIsPc) {        
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'A',bIsPc);
    }

    function MenuSetHidden(bIsPc) {      
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'H',bIsPc);
    }

    function MenuDelete() {
        DeleteAv(_productVersionID, _avDetailID, _productBrandID);
    }

    function ViewAvailMultipleAvs(ProductVersionID, ProductBrandID, bIsPc) {        
        var AvDetailIDs = "";
        var btnAvailMultipleAvs = document.getElementById(window.event.srcElement.id);
        var ParentCategory = btnAvailMultipleAvs.id.substring(19);

        $('input.' + ParentCategory + ':checkbox:checked').each(function () {
            if(AvDetailIDs == "")
                AvDetailIDs = $(this).attr('name').substring(5); 
            else
                AvDetailIDs += "," + $(this).attr('name').substring(5);
        });        
        
        if(AvDetailIDs.length > 0){
            ViewAvailability(ProductVersionID, ProductBrandID, AvDetailIDs, bIsPc);            
        }
        else {
            alert("Please check the AVs which you would like to view availability");
        }
        
    }

    function DeleteMultipleAvs(ProductVersionID, ProductBrandID) {
        var strID;
        var response = confirm("Are you sure you want to delete the selected records?");
        if (response) {
            var i;
            var checkBoxes = document.getElementsByTagName("input");
            var btnDelete = document.getElementById(window.event.srcElement.id);
            var ParentCategory = btnDelete.id.substring(9);
            for (i = 0; i < checkBoxes.length; i++) {
                if ((checkBoxes[i].id == "chkAv" + ParentCategory) && (checkBoxes[i].checked == true)) {
                    var AvDetailID = checkBoxes[i].name.substring(5);
                    strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/avDelete.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                    if (strID!="last BU") 
                    {var row = document.getElementById("AV" + AvDetailID);
                        if ((row.className == 'O') || (hidLastPublishDt.value == '') || (hidBusinessID.value != 1))
                            row.style.display = "none";
                        else
                            row.className = 'O';
                    }
                }
            }

            var base = document.getElementById("chkAll" + ParentCategory);
            base.checked = false;
            base.indeterminate = false;
        }
    }
    
    function MenuEditMrDates() {
    EditMrDates(_productVersionID, _productBrandID, _avDetailID);
    }

    function SetAvStatus(ProductVersionID, AvDetailID, ProductBrandID, Status,bIsPc) {        
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/avSetStatus.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&Status=" + Status, "", "dialogWidth:300px;dialogHeight:120px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        if(bIsPc && strID > 0)
        {
			var row = document.getElementById("AV" + AvDetailID);
			row.className = Status;
			row.status=Status;

			if (Status=='A'){
                
			    var ajaxurl = "<%=AppRoot %>/SupplyChain/GetAVDates.asp?AvDetailId=" + _avDetailID;
			    $.ajax({
			        url: ajaxurl,
			        type: "GET",
			        async: false,
			        success: function (data) {
			            var arrData = data.split(';');
			            for (var item in arrData){
			                var arrItem = item.split(':');
			                switch(arrItem(0)){
			                    case "CPLBlindDt":
			                        document.getElementById("CPLdt" + AvDetailID).innerHTML = arrItem(1);
			                        break;
			                    case "GeneralAvailDt":
			                        document.getElementById("tdGADate" + AvDetailID).innerHTML = arrItem(1);
			                        break;
			                    case "RASDiscontinueDt":
			                        document.getElementById("RASdt" + AvDetailID).innerHTML = arrItem(1);
			                        break;
			                    case "PHWebDate":
			                        document.getElementById("tdPAADDate" + AvDetailID).innerHTML = arrItem(1);
			                        break;
			                }
			            }			            
			        },
			        error: function (xhr, status, error) {
			            alert(error);
			        }
			    });
			}
            //if status is o or H , mark teh status the same for the localization avs			
			if (Status=='O' || Status == 'H')
			{    $("#GridViewContainer tr").each(function (i, avrow) {                  
			        if (avrow.parentID == AvDetailID)
			        {avrow.className = Status;
			            avrow.status=Status;}
			    });
			}
		}
    }

    function calcDates() {
        var q = new Date();
        var m = q.getMonth();
        var d = q.getDate();
        var y = q.getFullYear();
        var Today = new Date(y, m, d);

        if (isDate($("#hidRTP").val())) {

            var RTPDate = new Date($("#hidRTP").val());
            var GeneralAvailDt = RTPDate;
            var monday = getMonday(RTPDate);
            var firstDay = new Date(GeneralAvailDt.getFullYear(), GeneralAvailDt.getMonth(), 1);
            var PAADDate = new Date((monday.getMonth() + 1) + '/' + monday.getDate() + '/' + monday.getFullYear());
            var BlindDate;
            BlindDate = new Date(PAADDate.getFullYear(), (PAADDate.getMonth()-1), 1);

            $("#hidSA").val(((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear()));
            $("#hidGA").val(($("#hidRTP").val()));
            $("#hidPAAD").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
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


    function DeleteAv(ProductVersionID, AvDetailID, ProductBrandID) {
    var strID;
    var response = confirm("Are you sure you want to delete this record?");
    if (response) {
        strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/avDelete.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID, "", "dialogTop:0;dialogLeft:0;dialogWidth:200px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        
        if (strID!="last BU") 
        {   var row = document.getElementById("AV" + AvDetailID);
            if ((row.className == 'O') || (hidLastPublishDt.value == '') || (hidBusinessID.value != 1))
                row.style.display = "none";
            else
                row.className = 'O';
        }
        }
    }

    function ScmPublishRollback(ProductBrandID, UserName, IsProgramCoordinator) {
        var strID;
        if (IsProgramCoordinator == "True")
        {
            var response = confirm("You are about to Permanently delete all records of the most recent SCM publish. \n Do you want to continue?");
            if (response) {
                var parameters = "function=RollbackSCM&PBID=" + ProductBrandID + "&UserName=" + UserName;
                var request = null;
                //Initialize the AJAX variable.
                if (window.XMLHttpRequest) {        //Are we working with mozilla
                    request = new XMLHttpRequest(); //Yes -- this is mozilla.
                } else { //Not Mozilla, must be IE
                    request = new ActiveXObject("Microsoft.XMLHTTP"); 
                } //End setup Ajax.
                request.open("POST", "<%=AppRoot %>/SupplyChain/SCM_Rollback.asp", false);
                request.setRequestHeader("Content-type","application/x-www-form-urlencoded");
                request.send(parameters);
                if (request.responseText == '') {
                    document.location.reload();
                }
            }
        }
        else
        {
            alert("You do not have permission to roll back");
        }
    }

    function ShowObsolete(value) {
        var expireDate = new Date();

        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "ShowObsolete=" + value + ";expires=" + expireDate.toGMTString() + ";";

        window.location.reload(true);
    }

    function ShowReport(value) {
        var expireDate = new Date();

        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "ShowReport=" + value + ";expires=" + expireDate.toGMTString() + ";";

        window.location.reload(true);
    }

    function EditMrDates(ProductVersionID, ProductBrandID, AvDetailID) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/mrDatesFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&AVID=" + AvDetailID, "", "dialogWidth:325px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    function ImportAv(ProductVersionID, ProductBrandID) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/pasteFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:410px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        document.location.reload();
    }

    function CopyScm(ProductVersionID, ProductBrandID) {
    var strID;
    strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/copyFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    document.location.reload();
    }

    function LinkAv(ProductVersionID, ProductBrandID, AvID) {
    var sMode = null;
    var strID = null;
    if (AvID == null) {
    sMode = "LinkFrom";
    strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/linkFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&Function=" + sMode, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }
    else {
    sMode = "LinkTo";
    strID = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/linkFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&Function=" + sMode, "", "dialogWidth:500px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    document.location.reload();

    }

	function PublishScmOld(ProductVersionID, ProductBrandID) {
        if (ValidateData(ProductBrandID) == 1) {
            var answer = confirm("Click OK to continue Publishing the SCM");
            if (answer)
                location.href = "/iPulsar/ExcelExport/SCM.aspx?BID=" + ProductBrandID + "&PVID=" + ProductVersionID + "&Publish=True";
        }
    }

    function PublishScmNew(ProductVersionID, ProductBrandID) {
        //Check Features status before proceeding
        if($("#hidFeatureList").val() == "")
        {
            if (ValidateData(ProductBrandID) == 1) {
                ShowSCMPublishDialog("<%=AppRoot %>/SupplyChain/SCMPublish.asp?ProductBrandID=" + ProductBrandID + "&PVID=" + ProductVersionID , "Publish SCM", 780, 600);
            }
        }       
        else
        {
            var str = $("#hidFeatureList").val();
            while( str.indexOf(",") > -1)
            {
                str = str.replace(",", "\n");
            }
            alert("Please make sure the following Feature(s) are active before continuing :\n" + str);
        }
    }

    function OpenSCMReportList(ProductVersionID, ProductBrandID) {  
        var DlgWidth = adjustWidth(90);
        var DlgHeight = adjustHeight(95);

        $("#divOpenSCMReport").dialog({ width: DlgWidth, height: DlgHeight });
        $("#ifOpenSCMReport").attr("width", "98%");
        $("#ifOpenSCMReport").attr("height", "98%");
        $("#ifOpenSCMReport").attr("src", '<%=AppRoot %>/SupplyChain/ViewAllPublishedReports.asp?PVID=' + ProductVersionID + '&ProductBrandID=' + ProductBrandID);
        $("#divOpenSCMReport").dialog("option", "title", "Published SCM/Program Matrix Reports");
        $("#divOpenSCMReport").dialog("open");
    }

    function EOSReport(ProductBrandID)
    {
        location.href = "/ipulsar/Reports/SCM/SCM_EOS_Report.aspx?PBID=" + ProductBrandID;
    }

    function CloseSCMReportList()
    {
        $("#divOpenSCMReport").dialog("close");
    }

    function ShowFeatureSelectDialog(QueryString, Title, DlgWidth, DlgHeight) {
        var DlgWidth = adjustWidth(90);
        var DlgHeight = adjustHeight(95);
       
        //the dialog of feature do not reload after we close it. the code below is doing all that and we are all should call one function to open diag in this page.
        OpenPopUp(QueryString, DlgHeight, DlgWidth, Title, false, false, true, "divMultipleAVs", "ifMultipleAVs")
        
    }

    function ClosePopUpViewFromAvDetail_Features(Refresh,FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, aliasid, abbreviation, platform) {        
        if (document.getElementById('modalDialog').contentWindow != null)
        {
            document.getElementById('modalDialog').contentWindow.ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, aliasid, abbreviation, platform);
        }
        $("#divMultipleAVs").dialog("close");      
    }

    function OpenAddExistingSharedAV(url, width, height){
        OpenPopUp(url, height, width, "Add Existing Shared AV", true, false, true, "divAddExistingAvAsShared", "ifAddExistingAvAsShared")
    }

    function ShowAddExistingSharedAV(ProductVersionID, ProductBrandID) {
        var url = "../../IPulsar/Admin/SCM/AddExistingSharedAV.aspx?PVID=" + ProductVersionID + "&BID=" + ProductBrandID;
        
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);

        OpenAddExistingSharedAV(url, DlgWidth, DlgHeight);
        /*window.showModalDialog(url, "","dialogWidth:"+DlgWidth+"px;dialogHeight:"+DlgHeight+"px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
	
    }


    function CloseAddExistingAvAsShared()
    {
        $("#divOpenSCMReport").dialog("close");
    }
    function CloseAddExistingSharedAV(refresh)
    {
        $("#ifOpenSCMReport").attr("src", "");
        if(refresh == 1)
        {
            window.location.reload();
        }
        $("#divAddExistingAvAsShared").dialog("close");
    }

    function ValidateData(ProductBrandID) {
        //window.open("<%=AppRoot %>/SupplyChain/SCM_PM_Publish_ValidateData.asp?PBID=" + ProductBrandID);
        var parameters = "function=ValidateData&PBID=" + ProductBrandID;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {        //Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP"); 
        } //End setup Ajax.
        request.open("POST", "<%=AppRoot %>/SupplyChain/SCM_PM_Publish_ValidateData.asp", false);
        request.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        request.send(parameters);
        if (request.responseText != 'Success') {

            if (confirm(request.responseText + " \n Do you want to continue the publish?"))
                return 1
            else
            {
                document.getElementById("chkPublish").checked = false;
                return 0;
            }
		} else {
            return 1;
        }
    }              

    function clearAvHistory(ProductBrandID) {
        var answer = confirm("Are you sure you want to clear the SCM Change Log?");
        if (answer) {
            //jsrsExecute("<%=AppRoot %>/SupplyChain/rs_ClearHistory.asp", clearAvHistoryCallback, "ClearHistory", String(ProductBrandID));
            ajaxurl = "<%=AppRoot %>/SupplyChain/rs_ClearHistory.asp?ProductBrandID=" + ProductBrandID;
            $.ajax({
                url: ajaxurl,
                type: "POST",
                success: function (data) {
                    alert("Change Log Cleared.");
                },
                error: function (xhr, status, error) {
                    alert(error);               
                }

            });
        }
    }

    //************************************************************************
    //Description:  Show Project Matrix dates drop-down in a modal dialog
    //Function:     ExportProjectMatrix();
    //Modified:     Harris, Valerie (9/15/2016) - PBI 23434/Task 24367 
    //Note:         jquery.ui.min, juery, and shared.css files required
    //************************************************************************
    function ExportProjectMatrix(ProductBrandID){
        //Initialize modal form elements: ---
        $("#trCompare").show();
        $("#popup_submit").attr("disabled", false);
        $("#chkNewMatrix").attr("checked", false);
        $("#chkPublish").attr("checked", false);             
               
        if($("#selCompareDt option").length == 0) {
            $("#trCompare").hide();
            $("#chkNewMatrix").attr("checked", true);
        }  
        ajaxurl = "<%=AppRoot %>/SupplyChain/rs_ProgramMatrix_ManufacturingSites.asp?ProductBrandID=" + ProductBrandID;
        $.ajax({
            url: ajaxurl,
            type: "POST",
            success: function (data) {
                document.getElementById("selManufacturingSitesDiv").innerHTML = data;                  
            },
            error: function (xhr, status, error) {
                document.getElementById("selManufacturingSitesDiv").innerHTML = "<font face='verdana' size='1' color='red'>Error while pulling manufacturing sites</font>";               
            }
        });
        modalDialog.open({dialogTitle:'Program Matrix Options', dialogDivID:'modal_programmatrix', dialogHeight:200, dialogWidth:350, dialogResizable:false, dialogDraggable:true});
    }


    function getPublishDatesCallback(returnString) {
        var chkNewMatrix = document.getElementById("chkNewMatrix");
        var trCompare = document.getElementById("trCompare");
        document.getElementById("selCompareDtDiv").innerHTML = returnString;
        document.getElementById("popup_submit").disabled = false;
        if (returnString == "") {
            trCompare.style.display = "none";
            chkNewMatrix.checked = true;
        }
    }

    function SetDefaultDisplay(strList, CurrentUserID) {

        if (window.confirm("Are you sure you want to make this display list the default display that you see each time you view a product?")) {
            ajaxurl = "<%=AppRoot %>/DefaultProductTabRSUpdate.asp?CurrentUserID=" + CurrentUserID + "&List=" + strList;
            $.ajax({
                url: ajaxurl,
                type: "POST",
                success: function (data) {
                    if (data == "1") {
                        window.location.reload(true);
                    } else {
                        alert("Unable to save this tab as the default.");
                    }
                },
                error: function (xhr, status, error) {
                    alert(error);               
                }
            });
        }
    }

    function myTabSetCallback(returnstring) {
        if (returnstring == "1") {
            window.location.reload(true);
        } else {
            alert("Unable to save this tab as the default.");
        }
    }

    function txtSearch_onkeypress() {
        if (window.event.keyCode == 13)
            goSearch_onclick();
    }

    function gochange_onmouseover() {
        window.event.srcElement.style.cursor = "hand"
    }

    function goSearch_onclick() {

        if (txtSearch.value != "") {
            window.location.href = "http://houhpqexcal03.auth.hpicorp.net/MobileSE/Today/find.asp?Find=" + escape(txtSearch.value) + "&Type=Part";
        }
        else {
            window.alert("Please enter a part number first.");
            txtSearch.focus();
        }
    }

    function chkAv_onclick() {
        UpdateBase(window.event.srcElement);
       }

    function UpdateBase(chkClicked) {
        var i;
        var blnAllSame = true;
        var chkAv = document.getElementsByTagName("input");
        for (i = 0; i < chkAv.length; i++) {
            if (chkAv(i).className != "")
                if (chkAv(i).className == chkClicked.className) {
                if ((chkAv(i).checked != chkClicked.checked) || chkAv(i).indeterminate) {
                    blnAllSame = false;
                }
            }
        }
        var base = document.getElementById("chkAll" + chkClicked.value)
        if (blnAllSame) {
            base.indeterminate = false;
            base.checked = chkClicked.checked;
        }
        else {
            base.indeterminate = true;
            base.checked = true;
        }
  
    }

    function chkAll_onclick() {
        var i;
        var checkBoxes = document.getElementsByTagName("input");
        var chkAll = document.getElementById(window.event.srcElement.id);
        var FeatureCategory = window.event.srcElement.name.substring(6);
        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].id == "chkAv" + FeatureCategory) {
                if (checkBoxes[i].indeterminate) {
                    checkBoxes[i].indeterminate = false;
                }
                checkBoxes[i].checked = chkAll.checked;
            }
        }

    }

    function DocKitData(KMAT) {
        var strID;
        var url = "<%=AppRoot %>/SupplyChain/DocKitDataFrame.asp?Mode=add&KMAT=" + KMAT;
        modalDialog.open({ dialogTitle: 'Check Doc Kit Status', dialogURL: '' + url + '', dialogHeight: 700, dialogWidth: 645, dialogResizable: true, dialogDraggable: true });        

    }

    function FilterSCMByCategory(BusinessID, BID, UserID, PVID) {
        var Categories = document.getElementById("hidSCMCategories");
        var NoLocalization = document.getElementById("hidSCMNoLocalization");
        var GADateTo = document.getElementById("hidGADateTo");
        var GADateFrom = document.getElementById("hidGADateFrom");
        var SADateTo = document.getElementById("hidSADateTo");
        var SADateFrom = document.getElementById("hidSADateFrom");
        var EMDateTo = document.getElementById("hidEMDateTo");
        var EMDateFrom = document.getElementById("hidEMDateFrom");
        var ReleaseIDs = document.getElementById("hidReleaseIDs");

        var url = "<%=AppRoot %>/SupplyChain/FilterByCategoryFrame.asp?BusinessID=" + BusinessID + "&BID=" + BID + "&UserID=" + UserID + "&PVID=" + PVID + "&Categories=" + Categories.value + "&NoLocalization=" + NoLocalization.value + "&GADateTo=" + GADateTo.value + "&GADateFrom=" + GADateFrom.value + "&SADateTo=" + SADateTo.value + "&SADateFrom=" + SADateFrom.value + "&EMDateTo=" + EMDateTo.value + "&EMDateFrom=" + EMDateFrom.value+ "&ReleaseIDs=" + ReleaseIDs.value;

        //set the global varibles for the filter
        fiterBusinessID = BusinessID;
        filterBID = BID;
        filterUserID = UserID;
        filterPVID = PVID;
        SCM_SCMType = "SCM";
        var DlgWidth = 1100;
        var DlgHeight  = 675;     
        $("#divFilter").dialog({ width: DlgWidth, height: DlgHeight });
        $("#divFilter").dialog("option", "title", "Set Filter");
        $("#ifFilter").attr("src", url);             
        $("#divFilter").dialog("open");

    }
    //define some global varibles for the filter 
    var fiterBusinessID = "";
    var filterBID = "";
    var filterUserID = "";
    var filterPVID = "";
    var SCM_SCMType ="";
    function CloseFilterDialog(retValue) {
        var BusinessID = fiterBusinessID;
        var BID = filterBID;
        var UserID = filterUserID;
        var PVID = filterPVID;
        
        $("#divFilter").dialog("close");
        if (retValue != null) {
            window.location.replace("<%=AppRoot %>/SupplyChain/pmview.asp?ID=" + PVID + "&BID=" + BID + "&" + SCM_SCMType + "Categories=" + retValue.Categories + "&NoLocalization=" + retValue.NoLocalization + "&GADateTo=" + retValue.GADateTo + "&GADateFrom=" + retValue.GADateFrom + "&SADateTo=" + retValue.SADateTo + "&SADateFrom=" + retValue.SADateFrom + "&EMDateTo=" + retValue.EMDateTo + "&EMDateFrom=" + retValue.EMDateFrom + "&ReleaseIDs=" + retValue.ReleaseIDs + "&Class=Arrow1");
        }
    }
    function FilterPMByCategory(BusinessID, BID, UserID, PVID) {
        var Categories = document.getElementById("hidPMCategories")
        var url = "<%=AppRoot %>/SupplyChain/FilterByCategoryFrame.asp?BusinessID=" + BusinessID + "&BID=" + BID + "&UserID=" + UserID + "&PVID=" + PVID + "&Categories=" + Categories.value;

        fiterBusinessID = BusinessID;
        filterBID = BID;
        filterUserID = UserID;
        filterPVID = PVID;
        SCM_SCMType = "PM";
        var DlgWidth = 315;
        var DlgHeight  = 605;     
        $("#divFilter").dialog({ width: DlgWidth, height: DlgHeight });
        $("#divFilter").dialog("option", "title", "Set Filter");
        $("#ifFilter").attr("src", url);             
        $("#divFilter").dialog("open");
    }
   
    function ViewAvailabilityAv(bIsPc)
    {
        oPopup.hide();
        ViewAvailability(_productVersionID, _productBrandID, _avDetailID, bIsPc);
    }

    function ViewAvailability(PVID, BID, AvID, bIsPc){
        var DlgWidth = adjustWidth(90);
        var DlgHeight = adjustHeight(95);        

        ShowBUDialog("/IPulsar/SCM/SCM_BaseUnitAvailability.aspx?PVID=" + PVID + "&BID=" + BID + "&AvID=" + AvID, "SCM Base Unit Availability", DlgWidth, DlgHeight, bIsPc);
    }

    function ShowABTInfoDialog(ProductBrandID, EditABTInformation) {                               
        var DlgWidth = 700;
        var DlgHeight  = 600;
        //var DlgWidth = adjustWidth(300);
        //var DlgHeight = adjustHeight(0); 

        $("#divABTInfo").dialog({ width: DlgWidth, height: DlgHeight });
        $("#divABTInfo").dialog("option", "title", "ABT Information");
        $("#ifABTInfo").attr("src", "/IPulsar/SCM/ABT_Information.aspx?ProductBrandID=" + ProductBrandID + "&CanEdit=" + EditABTInformation);     
        $("#divABTInfo").dialog("open");

    }
	
    function ShowSCMManufacturingSitesDialog(ProductBrandID,BusinessSegmentID) {                 
        var DlgWidth = 800;
        var DlgHeight  = 600;

        $("#divSCMManufacturingSites").dialog({ width: DlgWidth, height: DlgHeight, modal: true });
        $("#divSCMManufacturingSites").dialog("option", "title", "SCM Manufacturing Sites");
        $("#ifSCMManufacturingSites").attr("src", "/IPulsar/SCM/SCM_ManufacturingSites.aspx?ProductBrandID=" + ProductBrandID + '&BusinessSegmentID=' + BusinessSegmentID);     
        $("#divSCMManufacturingSites").dialog("open");
    }
    function ShowChinaGPDialog(ProductBrandID,BusinessSegmentID) {                 
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);
        var DlgHeight = adjustHeight(95); 

        $("#divChinaGP").dialog({ width: DlgWidth, height: DlgHeight, modal: true });
        $("#divChinaGP").dialog("option", "title", "China GP Identifiers");
        $("#ifChinaGP").attr("src", "/IPulsar/SCM/ChinaGP.aspx?ProductBrandID=" + ProductBrandID);     
        $("#divChinaGP").dialog("open");
    }
    
    function CloseChinaGPPopup() {
        ClosePopup("divChinaGP");
        return false;
    }
    function ShowAddMultipleAVsDialog(ProductBrandID, IsDesktop, CurrentUserID, ProductVersionID) {
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95);
       
        if ($("#txtSecondAddAVsOpen").val() == "")
        {
            $("#txtSecondAddAVsOpen").val("1")
        }
        else //to solve the isue that second time the av history list opens, 
            //the pop-up opens higher than told in ie 10 and 11 regilar mode
        {   
            //if ($("#txtIe10_11_RegularMode").val() == "1")
            //DlgHeight = DlgHeight -150
        }

        $("#divMultipleAVs").dialog({ width: DlgWidth, height: DlgHeight, modal: true });
        $("#divMultipleAVs").dialog("option", "title", "Add Feature(s)"); //changed title(for combined AV functionality)
        $("#ifMultipleAVs").attr("width", "100%");
        $("#ifMultipleAVs").attr("height", "100%");
        $("#ifMultipleAVs").attr("src", "/IPulsar/SCM/SCM_AddMultipleFeatures.aspx?MultipleAVs=yes&ProductBrandID=" + ProductBrandID + "&IsDesktop=" + IsDesktop + "&CurrentUserID=" + CurrentUserID + "&ProductVersionID=" + ProductVersionID);     
        $("#divMultipleAVs").dialog("open");

        /*
        //Add Feature(s)
        var url = "/IPulsar/SCM/SCM_AddMultipleFeatures.aspx?MultipleAVs=yes&ProductBrandID=" + ProductBrandID + "&IsDesktop=" + IsDesktop + "&CurrentUserID=" + CurrentUserID + "&ProductVersionID=" + ProductVersionID;
        window.showModalDialog(url, "","dialogWidth:"+DlgWidth+"px;dialogHeight:"+DlgHeight+"px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
        */
    }
//-->
    // Base Part Information
    function ShowBasePartInformationDialog(ProductBrandID) {
        var DlgWidth = 1100;
        if(adjustWidth(95) < DlgWidth)
            DlgWidth = adjustWidth(95);

        var DlgHeight = adjustHeight(95); 
               
        /*$("#divBasePartInformation").dialog({ width: DlgWidth, height: DlgHeight });
        $("#divBasePartInformation").dialog("option", "title", "Base Part Information");
        $("#ifBasePartInformation").attr("src", "/IPulsar/SCM/SCM_BasePartInformation.aspx?ProductBrandID=" + ProductBrandID); 
        $("#divBasePartInformation").dialog("open");*/

        //Base Part Information
        var url = "/IPulsar/SCM/SCM_BasePartInformation.aspx?ProductBrandID=" + ProductBrandID;
        window.showModalDialog(url, "","dialogWidth:"+DlgWidth+"px;dialogHeight:"+DlgHeight+"px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
      
    }

    // ********************************** Pulsar Role-Permission Popup ****************************************//
    function ClosePopup(divID) {
        $('#' + divID).dialog('close');
        return false;
    }

    function loadIframe(iframeName, url) {
        var $iframe = $('#' + iframeName);
        $iframe.attr("width", "100%");
        $iframe.attr("height", "100%");
        if ($iframe.length) {
            $iframe.attr('src', url);
            return false;
        }
        return true;
    }

    function OpenPopUp(link, newHeight, newWidth, title, noScrollBar, hideCloseButton, Resizable, divID, ifrID) {
        var $divPopup = $('#' + divID);
        $divPopup.dialog({
            height: newHeight,
            width: newWidth,
            modal: true,
            title: title,
            resizable: Resizable,
            draggable: true,
            open: function (event, ui) {
                if (hideCloseButton)
                    $(this).parent().children().children('.ui-dialog-titlebar-close').hide();
                else
                    $(this).parent().children().children('.ui-dialog-titlebar-close').show();

                if (noScrollBar)
                    $divPopup.css('overflow', 'hidden');
            },
            close: function (event, ui) {
                //everytime the jquery dialog is closed trigger this event to clear the iframe so when dialogue is called again it will show blank first then load with the url
                $("#" + ifrID).attr("src", "");
            }
        });

        loadIframe(ifrID, link);
    }

    // ********************************** Pulsar Role-Permission Popup ****************************************//
    function showPulsarObjectPermission(div, ifrm) {
        var link = "/IPulsar/Admin/Areas/PulsarObjectPermission.aspx?Title=SCM&PageUrl=" + "Excalibur/SupplyChain/pmview.asp";

        var newHeight = 580;
        var newWidth = 700;
        var title = "What do I need to use the SCM?";
        var noScrollBar = true;
        var hideCloseButton = false;
        var Resizable = true;
        var divID = div;
        var ifrID = ifrm;
        OpenPopUp(link, newHeight, newWidth, title, noScrollBar, hideCloseButton, Resizable, divID, ifrID)  
    }
    // ********************************** Pulsar Role-Permission Popu ****************************************//

    function ClosePulsarObjectPermissionPopup() {
        ClosePopup("divPulsarObjectPermission");
        return false;
    }
        
    function closeModalDialog(bReload){
        modalDialog.cancel(bReload);
    };
    function window_onload() {
        modalDialog.load();
    }

    function chkPublish_Click(ProductBrandID, checked) {
        var ret = 0;
        if (checked) { 
            ret = ValidateData(ProductBrandID);
            if (ret == 1)
                $("#chkXrost").attr("disabled", true);
        }
        else
        $("#chkXrost").removeAttr("disabled");
    }

    function pushToXrost(ProductBrandID, checked) {
        var ret = 0;
        if (checked) {
            ret = ValidateData(ProductBrandID);
            if (ret == 1)
                $("input[name='chkPublish']").attr("disabled", true);
            else
                 $("#chkXrost").attr("checked", false); 
        }
        else
            $("input[name='chkPublish']").removeAttr("disabled");
        
    }
//-->
    </script>
</head>
<body onload="window_onload();">
    <div style="margin-bottom: 1.5em; margin-top: -1em; display: inline-block;">
        <div style="position: absolute; right: 15px;">
            <a href="#hrefPulsarObjectPermission" onclick="showPulsarObjectPermission('divPulsarObjectPermission', 'ifPulsarObjectPermission');">
                <img src="<%=AppRoot%>/images/QuestionMark.gif" style="padding: 0; border: 0;" /></a>
        </div>
    </div>

    <%IF sFusionRequirements = true THEN %>
    <span id="productNameTitle" style="font: bold medium Verdana;"><%= sProductName%> Information (Pulsar)</span><br />
    <br />
    <%ELSE%>
    <span id="productNameTitle" style="font: bold medium Verdana;"><%= sProductName%> Information (Legacy)</span><br />
    <br />
    <%END IF%>


    <%If bIsPc Then Response.Write strNoticetable %>
    <%if (clng(request("ID")) = 344 or clng(request("ID")) = 347 or clng(request("ID")) = 1107) then %>
    <td nowrap id="EditLink" style="display: none"></td>
    <%'malichi 07/19/2016, Product Backlog Item 16765: Marketing role needs permissions to Edit Product (Pulsar product)%>
    <%elseif (blnCanEditProduct) then%>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowProperties_Product(<%=PVID%>,0,1)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font></td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%end if%>
    <span id="loadingProgress"></span>
    <span id="RFLink" style="display: none;"><a href="javascript:RemoveFavorites(<%=PVID%>)"><font face="verdana" size="1">Remove From Favorites</font></a><font face="verdana" size="1" color="black"> | </font></span>
    <span id="AFLink" style="display: none;"><a href="javascript:AddFavorites(<%=PVID%>)"><font face="verdana" size="1">Add To Favorites</font></a><font face="verdana" size="1" color="black"> | </font></span>
    <span id="StatusLink" style="display: none;"><a href="<%=AppRoot %>/Productstatus.asp?Product=<%=sDisplayedProductName%>&ID=<%=PVID%>"><font face="verdana" size="1">Real-Time Status Report</font></a><font face="verdana" size="1" color="black"> | </font></span>
    <%if strDisplayedList <> CurrentUserDefaultTab and trim(strProdType) <> "2" then%>
    <span id="DefaultTabLink"><a href="javascript:SetDefaultDisplay('<%=strDisplayedList%>',<%=CurrentUserID%>)"><font face="verdana" size="1">Set Default List</font></a></span>
    <%end if%>

    <br>
    <br>

    <table id="menubar" class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0"
        cellpadding="2">
        <tr bgcolor="<%=strTitleColor%>">
            <td id="CellDCR" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('DCR',1)">Change&nbsp;Requests</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellAction" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Action',1)">Action&nbsp;Items</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellOTS" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('OTS',1)">Observations</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellAgency" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Agency',1)">Certifications</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellPMR" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('PMR',1)">SMR</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellCalls" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Calls',1)">Service</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellGeneral" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('General',1)">General</a>&nbsp;&nbsp;&nbsp;</font></td>
        </tr>
        <tr bgcolor="<%=strTitleColor%>">
            <td id="CellRequirements" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Requirements',1)">Requirements</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellCountry" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Country',1)">Localization</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellLocal" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Local',1)">Images</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellDeliverables" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Deliverables',1)">Deliverables</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellSchedule" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Schedule',1)">Schedule</a>&nbsp;&nbsp;&nbsp;</font></td>
            <td id="CellSCMb" width="10" bgcolor="wheat"><font size="1" color="black">Supply&nbsp;Chain</font></td>
            <td id="CellDocuments" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Documents',1)">Documents</a>&nbsp;&nbsp;&nbsp;</font></td>
        </tr>
    </table>
    <%

Response.flush

'
' Get the list of Brands for the product.
'
Set cmd = dw.CreateCommAndSP(cn, "usp_GetBrands4Product")
dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, PVID
dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
Set rs = dw.ExecuteCommAndReturnRS(cmd)
rs.Sort = "SortingBrandName" 'sp already sorted by that field, force again here
bFirstWrite = True

If Not rs.EOF Then
		
    %>
    <br>
    <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
        <tr>
            <td valign="top">
                <table>
                    <tr>
                        <td valign="top">
                            <font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font>
                        </td>
                    </tr>
                </table>
                <td width="100%">
                    <table>
                        <tr>
                            <td>
                                <b>Report:</b>
                            </td>
                            <td width="100%">
                                <% 
	                                Select Case ShowReport
		                                Case "scm"
                                %>
                                SCM 
                                  | <a href="#" onclick="ShowReport('pm');">Program Matrix</a>
                                <%
		                                Case "pm"
                                %>
                                <a href="#" onclick="ShowReport('scm');">SCM</a> | Program Matrix
                                <%
		                                Case Else
                                %>
                                SCM | <a href="#" onclick="ShowReport('pm');">Program Matrix</a>
                                <%
	                                End Select
                                %>
                            </td>
                        </tr>
                        <% If ShowReport = "scm" Or ShowReport = "pm" Or ShowReport = "" Then %>
                        <tr>
                            <td>
                                <b>Brand:</b>
                            </td>
                            <td width="100%">
                                <%	
   'handle the situation that the passed in brand is deleted           
    dim bBrandDeleetd
        bBrandDeleetd =false
    if Request("BID") <> "" then
        rs.Filter="ProductBrandID = " & Request("BID")
        if rs.recordcount=0 then
            bBrandDeleetd =true
        end if
        rs.Filter = ""
    end if                                  
                                    
	Do Until rs.EOF
			
		If ((Request("BID") = ""  or bBrandDeleetd) And m_BrandID = "") or (CLng(rs("ProductBrandID")) = CLng(Request("BID"))) or (CLng(rs("CombinedProductBrandId")) = CLng(Request("BID"))) Then   'This part is when the Brand is currently being displayed so there is no link                                                           
            m_BrandID = rs("ProductBrandID")		
            
            if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
                m_BrandName = rs("CombinedName")
                m_BrandID = rs("CombinedProductBrandId")		            
		    else
                m_BrandName = rs("Name") 'no combined name so display the name from the Brand table
            end if
            
            if rs("ShortName") = "" then	'the shortname field is used at the top of the grid if it exists. If not, it's the old way of doing things.
			    m_ShortName = rs("streetname2") & " " & rs("SeriesSummary")
			else
				m_ShortName = rs("ShortName")
			end if

			m_LastPublishDt = rs("LastPublishDt") & ""
			If rs("LastPublishDt") & "" = "" Then
			    bUnpublished = True
			ElseIf DateDiff("h", m_LastPublishDt, Now()) <= 8 Then
			    bShowPublishRollback = True
			End If

            m_BusinessID = rs("BusinessID") & ""                     
            m_BusinessSegmentID = rs("BusinessSegmentID") & ""                     

            if sBrandDisplayed <> m_BrandName then 'have to handle if it is a combined Brand name because we don't want to display the same name multiple times
				If Not bFirstWrite Then
				    Response.Write "&nbsp;|&nbsp;"
				end if
				Response.Write server.HTMLEncode(m_BrandName)
			end if	          
           
            sBrandDisplayed = m_BrandName
		Else	' below part is when the Brand is not currently being displayed so there needs to be a link on the other Brands
			if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
				if sBrandDisplayed <> rs("CombinedName") then
					If Not bFirstWrite Then
						Response.Write "&nbsp;|&nbsp;"
					end if
					Response.Write "<a href=""javascript:BrandLink_onClick(" & rs("combinedProductBrandId") & ")"">" & server.HTMLEncode(rs("CombinedName")) & "</a>"
				end if
				sBrandDisplayed = rs("CombinedName")
			else
					if sBrandDisplayed <> rs("Name") then
						If Not bFirstWrite Then
							Response.Write "&nbsp;|&nbsp;"
						end if
			            
                        Response.Write "<a href=""javascript:BrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"
					end if
					sBrandDisplayed = rs("Name")
			end if
		End If

		bFirstWrite = False
		rs.MoveNext
	Loop
	
	rs.Close
	
                                    


                       

' Get Features Status (Not Active)

Dim rsFeature
Set rsFeature = Server.CreateObject("ADODB.RecordSet")
Set cmd = dw.CreateCommAndSP(cn, "usp_GetFeaturesNotActive_SCM")
dw.CreateParameter cmd, "@p_ProductBrandId", adInteger, adParamInput, 8, m_BrandID
Set rsFeature = dw.ExecuteCommAndReturnRS(cmd)

If rsFeature.EOF Then
  featurelist = ""
 Else
   Do While Not rsFeature.EOF

    featurelist = rsFeature.Fields("FeatureName").value + ","  + featurelist   
    rsFeature.MoveNext()
   Loop
 End If
rsFeature.Close  

                                    
'
' Get KMAT
'
Set cmd = dw.CreateCommandSQL(cn, "SELECT KMAT,ShowPhWebActionItems From Product_Brand with (NOLOCK) WHERE ID=" & m_BrandID)
Set rs = dw.ExecuteCommandReturnRS(cmd)
If not rs.EOF then
    sKmat = rs("KMAT") & ""
    ShowPhWebActionItems = rs("ShowPhWebActionItems") & ""
    'Response.Write(sKmat)
End If
rs.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvsNotInKmatBom")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
Set rs = dw.ExecuteCommAndReturnRS(cmd)
orphanCount = rs.RecordCount
rs.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_ViewUnreleasedAvs")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
Set rs = dw.ExecuteCommAndReturnRS(cmd)
UnReleasedAVCount = rs.RecordCount
rs.Close


Set cmd = dw.CreateCommAndSP(cn, "usp_SelectHiddenAvs")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
Set rs = dw.ExecuteCommAndReturnRS(cmd)
UnpublishedAvsCount = rs.RecordCount
rs.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvWithMissingData")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
Set rs = dw.ExecuteCommAndReturnRS(cmd)
AVsWithMissingDataCount = rs.RecordCount
rs.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvsMissingDeliverableRoot")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, PVID
Set rs = dw.ExecuteCommAndReturnRS(cmd)
AVsWithMissingDelRootCount = rs.RecordCount
rs.Close
                                %>
                            </td>
                        </tr>
                        <% If ShowReport = "scm" Or ShowReport = "" Then%>
                        <tr>
                            <td>
                                <b>Status:</b>
                            </td>
                            <td width="100%" colspan="2">
                                <% 
	dim strShowObsolete
	strShowObsolete = ""
	on error resume next
	strShowObsolete = Request.Cookies("ShowObsolete")
	on error goto 0
	Select Case strShowObsolete
		Case "active"
                                %>
                                Active | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | <a href="#" onclick="ShowObsolete('hidden');">Hidden</a> | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
		Case "obsolete"
                                %>
                                <a href="#" onclick="ShowObsolete('active');">Active</a> | Obsolete | <a href="#" onclick="ShowObsolete('hidden');">Hidden</a> | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
        Case "hidden"           
                                %>
                                <a href="#" onclick="ShowObsolete('active');">Active</a> | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | Hidden | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
        Case "all"              'all is no longer a default when no cookie is set yet
                                %>
                                <a href="#" onclick="ShowObsolete('active');">Active</a> | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | <a href="#" onclick="ShowObsolete('hidden');">Hidden</a> | All
                                <%
		Case Else               'active is no longer a default when no cookie is set yet
                                %>
                                Active | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | <a href="#" onclick="ShowObsolete('hidden');">Hidden</a> | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
	End Select
                                %>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <b>Filter:</b>
                            </td>
                            <td width="100%" colspan="2">
                                <% 
	                                sSCMCategories = Request("SCMCategories")
                                    intNoLocalization = Request("NoLocalization")
                                    sGADateTo = Request("GADateTo")
                                    sGADateFrom = Request("GADateFrom")
                                    sSADateTo = Request("SADateTo")
                                    sSADateFrom = Request("SADateFrom")
                                    sEMDateTo = Request("EMDateTo")
                                    sEMDateFrom = Request("EMDateFrom")
                                    sReleaseIDs = Request("ReleaseIDs")
                                    if sSCMCategories<>"" or (intNoLocalization<>"0" and intNoLocalization<>"") or sGADateTo<>"" or sGADateFrom<>"" or sSADateTo<>"" or sSADateFrom<>"" or sEMDateTo<>"" or sEMDateFrom<>"" or sReleaseIDs<>"" then
                                        sFilterApplied ="1"
                                    end if 
                                %>
                                <a href="#" onclick="FilterSCMByCategory(<%=m_BusinessID%>,<%=m_BrandID%>,<%=m_CurrentUserId%>,<%=PVID%>);">Set Filter</a>
                                <%if sFilterApplied ="1" then %>
                                <label id="lblFilterapplied">AVs displayed are filtered</label>
                                <%end if %>
                            </td>
                        </tr>
                        <% Else %>
                        <tr>
                            <td>
                                <b>Filter:</b>
                            </td>
                            <td width="100%" colspan="2">
                                <% 
	                            Select Case Request("PMCategories")
		                            Case ""
                                %>
                                <a href="#" onclick="FilterPMByCategory(<%=m_BusinessID%>,<%=m_BrandID%>,<%=m_CurrentUserId%>,<%=PVID%>);">By Category</a>
                                <%
		                            Case Else
		                              sPMCategories = Request("PMCategories")
                                %>
                                <a href="#" style="color: Red;" onclick="FilterPMByCategory(<%=m_BusinessID%>,<%=m_BrandID%>,<%=m_CurrentUserId%>,<%=PVID%>);">By Category</a>
                                <%
	                            End Select
                                %>
                            </td>
                        </tr>
                        <% End If
                    End If %>
                    </table>
                </td>
        </tr>
    </table>
    <%
Else
    Response.Write "<br /><font color=red size=3>No brand information available.  Unable to initialize the SCM module.</font><br />"
    response.End
End If
    %>
    <% 

'Check for Multiple Series Number

Set cmd = dw.CreateCommandSQL(cn, "SELECT SeriesSummary From Product_Brand with (NOLOCK) WHERE ID=" & m_BrandID)
'Response.Write(m_BrandID)
Set rs = dw.ExecuteCommandReturnRS(cmd)
If not rs.EOF then
    strSeriesSummary = Trim(rs("SeriesSummary") & "")
End If
rs.Close

bMultipleSeriesNumbers = False
If instr(strSeriesSummary, ",") > 0 Then
    bMultipleSeriesNumbers = True
End If

    %>
    <%
'##########################################################
'#
'# Draw Menu Links
'#
'##########################################################
    %>
    <br>
    <span style="font-size: xx-small; font-family: Verdana">
        <% If ShowReport = "scm" Or ShowReport = "pm" Then %>
        <span style="color: Black; font-weight: bold; cursor: pointer;">Actions:&nbsp;</span>
        <%If bIsPc Then%>
        <%If ShowPhWebActionItems = "True" Then%>
        <a href="#" onclick="EditKMAT(<%=PVID%>, <%=m_BrandID%>);">Program Data</a> |
            <% Else %>
        <a href="#" onclick="EditKMAT(<%=PVID%>, <%=m_BrandID%>);">Program Data</a> <span style="color: Red; font-weight: bold; cursor: pointer;">(Action Items Off)</span> |
            <%End If%>
        <%End If%>
        <%If bIsPc or m_IsMarketingUser Then%>
        <a href="#" onclick="SCMWorkSheetShowHide(<%=m_BrandID%>,'<%=m_IsDesktop%>','<%=bIsPc%>','<%=m_IsMarketingUser%>')">SCM Wrksht</a> |
            <a href="#" onclick="UploadScm(<%=m_BrandID%>,'<%=m_IsDesktop%>','<%=sProdVersionBSAMFlag%>', '<%=CurrentUserName%>','<%=bIsPc%>','<%=m_IsMarketingUser%>');">Upload Wrksht</a>
        |     
        <%End If%>
        <a href="/ipulsar/Reports/SCM/SCM_ExportSCM_DT.aspx?PBID=<%=m_BrandID%>&PVID=<%=PVID%>">Export SCM</a>
        |
        <%If bIsPc or (m_IsSCMPublishCoordinator and m_BusinessID=2) Then%>
        <a href="#" onclick="PublishScmNew(<%=PVID%>, <%=m_BrandID%>);">Publish SCM</a> |
        <%End If%>
        <a href="#" onclick="ExportProjectMatrix(<%=m_BrandID%>)">Program Matrix</a> |    
       
        <%If bUnpublished Then %>
        <%if (blnCanClearChangeLog) then%>
        <a href="javascript:clearAvHistory(<%=m_BrandID%>)">Clear Change Log</a> |
        <%end if%>
        <%End If 'bUnpublished %>
        <a href="<%=AppRoot %>/SupplyChain/ChangeLog.aspx?ID=<%=PVID%>&BID=<%=m_BrandID%>" target="New">Show Change Log</a>
        | <a href="#" onclick="ShowABTInfoDialog(<%=m_BrandID%>, <%=EditABTInformation%>);">ABT Information</a>
        <%If m_IsProgramCoordinator or m_IsConfigurationManager Then%>
            | <a href="#" onclick="ShowSCMManufacturingSitesDialog(<%=m_BrandID%>, <%=m_BusinessSegmentID%>);">Assign Manufacturing Sites</a>
        <%End IF %>
         | <a href="#" onclick="ShowChinaGPDialog(<%=m_BrandID%>);">Edit China GP Identifiers</a>
        <%  If bShowPublishRollback Then %>
            | <a href="javascript:ScmPublishRollback(<%=m_BrandID%>,'<%=CurrentUserName%>','<%=m_IsProgramCoordinator%>')">Rollback SCM Publish</a>
        <%  End If 'bShowPublishRollback %>
        <br />
        <br />
        <span style="color: Black; font-weight: bold; cursor: pointer;">Tools:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
        <a href="#" onclick="QuickSearch()">QuickSearch</a>
        <%If bIsPc Then%>          
            | <a href="#" onclick="ShowAddMultipleAVsDialog('<%=m_BrandID%>','<%=m_IsDesktop%>','<%=m_CurrentUserId%>','<%=PVID%>');">Add AV(s)</a>
        <!--Renamed link for combined AV functionality-->
        | <a href="#" onclick="ShowAddExistingSharedAV(<%=PVID%>, <%=m_BrandID%>);">Add Existing Shared AV</a>
        | <a href="#" onclick="LocalizeAVs(<%=PVID%>, '<%=strSeriesSummary%>', '<%=bMultipleSeriesNumbers%>', <%=m_BrandID%>, '<%=CurrentUserName%>', '<%=sFusionRequirements%>');">Add Localized AVs</a>

        | <a href="#" onclick="ExportRPN(<%=m_BrandID%>);">Export RPN</a>
        | <a href="#" onclick="ExportRPNToS4(<%=m_BrandID%>);">Export RPN to S4</a>
        | <a href="#" onclick="UploadAvData(<%=m_BrandID%>,<%=PVID%>,'<%=CurrentUserName%>')">Import AV Data</a>
        | <a href="#" onclick="ImportPN(<%=PVID%>)">Import PN from S4</a>
        | <a href="#" onclick="AddSAs()">Assign SAs</a>
        | <a href="#" onclick="StructureBOM('<%=CurrentUserName%>', <%=m_BrandID%>, '<%=sKmat%>', <%=m_BusinessID%>, <%=PVID%>)">Structure BOM</a>
        | <a href="#" onclick="ShowSCMBUInformation(<%=PVID%>,<%=m_BrandID%>,'<%=bIsPc%>')">SCM Base Unit Information</a>
        <% End If %>
        <br />
        <br />
        <span style="color: Black; font-weight: bold; cursor: pointer;">Reports:</span>
        <%If bIsPc Then%>
        <% If orphanCount > 0 Then %>
        <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
            onclick="showOrphanedAvReport(<%=m_BrandID%>)">Orphaned AVs</span> |
            <% Else %>
        <a href="#" onclick="showOrphanedAvReport(<%=m_BrandID%>)">Orphaned AVs</a> |
            <% End if 'orphanCount %>
        <% If UnReleasedAVCount > 0 Then %>
        <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
            onclick="showUnreleasedAvReport(<%=m_BrandID%>)">Unreleased AVs</span> |
            <% Else %>
        <a href="#" onclick="showUnreleasedAvReport(<%=m_BrandID%>)">Unreleased AVs</a> | 
            <% End if 'UnReleasedAVCount %>


        <%End if%>

        <%  If bIsPc or bIsPDM Then %>
        <%If ShowPhWebActionItems = "True" And PendingAvActions = "True" Then%>
        <a href="/iPulsar/ExcelExport/PendingPhWebActions.aspx?PVID=<%=PVID%>">Pending
			AV Actions</a> <span style="color: Red; font-weight: bold; cursor: pointer;"></span>|
        <% End If 'ShowPhWebActionItems
         End If %>
        <a href="#" onclick="AvActionScorecard(<%=PVID%>, <%=m_CurrentUserId%>);">AV Action Scorecard</a>
        <span style="color: Red; font-weight: bold; cursor: pointer;"></span>
        <%If bIsPc Then%>
        | <a href="#" onclick="DocKitData('<%=sKmat%>');">Check Doc Kit Status</a>
        <%If UnpublishedAvsCount > 0 Then %>
        | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
            onclick="showUnpublishedAVs(<%=m_BrandID%>)">Hidden AVs</span>
        <% Else %>
        | <a href="#" onclick="showUnpublishedAVs(<%=m_BrandID%>)">Hidden AVs</a>
        <% End if 'Unpublished AVs %>
        <%End If
        If bIsPc or bIsPDM Then %>
        <% If AVsWithMissingDataCount > 0 Then %>
            | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
                onclick="showAVsWithMissingData(<%=m_BrandID%>,'<%=bIsPc%>')">AVs Missing Corporate Data</span>
        <% Else %>
            | <a href="#" onclick="showAVsWithMissingData(<%=m_BrandID%>,'<%=bIsPc%>')">AVs Missing Corporate Data</a>
        <% End if 'AVs Missing Corporate Data %>
        <span style="display: none"><a href="#" onclick="CopyScm(<%=PVID%>, <%=m_BrandID%>);">Copy Existing SCM</a> |</span>
        <% If AVsWithMissingDelRootCount > 0 Then %>
            | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
                onclick="showAVsWithMissingDelRoot(<%=m_BrandID%>,<%=PVID%>)">AVs Missing Deliverable Root</span>
        <% Else %>
            | <a href="#" onclick="showAVsWithMissingDelRoot(<%=m_BrandID%>,<%=PVID%>)">AVs Missing Deliverable Root</a>
        <% End if 'AVs Missing Deliverable Root %>
        <%End If%>
        | <a href="#" onclick="EOSReport(<%=m_BrandID%>)">End of Sales Report</a> <span style="color: Red; font-weight: bold; cursor: pointer;"></span>
        | <a href="#" onclick="OpenSCMReportList(<%=PVID%>,<%=m_BrandID%>)">SCM/Program Matrix Reports</a> <span style="color: Red; font-weight: bold; cursor: pointer;"></span>

        <%
    End If 'ShowReport = "scm" Or ShowReport = "pm"
        %>
    </span>
    <p>
        <span style="font-size: x-small; font-weight: bold;">
            <!--Use Combined brand name if it is a combinedbrand else use Name from Brand table - task 16893 -->
            <%= m_BrandName%>
            <%
            Select Case ShowReport
                Case "scm"
                    Response.Write "&nbsp;-&nbsp;Supported Configuration Matrix <span id=""lblAvCount"" style=""font-size:xx-small; font-weight: normal;""></span><span style=""font-size:xsmall; color: Red; font-weight: bold;"">&nbsp; - Global GBU View</span>"
                Case "pm"
                    Response.Write "&nbsp;-&nbsp;Program Matrix"
                Case Else
                    Response.Write "&nbsp;-&nbsp;Supported Configuration Matrix<span style=""font-size:xsmall; color: Red; font-weight: bold;""> - Global GBU View<span>"
            End Select
            %></span>
        <div id="scmLoading">
            <span style="font: Bold x-small Tahoma; color: Red;">Loading
                Please Wait...</span>
        </div>
        <%

Response.Flush

If ShowReport = "" Or ShowReport = "scm" Then
'##########################################################
'#
'# Draw SCM
'#
'##########################################################
Dim strCategoryIDName : strCategoryIDName = ""
Dim strCategoryName : strCategoryName = ""

    Set cmd = dw.CreateCommAndSP(cn, "usp_SelectScmDetail_Pulsar")
    strCategoryIDName = "SCMCategoryID"
    strCategoryName = "SCMCategoryName"

dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, PVID
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID

strShowObsolete = ""
on error resume next
strShowObsolete = Request.Cookies("ShowObsolete")
on error goto 0
'response.Write("<div>ShowObsolete Status (A or O or H): " & strShowObsolete & "</div>") 
Select Case strShowObsolete
	Case "active"
		dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
	Case "obsolete"
		dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "O"
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
    Case "hidden"
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "H"
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
    Case "all" 'all is no longer a default when no cookie is set yet
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, ""
        dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
	Case Else 'active is no longer a default when no cookie is set yet
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
        dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, "H"
End Select

If request("SCMCategories") <> "" Then
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, Request("SCMCategories")
Else
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, ""
End If  

If request("NoLocalization") <> "" Then
    dw.CreateParameter cmd, "@p_NoLocalization", adVarchar, adParamInput, 500, Request("NoLocalization")
Else
    dw.CreateParameter cmd, "@p_NoLocalization", adVarchar, adParamInput, 500, ""
End If

If request("GADateTo") <> "" Then
    dw.CreateParameter cmd, "@p_GADateTo", adVarchar, adParamInput, 500, Request("GADateTo")
Else
    dw.CreateParameter cmd, "@p_GADateTo", adVarchar, adParamInput, 500, ""
End If

If request("GADateFrom") <> "" Then
    dw.CreateParameter cmd, "@p_GADateFrom", adVarchar, adParamInput, 500, Request("GADateFrom")
Else
    dw.CreateParameter cmd, "@p_GADateFrom", adVarchar, adParamInput, 500, ""
End If

If request("SADateTo") <> "" Then
    dw.CreateParameter cmd, "@p_SADateTo", adVarchar, adParamInput, 500, Request("SADateTo")
Else
    dw.CreateParameter cmd, "@p_SADateTo", adVarchar, adParamInput, 500, ""
End If

If request("SADateFrom") <> "" Then
    dw.CreateParameter cmd, "@p_SADateFrom", adVarchar, adParamInput, 500, Request("SADateFrom")
Else
    dw.CreateParameter cmd, "@p_SADateFrom", adVarchar, adParamInput, 500, ""
End If

If request("EMDateTo") <> "" Then
    dw.CreateParameter cmd, "@p_EMDateTo", adVarchar, adParamInput, 500, Request("EMDateTo")
Else
    dw.CreateParameter cmd, "@p_EMDateTo", adVarchar, adParamInput, 500, ""
End If

If request("EMDateFrom") <> "" Then
    dw.CreateParameter cmd, "@p_EMDateFrom", adVarchar, adParamInput, 500, Request("EMDateFrom")
Else
    dw.CreateParameter cmd, "@p_EMDateFrom", adVarchar, adParamInput, 500, ""
End If

If request("ReleaseIDs") <> "" Then
    dw.CreateParameter cmd, "@p_ReleaseIDs", adVarchar, adParamInput, 500, Request("ReleaseIDs")
Else
    dw.CreateParameter cmd, "@p_ReleaseIDs", adVarchar, adParamInput, 500, ""
End If

Set rs = dw.ExecuteCommAndReturnRS(cmd)
iAvCount = 0
If rs.EOF Then
        %>
        <table id="TableScm" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <font face="Verdana" size="2">No Av Items found for this brand.</font>
                </td>
            </tr>
        </table>
        <%
Else
        %>
        <div id="GridViewContainer" class="GridViewContainer" style="width: 100%; height: 100%;">
            <table id="Table1" class="Table" width="100%">
                <col width="90">
                <col width="90">
                <col width="200">
                <col width="145">
                <col width="145">
                <col width="145">
                <col width="70" align="center" />
                <col width="70" align="center" />
                <col width="70" align="center" />
                <col width="70" align="center" />
                <col width="120">
                <col width="70">
                <col width="70">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50" align="center">
                <col width="50" align="center">
                <col width="50" align="center">
                <col width="40" align="center">
                <col width="40" align="center">
                <col width="50" align="center">
                <col width="50" align="center">
                <col width="50" align="center" />

                <thead>
                    <tr class="FrozenHeader">
                        <th>AV No.</th>
                        <th>Feature ID</th>
                        <th>GPG Description</th>
                        <th>Marketing Short Description<br />
                            (40 Char)</th>
                        <th>Marketing Long Description<br />
                            (100 Char)</th>
                        <th>Release(s)</th>
                        <th>PA:AD<br />
                            (Intro Date)</th>
                        <th>Select Availability<br />
                            (SA) Date </th>
                        <th>General Availability<br />
                            (GA) Date </th>
                        <th>End of Manufacturing<br />
                            (EM) Date</th>
                        <th>Global Series Config
                            <br />
                            Planned End</th>
                        <th>Configuration Rules</th>
                        <th>Rule Syntax</th>
                        <th>AVID</th>
                        <th>Group 1</th>
                        <th>Group 2</th>
                        <th>Group 3</th>
                        <th>Group 4</th>
                        <th>Group 5</th>
                        <th>Group 6</th>
                        <th>Group 7</th>
                        <th>IDS</th>
                        <th>IDS CTO</th>
                        <th>RCTO</th>
                        <th>RCTO CTO</th>
                        <% If (sProdVersionBSAMFlag = "True") Then %>
                        <th>BSAM -B</th>
                        <% End If %>
                        <th>UPC</th>
                        <th>Weight<br />
                            (in oz)</th>
                        <th>Product<br />
                            Line</th>
                        <th style="display: none">Feature Category
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <%
	Do Until rs.EOF
	   If (iSCMCategoryID <> rs(strCategoryIDName)) Then
			iSCMCategoryID = rs(strCategoryIDName)
			sSCMCategory = rs(strCategoryName)
                    %>
                    <tr class="FeatureCategory" pvid="<%=PVID%>" fcid="<%=iSCMCategoryID%>" bid="<%=m_BrandID%>" onmouseover="return FRMO()" onmouseout="return FRMOut()" onclick="return FROC()">
                        <td colspan="3">
                            <span style="font-weight:bold;">
                                <%=sSCMCategory%></span>
                            <%If bIsPc Then%>
                            <br />
                            <input type="checkbox" id="chkAll<%=rs(strCategoryIDName)%>" name="chkAll<%=rs(strCategoryIDName)%>"
                                style="width: 16px; height: 16px" onclick="chkAll_onclick()">
                            <%End If%>
                            <input type="button" value="View Availability" id="btnAvailMultipleAvs<%=rs(strCategoryIDName)%>" name="btnAvailMultipleAvs"
                                class="button2" style="width: 100px" onclick="ViewAvailMultipleAvs(<%=PVID%>,<%=m_BrandID%>,'<%=bIsPc%>')">
                        </td>
                        <td id="tdCategoryMD<%=rs(strCategoryIDName)%>">
                            <%= PrepForWeb(rs("CategoryMarketingDescription"))%>
                        </td>
                        <td colspan="7"></td>
                        <td id="tdCategoryRules<%=rs(strCategoryIDName)%>">
                            <%= PrepForWeb(rs("CategoryRules"))%>
                        </td>
                        <td id="tdCategoryRuleSyntax<%=rs(strCategoryIDName)%>">
                            <%= PrepForWeb(rs("CategoryRuleSyntax"))%>
                        </td>
                        <td><%= PrepForWeb(rs("CategoryAvId"))%></td>
                        <td id="tdCatMin<%=rs(strCategoryIDName)%>">
                            <% sDisplayMinMax = PrepForWeb(rs("CatMin"))
							if sDisplayMinMax <> "&nbsp;" then 
								response.write "MIN=" & sDisplayMinMax
							else
								response.write sDisplayMinMax
							end if
                            %>
                        </td>
                        <td id="tdCatMax<%=rs(strCategoryIDName)%>">
                            <% sDisplayMinMax = PrepForWeb(rs("CatMax"))
							if sDisplayMinMax <> "&nbsp;" then 
								response.write "MAX=" & sDisplayMinMax
							else
								response.write sDisplayMinMax
							end if
                            %>
                        </td>
                        <td colspan="21"></td>
                    </tr>
                    <%
            Response.Flush
		End If
		If rs.Fields(0).Value & "" <> "" Then
                    %>
                    <tr class='<%=rs("Status")%>' id="AV<%=rs.Fields(0).Value%>" pvid="<%=PVID%>" avid="<%=rs.Fields(0).Value%>" mktdesc="<%=rs("MarketingDescription")%>" mktdescpmg="<%=rs("MarketingDescriptionPMG")%>"
                        status="<%=rs("Status")%>" bid="<%=rs("ProductBrandID")%>" parentid="<%=rs("ParentID")%>" bscmpublished="<%=rs("bSCMPublished")%>" onmousedown="return AVMD(<%=PVID%>,<%=rs.Fields(0).Value%>)" onmouseover="return AVMO()" onmouseout="return AVMOut()" onclick="return AVOC()">
                        <td nowrap>
                            <%If bIsPc Then%>
                            <div style="float: left; width: 18%">
                                <input class='<%=rs(strCategoryIDName)%>' type='checkbox' id='chkAv<%=rs(strCategoryIDName)%>'
                                    name='chkAv<%=rs("AvDetailID")%>' value='<%=rs(strCategoryIDName)%>' style="width: 16px; height: 16px"
                                    onclick="return chkAv_onclick()">
                            </div>
                            <%End If%>
                            <div id="divAVNo<%=rs.Fields(0).Value%>" style="float: left; width: 80%"><%=PrepForWeb(rs("AvNo"))%></div>
                        </td>
                        <td><%=PrepForWeb(rs("FeatureID"))%></td>
                        <%if rs("GPGDescSysUpdate") = 0 then%>
                        <td id="Gdsc<%=rs.Fields(0).Value%>" style="color: Gray" nowrap><%=PrepForWeb(rs("GPGDescription"))%></td>
                        <%else%>
                        <td id="Gdsc1<%=rs.Fields(0).Value%>" nowrap><%=PrepForWeb(rs("GPGDescription"))%></td>
                        <%end if %>
                        <%if rs("MktDescSysUpdate") = 0 then%>
                        <td id="Mdsc<%=rs.Fields(0).Value%>" style="color: Gray" nowrap><%=PrepForWeb(rs("MarketingDescription"))%></td>
                        <%else%>
                        <td id="Mdsc1<%=rs.Fields(0).Value%>" nowrap><%=PrepForWeb(rs("MarketingDescription"))%></td>
                        <%end if %>
                        <%if rs("MktDescPMGUSysUpdate") = 0 then%>
                        <td id="MdscPMG<%=rs.Fields(0).Value%>" style="color: Gray" nowrap><%=PrepForWeb(rs("MarketingDescriptionPMG"))%></td>
                        <%else%>
                        <td id="MdscPMG1<%=rs.Fields(0).Value%>" nowrap><%=PrepForWeb(rs("MarketingDescriptionPMG"))%></td>
                        <%end if %>
                        <td id="rel<%=rs.Fields(0).Value%>" nowrap><%=PrepForWeb(rs("Releases"))%></td>

                        <%if rs("PAADDate") = 0 then%>
                        <td style="color: Gray"></td>
                        <%else%>
                        <td id="tdPAADDate<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("PAADDate"))%></td>
                        <%end if %>

                        <%if rs("CplBlindSysUpdate") = 0 then%>
                        <td id="CPLdt<%=rs.Fields(0).Value%>" style="color: Gray"><%=PrepForWeb(rs("CPLBlindDt"))%></td>
                        <%else%>
                        <td id="CPLdt<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("CPLBlindDt"))%></td>
                        <%end if %>
                        <%if rs("GeneralAvailSysUpdate") = 0 then%>
                        <td id="tdGADate<%=rs.Fields(0).Value%>" style="color: Gray"><%=PrepForWeb(rs("GeneralAvailDt"))%></td>
                        <%else%>
                        <td id="tdGADate<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("GeneralAvailDt"))%></td>
                        <%end if %>
                        <%if rs("RasDiscoSysUpdate") = 0 then%>
                        <td id="RASdt<%=rs.Fields(0).Value%>" style="color: Gray"><%=PrepForWeb(rs("RASDiscontinueDt"))%></td>
                        <%else%>
                        <td id="RASdt<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("RASDiscontinueDt"))%></td>
                        <%end if %>
                        <td id="GSDt<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("GSEndDt"))%></td>
                        <td id="rules<%=rs.Fields(0).Value%>" nowrap><%=PrepForWeb(rs("ConfigRules"))%></td>
                        <td id="rulesX<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("RuleSyntax"))%></td>
                        <td id="tdAvId<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("AvId"))%></td>
                        <td id="g1<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group1"))%></td>
                        <td id="g2<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group2"))%></td>
                        <td id="g3<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group3"))%></td>
                        <td id="g4<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group4"))%></td>
                        <td id="g5<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group5"))%></td>
                        <td id="g6<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group6"))%></td>
                        <td id="g7<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Group7"))%></td>
                        <td id="ids<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("IdsSkus_YN"))%></td>
                        <td id="idsc<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("IdsCto_YN"))%></td>
                        <td id="rcto<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("RctoSkus_YN"))%></td>
                        <td id="rctoc<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("RctoCto_YN"))%></td>
                        <% if (sProdVersionBSAMFlag = "True") Then %>
                        <td id="tdBSAMB<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("BSAMBparts_YN"))%></td>
                        <% End If %>
                        <td><%=PrepForWeb(rs("UPC"))%></td>
                        <td id="wt<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("Weight"))%></td>
                        <td id="PL<%=rs.Fields(0).Value%>"><%=PrepForWeb(rs("ProductLineName"))%></td>
                        <td id="SCMCat<%=rs.Fields(0).Value%>" style="display: none"><%=PrepForWeb(rs(strCategoryIDName))%></td>
                    </tr>
                    <%
                iAvCount = iAvCount + 1        
		End If
        Response.Flush
		rs.MoveNext
	Loop
                    %>
                </tbody>
            </table>
        </div>
    </p>
    <%
End If
    %>
    <% ElseIf ShowReport = "pm" Then %>
    <%
'##########################################################
'#
'# Draw Program Matrix
'#
'##########################################################
Set cmd = dw.CreateCommAndSP(cn, "usp_SelectProgramMatrixSS_Pulsar")
dw.CreateParameter cmd, "@p_KMAT", adVarchar, adParamInput, 10, sKmat
If request("PMCategories") <> "" Then
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, Request("PMCategories")
Else
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, ""
End If  
Set rs = dw.ExecuteCommAndReturnRS(cmd)

Dim iVersions
Dim rsVer
Set rsVer = Server.CreateObject("ADODB.RecordSet")
Set cmd = dw.CreateCommAndSP(cn, "usp_SelectKmatVersions")
dw.CreateParameter cmd, "@p_ProductBrandId", adInteger, adParamInput, 8, m_BrandID
Set rsVer = dw.ExecuteCommAndReturnRS(cmd)
iVersions = 0
Do Until rsVer.EOF
    iVersions = iVersions + 1
    rsVer.MoveNext
Loop

iVersions = iVersions - 1

rsVer.Close

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvVersions")
dw.CreateParameter cmd, "@p_ProductBrandId", adInteger, adParamInput, 8, m_BrandID
dw.CreateParameter cmd, "@p_CompareDt", adDate, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_Publish", adBoolean, adParamInput, 1, 0
Set rsVer = dw.ExecuteCommAndReturnRS(cmd)

Dim sLastAv
Dim sLastSa
Dim sLastComponent
Dim sLastComponent_L5
Dim sLastComponent_L6
Dim sLastCategory

If Not rs.EOF Then
    %>
    <p style="font-size: xx-small; font-family: Verdana; color: Red; font-weight: bold;">
        This is not the Published Matrix. Working Data Snapshot taken at
        <%= rs("ExportTime") %>
    </p>
    <% End If %>
    <div id="GridViewContainer" class="GridViewContainer" style="width: 100%; height: 100%;">
        <table class="MatrixTable">
            <thead>
                <tr class="FrozenHeader">
                    <th>Category / Manufacturing Comments</th>
                    <th colspan="<%= iVersions + 1 %>">Product Version</th>
                    <th>Description</th>
                    <th>AV<br />
                        Level 2</th>
                    <th>SA<br />
                        Level 3</th>
                    <th>Component<br />
                        Level 4</th>
                    <th>Component<br />
                        Level 5</th>
                    <th>Component<br />
                        Level 6</th>
                    <th>Qty</th>
                    <th>SAP Rev</th>
                    <th>ZWAR = X<br />
                        PRI/ALT</th>
                    <th>ROHS</th>
                    <th>UPC</th>
                    <th>IDS</th>
                    <th>IDS-CTO</th>
                    <th>RCTO</th>
                    <th>RCTO-CTO</th>
                    <% If (sProdVersionBSAMFlag = "True") Then %>
                    <th>BSAM -B</th>
                    <% End If %>
                    <th>Config Rules</th>
                    <th>Rule Syntax</th>
                    <th>AVID</th>
                    <th>Group 1</th>
                    <th>Group 2</th>
                    <th>Group 3</th>
                    <th>Group 4</th>
                    <th>Group 5</th>
                    <th>Group 6</th>
                    <th>Group 7</th>
                </tr>
            </thead>
            <tbody>
                <%
Do Until rs.EOF
	If sLastCategory <> Trim(rs("FeatureCategory") & "") Then
		sLastCategory = Trim(rs("FeatureCategory") & "")
		PmDrawCategoryRow rs, iVersions
	End If

	If Trim(rs("AvNo") & "") <> "" Then
		If sLastAv <> Trim(rs("AvNo") & "") Then
			PmDrawAvRow rs, rsVer, iVersions
		ElseIf sLastSa <> Trim(rs("SaNo") & "") AND Trim(rs("SaNo") & "") <> "" Then
			PmDrawSaRow rs, iVersions
		ElseIf sLastComponent <> Trim(rs("CompNo") & "") AND Trim(rs("CompNo") & "") <> "" AND Trim(rs("SaNo") & "") <> "" Then
			PmDrawComponentRow rs, iVersions
        ElseIf sLastComponent_L5 <> Trim(rs("CompNo_L5") & "") AND Trim(rs("CompNo_L5") & "") <> "" AND Trim(rs("CompNo") & "") <> "" AND Trim(rs("SaNo") & "") <> "" Then
			PmDrawComponentRow_L5 rs, iVersions
        ElseIf sLastComponent_L6 <> Trim(rs("CompNo_L6") & "") AND Trim(rs("CompNo_L6") & "") <> "" AND Trim(rs("CompNo_L5") & "") <> "" AND Trim(rs("CompNo") & "") <> "" AND Trim(rs("SaNo") & "") <> "" Then
			PmDrawComponentRow_L6 rs, iVersions
		End If
	End If    
	rs.MoveNext
    Response.Flush()
Loop
rs.Close
rsVer.Close


Sub PmDrawCategoryRow(rs, iVersions)
    If (sProdVersionBSAMFlag = "True") Then
        Response.Write "<tr class=""MatrixFeatureCategory""><td colspan=" & 28 + iVersions & ">"
    else
        Response.Write "<tr class=""MatrixFeatureCategory""><td colspan=" & 27 + iVersions & ">"
    end if
    Response.Write rs("FeatureCategory") & "</td></tr>"
End Sub

Sub PmDrawAvRow(rs, rsVer, iVersions)
    sLastAv = Trim(rs("AvNo") & "")
    
    rsVer.Filter= "SortAvNo = '" & rs("AvNo") & "'"

    Response.Write "<tr class=""MatrixAvRow""><td>"
    Response.Write rs("AVManufacturingNotes") & "" 
    Response.Write "</td>"
    Do Until rsVer.EOF
        Response.Write "<td>"
        Response.Write rsVer("Version")
        Response.Write "</td>"
        rsVer.MoveNext
    Loop
    Response.Write "<td nowrap>"
    Response.Write rs("GPGDescription") & ""
    Response.Write "</td><td nowrap>"
    Response.Write rs("AvNo") & ""
    Response.Write "</td><td colspan=""5""></td><td>"
    Response.Write rs("RevisionLevel") & ""
    Response.Write "</td><td>"
    Response.Write rs("ZWAR") & ""
    Response.Write "</td><td>"
    Response.Write rs("ROHS") & ""
    Response.Write "</td><td>"
    Response.Write rs("UPC") & ""
    Response.Write "</td><td>"
    Response.Write rs("IDSSKUS") & ""
    Response.Write "</td><td>"
    Response.Write rs("IDSCTO") & ""
    Response.Write "</td><td>"
    Response.Write rs("RCTOSKUS") & ""
    Response.Write "</td><td>"
    Response.Write rs("RCTOCTO") & ""
    Response.Write "</td><td>"
    If (sProdVersionBSAMFlag = "True") Then
        Response.Write rs("bsambparts") & ""
        Response.Write "</td><td>"
    End If
    Response.Write rs("ConfigRules") & ""
    Response.Write "</td><td>"
    Response.Write rs("RuleSyntax") & ""
    Response.Write "</td><td>"
    Response.Write rs("AvId") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group1") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group2") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group3") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group4") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group5") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group6") & ""
    Response.Write "</td><td>"
    Response.Write rs("Group7") & ""
    Response.Write "</td></tr>"

    PmDrawSaRow rs, iVersions
End Sub

Sub PmDrawSaRow(rs, iVersions)
Dim i

If Trim(rs("SaNo") & "") <> "" Then
    sLastSa = Trim(rs("SaNo") & "")
    Response.Write "<tr><td></td>"
    For i = 0 to iVersions
        Response.Write "<td></td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;"
    Response.Write rs("saGPGDescription") & ""
    Response.Write "</td><td></td><td nowrap>"
    Response.Write rs("SaNo") & ""
    Response.Write "</td><td></td><td></td><td></td><td>"
    Response.Write rs("saQuantity") & ""
    Response.Write "</td><td>"
    Response.Write rs("saRevisionLevel") & ""
    Response.Write "</td><td>"
    Response.Write rs("saZWAR") & ""
    Response.Write "</td><td>"
    Response.Write rs("saROHS") & ""
    Response.Write "</td><td colspan=""15""></td></tr>"
    
    PmDrawComponentRow rs, iVersions
End If
End Sub

Sub PmDrawComponentRow(rs, iVersions)
Dim i

If Trim(rs("CompNo") & "") <> "" Then
    sLastComponent = Trim(rs("CompNo") & "")
    Response.Write "<tr><td></td>"
    For i = 0 to iVersions
        Response.Write "<td></td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription") & ""
    Response.Write "</td><td></td><td></td><td nowrap>"
    Response.Write rs("CompNo") & ""
    Response.Write "</td><td></td><td></td><td>"
    Response.Write rs("CompQuantity") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS") & ""
    Response.Write "</td><td colspan=""15""></td></tr>"
    PmDrawComponentRow_L5 rs, iVersions
End If
End Sub

Sub PmDrawComponentRow_L5(rs, iVersions)
Dim i

If Trim(rs("CompNo_L5") & "") <> "" Then
    sLastComponent_L5 = Trim(rs("CompNo_L5") & "")
    Response.Write "<tr><td></td>"
    For i = 0 to iVersions
        Response.Write "<td></td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription_L5") & ""
    Response.Write "</td><td></td><td></td><td></td><td nowrap>"
    Response.Write rs("CompNo_L5") & ""
    Response.Write "</td><td></td><td>"
    Response.Write rs("CompQuantity_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS_L5") & ""
    Response.Write "</td><td colspan=""15""></td></tr>"
    PmDrawComponentRow_L6 rs, iVersions
End If
End Sub

Sub PmDrawComponentRow_L6(rs, iVersions)
Dim i

If Trim(rs("CompNo_L6") & "") <> "" Then
    sLastComponent_L6 = Trim(rs("CompNo_L6") & "")
    Response.Write "<tr><td></td>"
    For i = 0 to iVersions
        Response.Write "<td></td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription_L6") & ""
    Response.Write "</td><td></td><td></td><td></td><td></td><td nowrap>"
    Response.Write rs("CompNo_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompQuantity_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS_L6") & ""
    Response.Write "</td><td colspan=""15""></td></tr>"
End If
End Sub
                %>
            </tbody>
        </table>
    </div>
    <%End If%>
    <!--Popups-->
    <div id="PopUpMenu" class="hidden">
        <ul id="menu">
            <li class="default"><a href="#" onclick="parent.location.href='javascript:MenuProperties();'">Properties</a></li>
            <li id="spacer">
                <hr width="95%">
            </li>
            <li id="obsolete"><a href="#" onclick="parent.location.href='javascript:MenuSetObsolete(\'<%=bIsPc%>\');'">Obsolete This AV</a></li>
            <li id="activate"><a href="#" onclick="parent.location.href='javascript:MenuSetActive(\'<%=bIsPc%>\');'">Activate This AV</a></li>
            <li id="Li2">
                <hr width="95%" />
            </li>
            <li id="hide"><a href="#" onclick="parent.location.href='javascript:MenuSetHidden(\'<%=bIsPc%>\');'">Hide This AV (SCM/PM)</a> </li>
            <li id="unhide"><a href="#" onclick="parent.location.href='javascript:MenuSetActive(\'<%=bIsPc%>\' );'">Unhide This AV (SCM/PM)</a> </li>
            <%'if m_IsDesktop then %>
            <li id="Li3">
                <hr width="95%" />
            </li>
            <li id="Avail"><a href="#" onclick="parent.location.href='javascript:ViewAvailabilityAv(\'<%=bIsPc%>\');'">View Availability</a></li>
            <%'end if %>
        </ul>
    </div>

    <div id="modal_programmatrix" class="content-dialog hide">
        <div class="menu_form_body">
            <form method="post" action="/IPulsar/Reports/SCM/ProgramMatrix.aspx?PBID=<%=m_BrandID %>" id="popup_form">
                <input type="hidden" name="ProductBrandID" value="<%=m_BrandID %>" />
                <input type="hidden" name="ServiceFamilyPn" value="<%=sServiceFamilyPn %>" />
                <table>
                    <%If sKmat = "" Then %>
                    <tr>
                        <td width="100%" colspan="2">
                            <span style="color: Red; text-align: center; font-weight: bold;">KMAT is not saved in
                            Program Data.<br />
                                BOM Information will not be available.</span>
                        </td>
                    </tr>
                    <%End If 'sKmat = "" %>
                    <tr id="trCompare">
                        <th style="white-space: nowrap; text-align: left">Comparison:</th>
                        <td>
                            <div id="selCompareDtDiv">
                                <select class="form" name="selCompareDt" id="selCompareDt">
                                    <%'---Get Program Matrix Dates: --- 
                                    Call GetProgramMatrixDates(True, m_BrandID)			
	                                    If Not (oRSProgramMatrixDates Is Nothing) Then
		                                    If Not oRSProgramMatrixDates.EOF Then 
			                                    Do While Not oRSProgramMatrixDates.EOF%>
                                    <option value="<%=oRSProgramMatrixDates("ExportTime")%>">
                                        <%=DayOfWeek(oRSProgramMatrixDates("ExportTime")) %>&nbsp;<%=oRSProgramMatrixDates("ExportTime")%>
                                    </option>
                                    <% oRSProgramMatrixDates.Movenext								
					                            Loop
		                                    End If
	                                    End If
                                    Call GetProgramMatrixDates(False, Empty)  %>
                                </select>
                            </div>
                        </td>
                    </tr>
                    <tr id="trManufacturingSites">
                        <th style="white-space: nowrap; text-align: left">Manufacturing Site:</th>
                        <td>
                            <div id="selManufacturingSitesDiv">
                            </div>
                        </td>
                    </tr>
                    <% If (bIsPc And ShowReport <> "spb") Then %>
                    <tr>
                        <th style="white-space: nowrap; text-align: left">Publish:</th>
                        <td>
                            <input type="checkbox" name="chkPublish" title="Publish" onclick="chkPublish_Click(<%= m_BrandID %>, this.checked);" />
                        </td>

                    </tr>
                    <tr>
                        <th style="text-align: left">Push to x-ROST:</th>
                        <td>
                            <input type="checkbox" name="chkXrost" id="chkXrost" title="XRost" onclick="pushToXrost(<%= m_BrandID %>,this.checked);" />
                        </td>
                    </tr>
                    <%end if%>
                    <tr>
                        <th style="white-space: nowrap; text-align: left">New&nbsp;Matrix:</th>
                        <td>
                            <input type="checkbox" id="chkNewMatrix" name="chkNewMatrix" title="Publish" />
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>
                            <input class="btn" type="submit" id="popup_submit" value="Export" onclick="closeModalDialog(false);" />
                        </td>
                    </tr>
                </table>
            </form>
        </div>
    </div>
    <input type="hidden" id="txtClass" name="txtClass" value="<%=sClass%>" />
    <input type="hidden" id="txtID" name="txtID" value="<%=PVID%>" />
    <input type="hidden" id="txtFavs" name="txtFavs" value="<%=sFavs%>" />
    <input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=sFavCount%>" />
    <input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>" />
    <input type="hidden" id="hidLastPublishDt" value="<%=m_LastPublishDt %>" />
    <input type="hidden" id="hidBusinessID" value="<%=m_BusinessID %>" />
    <input type="hidden" id="hidBusinessSegmentID" value="<%=m_BusinessSegmentID %>" />
    <input type="hidden" id="hidAvCount" value="<%= iAvCount %>" />
    <input type="hidden" id="hidSCMCategories" value="<%= sSCMCategories %>" />
    <input type="hidden" id="hidGADateTo" value="<%= sGADateTo %>" />
    <input type="hidden" id="hidGADateFrom" value="<%= sGADateFrom %>" />
    <input type="hidden" id="hidSADateTo" value="<%= sSADateTo %>" />
    <input type="hidden" id="hidSADateFrom" value="<%= sSADateFrom %>" />
    <input type="hidden" id="hidEMDateTo" value="<%= sEMDateTo %>" />
    <input type="hidden" id="hidEMDateFrom" value="<%= sEMDateFrom %>" />
    <input type="hidden" id="hidReleaseIDs" value="<%= sReleaseIDs %>" />
    <input type="hidden" id="hidSCMNoLocalization" value="<%= intNoLocalization %>" />
    <input type="hidden" id="hidPMCategories" value="<%= sPMCategories %>" />
    <input type="hidden" id="hidProductBrandID" value="<%= m_BrandID %>" />
    <input type="hidden" id="hidRefreshPage" value="0" />
    <input type="hidden" id="txtNonstandardpublishName" name="txtNonstandardpublishName" value="" />
    <input type="hidden" id="txtNonStandarversion" name="txtNonStandarversion" value="" />
    <input type="hidden" id="txtReason" name="txtReason" value="" />
    <input type="hidden" id="txtIe10_11_RegularMode" name="txtIe10_11_RegularMode" value="<%=bIe10_11_RegularMode%>" />
    <input type="hidden" id="txtSecondAVhistoryOpen" name="txtSecondAVhistoryOpen" value="" />
    <input type="hidden" id="txtSecondAddAVsOpen" name="txtSecondAddAVsOpen" value="" />
    <input type="hidden" id="hidIsDesktop" name="hidIsDesktop" value="<%= m_IsDesktop%>" />
    <input type="hidden" id="hidFeatureList" name="hidFeatureList" value="<%= featurelist%>" />
    <input type="hidden" id="hidEOM" name="hidEOM" value="<%=sEOM%>" />
    <input type="hidden" id="hidRTP" name="hidRTP" value="<%=sRTP%>" />
    <input type="hidden" id="hidSA" name="hidSA" value="" />
    <input type="hidden" id="hidGA" name="hidGA" value="" />
    <input type="hidden" id="hidPAAD" name="hidPAAD" value="" />
    <input type="hidden" id="hidIsPC" name="hidIsPC" value="<%=bIsPc%>" />

    <div id="divAVPropertiesdialog" style="display: none;">
        <iframe frameborder="0" name="frmAVPropertiesdialog" id="frmAVPropertiesdialog" style="height: 100%; width: 100%" frameborder="0" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divAVCategorydialog" style="display: none;">
        <iframe frameborder="0" name="frmAVCategorydialog" id="frmAVCategorydialog" style="height: 100%; width: 100%" frameborder="0" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div style="display: none;">
        <div id="iframeDialog">
            <iframe frameborder="0" name="modalDialog" id="modalDialog" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
        </div>
    </div>

    <div style="display: none;">
        <div id="divOpenSCMReport">
            <iframe frameborder="0" name="ifOpenSCMReport" id="ifOpenSCMReport"></iframe>
        </div>
    </div>

    <div id="divAddExistingAvAsShared" style="display: none;">
        <iframe frameborder="0" name="ifAddExistingAvAsShared" id="ifAddExistingAvAsShared"></iframe>
    </div>

    <div id="divAVChangeHistory" style="display: none;">
        <iframe frameborder="0" name="ifAVChangeHistory" id="ifAVChangeHistory" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divSCMCatDetails" style="display: none;">
        <iframe frameborder="0" name="ifSCMCatDetails" id="ifSCMCatDetails" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divABTInfo" style="display: none;">
        <iframe frameborder="0" name="ifABTInfo" id="ifABTInfo" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divSCMManufacturingSites" style="display: none;">
        <iframe frameborder="0" name="ifSCMManufacturingSites" id="ifSCMManufacturingSites" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
    <div id="divChinaGP" style="display: none;">
        <iframe frameborder="0" name="ifChinaGP" id="ifChinaGP" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
    <div id="divMultipleAVs" style="display: none;">
        <iframe frameborder="0" name="ifMultipleAVs" id="ifMultipleAVs" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divBasePartInformation" style="display: none;">
        <iframe frameborder="0" name="ifBasePartInformation" id="ifBasePartInformation" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div style="display: none;">
        <div id="divFeatureCreateDialog">
            <iframe frameborder="0" name="ifFeatureCreateDialog" id="ifFeatureCreateDialog"></iframe>
        </div>
    </div>

    <div id="divFeaturePropertiesDialog" style="display: none;">
        <iframe frameborder="0" name="ifFeaturePropertiesDialog" id="ifFeaturePropertiesDialog" style="height: 102%; width: 100%" frameborder="0" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divPulsarObjectPermission" style="display: none;">
        <iframe frameborder="0" name="ifPulsarObjectPermission" id="ifPulsarObjectPermission" style="height: 102%; width: 100%" frameborder="0" marginheight="0" marginwidth="0"></iframe>
    </div>

    <div id="divSCMWorksheetShowHide" style="display: none;">
        <iframe frameborder="0" name="ifSCMWorksheetShowHide" id="ifSCMWorksheetShowHide" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
    <div id="divFilter" style="display: none;">
        <iframe frameborder="0" name="ifFilter" id="ifFilter" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
</body>
</html>
<script type="text/javascript">
    $('#loadingProgress').hide();
    $('#loadingProgress').spin('medium', '#0096D6');

    $(window).on('beforeunload', function(){
        var targetUrl = window.location.href;

        $('#loadingProgress').show();
        setTimeout(function () {
            if (targetUrl === window.location.href && document.readyState !== 'loading') {
                $('#loadingProgress').hide();
            }
        }, 5000);
    });

    $(document).on('readystatechange', function () {
        if (document.readyState === 'interactive' || document.readyState === 'complete') {
            $('#loadingProgress').hide();
        }
    });
</script>
<% 
Function DayOfWeek(InputDate)

    DIM iDay, strDayName
    iDay = DatePart("w", InputDate)

    SELECT CASE iDay
    Case "1" strDayName = "Sun"
    Case "2" strDayName = "Mon"
    Case "3" strDayName = "Tue"
    Case "4" strDayName = "Wed"
    Case "5" strDayName = "Thu"
    Case "6" strDayName = "Fri"
    Case "7" strDayName = "Sat"
    END SELECT

    DayOfWeek = strDayName

End Function

    Function ValidRow(drow, BomLevel)
        Dim sKitMatlType : sKitMatlType = drow("SpsMatlType") & ""
        Dim sSaMatlType : sSaMatlType = drow("SaMatlType") & ""
        Dim sPartMatlType : sPartMatlType = drow("PartMatlType") & ""
        Dim bValidType : bValidType = False

        Dim sKitXplantStatus : sKitXplantStatus = drow("SpsXplantStatus") & ""
        Dim sSaXplantStatus : sSaXplantStatus = drow("SaXplantStatus") & ""
        Dim sPartXplantStatus : sPartXplantStatus = drow("PartXplantStatus") & ""
        Dim bValidStatus : bValidStatus = False

        Select Case BomLevel
            Case "kit"
                bValidType = ValidMaterial(sKitMatlType)
                bValidStatus = ValidStatus(sKitXplantStatus)
            Case "sa"
                bValidType = ValidMaterial(sKitMatlType) And ValidMaterial(sSaMatlType)
                bValidStatus = ValidStatus(sKitXplantStatus) And ValidStatus(sSaXplantStatus)
            Case "part"
                bValidType = ValidMaterial(sKitMatlType) And ValidMaterial(sSaMatlType) And ValidMaterial(sPartMatlType)
                bValidStatus = ValidStatus(sKitXplantStatus) And ValidStatus(sSaXplantStatus) And ValidStatus(sPartXplantStatus)
        End Select

        ValidRow = bValidType And bValidStatus

    End Function

    Function ValidMaterial( MatlType )
        Select Case UCASE(TRIM(MatlType))
            Case "HALB"
                ValidMaterial = True
            Case "FERT"
                ValidMaterial = True
            Case "ROH"
                ValidMaterial = True
            Case Else
                ValidMaterial = False
        End Select
    End Function

    Function ValidStatus( XplantStatus )
        Select Case UCASE(TRIM(XplantStatus))
            Case "C2"
                ValidStatus = True
            Case "C5"
                ValidStatus = True
            Case Else
                ValidStatus = False
        End Select
    End Function
    
    Set cn = Nothing
%>

<%
    '--- CLOSE OBJECTS: ---
    Call OpenDBConnection(PULSARDB(), False)	    'Close database connection, oConnect
%>

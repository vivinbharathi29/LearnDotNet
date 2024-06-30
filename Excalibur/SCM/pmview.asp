<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%
Server.ScriptTimeout = 480
Response.Clear
'response.Buffer = false
' --- GLOBAL & OPTIONAL INCLUDES: ---%>
<!--#INCLUDE FILE="../includes/oConnect.asp"-->
<!--#INCLUDE FILE="../includes/orsProgramMatrix.asp"-->
<%
'--- INSTANTIATE OBJECTS: ---
Call OpenDBConnection(PULSARDB(), True)			'Open database connection, oConnect.

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
Dim iFeatureCategoryID, sFeatureCategory, chAvValue
Dim iAvCount
Dim bFirstWrite
Dim bIsPc : bIsPc = False
Dim bIsPDM : bIsPDM = False
dim bIsRpdm : bIsRpdm = False
Dim bUnpublished : bUnpublished = False
Dim bShowPublishRollback : bShowPublishRollback = False
Dim m_BrandID : m_BrandID = ""
Dim m_BrandName : m_BrandName = ""
Dim m_ShortName : m_ShortName = ""
Dim m_LastPublishDt : m_LastPublishDt = ""
Dim m_BusinessID : m_BusinessID = ""
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
Dim UnpublishedAvsCount
Dim AVsWithMissingDataCount
Dim AVsWithMissingDelRootCount
Dim ShowReport
Dim sProdVersionBSAMFlag
Dim sProgramName
Dim ShowPhWebActionItems
Dim PendingAvActions : PendingAvActions = "True"
Dim sSCMCategories : sSCMCategories = ""
Dim sPMCategories : sPMCategories = ""
Dim strSeriesSummary : strSeriesSummary = ""
Dim bMultipleSeriesNumbers : bMultipleSeriesNumbers = False
Dim sFusionRequirements : sFusionRequirements = false

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
Dim m_IsOdmPinPm
Dim m_IsMarketingUser
Dim MarketingProductCount : MarketingProductCount = 0
'##############################################################################	
'
' Create Security Object to get User Info
'
	
	bIsPc = False
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

    m_IsOdmPinPm = false
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

    'add the permission from the Users and Roles to the Pulsar products
    If Not m_IsMarketingUser Then
		MarketingProductCount = rs("MarketingProductCount")
        if MarketingProductCount > 0 then
            m_IsMarketingUser = True
        end if
	End If

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

Set cmd = dw.CreateCommAndSP(cn, "spGetProductVersion")
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
    if rs("SMID") & "" <> "" then
        PMID = PMID & "_" & rs("SMID")
    end if
    PMID = "_" & PMID & "_"
	strSCMPath = rs("SCMPath") & ""
    strProdType = rs("TypeID") & ""
    sProdVersionBSAMFlag = rs("BSAMFlag") & ""
    sFusionRequirements = rs("FusionRequirements") 

    if (rs("ODMPIMPMID") & "") = CurrentUserID then
        m_IsOdmPinPm = true
    end if

End If				
rs.Close

Function PrepForWeb( value )
	Dim myString
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
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

'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
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

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!--meta http-equiv="X-UA-Compatible" content="IE=8" /-->

    <title>Excalibur PM View</title>
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/scm/style.css" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/scm/sample.css" />
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script src="<%= AppRoot %>/includes/client/json2.js" type="text/javascript"></script>
    <script type="text/javascript" src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js"></script>
    <script type="text/javascript" src="<%= AppRoot %>/includes/client/popup.js"></script>
    <script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/spin.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/jquery.spin.js" type="text/javascript"></script>
    <script src="/pulsar2/js/userfavorite.js" type="text/javascript"></script>

    <script type="text/javascript">
    <!--
    $(function () {
        $("#scmLoading").hide();

        var strFavorites = "," + $("#txtFavs").val();
        var found = strFavorites.indexOf(",P" + $.trim($("#txtID").val()) + ",");

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

        $("tr").on("contextmenu", function () {
            return false;
        });

    });

    function adjustWidth(percent) {
        return document.documentElement.offsetWidth * (percent / 100);
    }

    function adjustHeight(percent) {
        return (document.documentElement.offsetHeight * (percent / 100));
    }

    String.prototype.trim = function () {
        return this.replace(/^\s+|\s+$/g, "");
    }

    String.prototype.ltrim = function () {
        return this.replace(/^\s+/, "");
    }

    String.prototype.rtrim = function () {
        return this.replace(/\s+$/, "");
    }

    function AddAV(PVID, BID, CurrentUserId) {
        var strID
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/avFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID, "", "dialogWidth:900px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        window.location.reload();
    }

    function LocalizeAVs(PVID, strSeriesSummary, bMultipleSeriesNumbers, BID, UserName, KMATValue) {
        //if (bMultipleSeriesNumbers == 'True') {
        var strID;
        var url = "<%=AppRoot %>/SCM/LocalizeAVsSeriesNumFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID + "&strSeriesSummary=" + strSeriesSummary + "&UserName=" + UserName + "&KMAT=" + KMATValue;
        modalDialog.open({ dialogTitle: 'Add Localized AVs', dialogURL: '' + url + '', dialogHeight: 460, dialogWidth: 390, dialogResizable: true, dialogDraggable: true });
    }

    function OpenLocalizeAVs(strPath) {
        if (strPath == null) {
            return;
        } else {
            //close LocalizeAVs open modal dialog: ---
            modalDialog.cancel();

            //configure width and height and then open modal dialog with selected AV
            var DlgWidth = 1100;
            if (adjustWidth(95) < DlgWidth) {
                DlgWidth = adjustWidth(95);
            }
            var DlgHeight = adjustHeight(95);

            //var retVal = window.parent.showModalDialog(strPath, "", "dialogWidth:" + DlgWidth + "px;dialogHeight:" + DlgHeight + "px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
            modalDialog.open({ dialogTitle: 'Add Localized AVs', dialogURL: '' + strPath + '', dialogHeight: DlgHeight, dialogWidth: DlgWidth, dialogResizable: true, dialogDraggable: true });
        }
    }

    function ReloadLocalizeAVs(retVal) {
        if (retVal == "YES") {
            //close LocalizeAVs open modal dialog: ---
            modalDialog.cancel();
            //reload parent page
            window.location.reload();
        }
    }

    function StructureBOM(User, BID, KMAT, BusinessID, PVID) {
        //alert(User + ',' + BID + ',' + KMAT)
        var strID;
        var url = "<%=AppRoot %>/SCM/StructureBOMFrame.asp?Mode=add&User=" + User + "&BID=" + BID + "&KMAT=" + KMAT + "&BusinessID=" + BusinessID + "&PVID=" + PVID;
        modalDialog.open({ dialogTitle: 'Structure BOM', dialogURL: '' + url + '', dialogHeight: 250, dialogWidth: 410, dialogResizable: true, dialogDraggable: true });
        /*strID = window.parent.showModalDialog(url, "", "dialogWidth:360px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        window.location.reload();*/
    }

    function AvActionScorecard(UserId) {
        var strID;
        var url = "<%=AppRoot %>/SCM/AvActionScorecardFrame.asp?UserID=" + UserId;
        strID = window.parent.showModalDialog(url, "", "dialogWidth:549px;dialogHeight:475px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        window.close();
    }

    function UploadAvData(BID, PVID, UserName) {
        var strID;
        var url = "<%=AppRoot %>/SCM/UploadAvDataFrame.asp?PVID=" + PVID + "&BID=" + BID + "&UserName=" + UserName
        strID = window.parent.showModalDialog(url, "", "dialogWidth:475px;dialogHeight:500px;edge: Sunken;center:Yes; help: No;resizable: No;status: No");
        window.location.reload();
    }

    function AddSAs() {
        var strID;
        strID = window.open("<%=AppRoot %>/MobileSE/SubAssembly.asp?SAType=223&Business=100&Family=100", "", "width=1100,height=800,toolbar=0,resizable=1,scrollbars=1")
    }

    function QuickSearch() {
        var strID;
        var url = "<%=AppRoot %>/MobileSE/QuickSearch.asp"
        modalDialog.open({ dialogTitle: 'Quick Search', dialogURL: '' + url + '', dialogHeight: 175, dialogWidth: 350, dialogResizable: true, dialogDraggable: true });
        /*strID = window.parent.showModalDialog(url, "", "dialogWidth:220px;dialogHeight:50px;edge: Sunken;center:Yes; help: No;resizable: No;status: No");
        window.close()*/
    }

    function UploadScm(BID) {
        var strID
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/UploaderFrame.asp?BID=" + BID, "", "dialogWidth:500px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        if (strID == null)
            return;
        window.location.reload();
    }

    function EditKMAT(PVID, BID) {
        var strID
        //window.parent.showModalDialog("<%=AppRoot %>/SCM/kmatFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        modalDialog.open({ dialogTitle: 'Edit Program Data', dialogURL: "<%=AppRoot %>/SCM/kmatFrame.asp?Mode=add&PVID=" + PVID + "&BID=" + BID + "", dialogHeight: 700, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
        //window.location.reload();
    }

    function showOrphanedAvReport(BID) {
        var strID
        strID = window.open("<%=AppRoot %>/SCM/avOrphansFrame.asp?BID=" + BID, "OrphanedAVReport", "width=450,height=500,toolbar=0,resizable=1")
    }

    function showUnpublishedAVs(BID) {
        var strID;
        strID = window.parent.showModalDialog("<%= AppRoot %>/SCM/UnpublishedAVsFrame.asp?BID=" + BID, "", "dialogWidth:650px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
        //strID = window.open("<%= AppRoot %>/SCM/UnpublishedAVsFrame.asp?BID=" + BID, "", "width=650,height=510,toolbar=0,resizable=0,scrollbars=0,center=1")
        window.location.reload();
    }

    function showAVsWithMissingData(BID, IsPDM) {
        var strID;
        modalDialog.open({ dialogTitle: 'AVs Missing Corporate Data', dialogURL: '<%= AppRoot %>/SCM/AvsMissingDataFrame.asp?BID=' + BID + '&IsPDM=' + IsPDM + '', dialogHeight: 600, dialogWidth: (GetWindowSize('width')), dialogResizable: true, dialogDraggable: true });

        /*strID = window.parent.showModalDialog("<%= AppRoot %>/SCM/AvsMissingDataFrame.asp?BID=" + BID + "&IsPDM=" + IsPDM, "", "dialogWidth:1325px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
        //strID = window.open("<%= AppRoot %>/SCM/AvsMissingDataFrame.asp?BID=" + BID, "", "width=650,height=510,toolbar=0,resizable=0,scrollbars=0,center=1")
        window.location.reload();*/
    }

    function showAVsWithMissingDelRoot(BID, PVID) {
        modalDialog.open({ dialogTitle: 'AVs Missing Deliverable Root', dialogURL: '<%= AppRoot %>/SCM/AvsMissingDelRootFrame.asp?BID=' + BID + '&PVID=' + PVID + '', dialogHeight: 560, dialogWidth: 700, dialogResizable: true, dialogDraggable: true });
        /*var strID;
        strID = window.parent.showModalDialog("<%= AppRoot %>/SCM/AvsMissingDelRootFrame.asp?BID=" + BID + "&PVID=" + PVID, "", "dialogWidth:650px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
        window.location.reload();*/
        //        var returnValue;
        //        returnValue = window.parent.showModalDialog("<%= AppRoot %>/SCM/AvsMissingDelRootFrame.asp?BID=" + BID + "&PVID=" + PVID, "", "dialogWidth:650px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
        //        if (returnValue != undefined) {
        //            window.location.reload();
        //        } else {
        //            showAVsWithMissingDelRoot(BID, PVID);
        //        }
    }

    function SelectTab(strStep, blnLoad) {
        var i;
        var expireDate = new Date();
        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "PMTab=" + strStep + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";
        CurrentState = strStep;
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
            //jsrsExecute("<%=AppRoot %>/FavoritesRSupdate.asp", myCallback, "UpdateFavs", Array(strFavorites, String(FavCount), txtUser.value));
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
    function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
        $("#modalDialog").attr("width", "98%");
        $("#modalDialog").attr("height", "98%");
        $("#modalDialog").attr("src", QueryString);
        $("#iframeDialog").dialog("option", "title", Title);
        $("#iframeDialog").dialog("open");
    }

    function ShowProperties_Product(DisplayedID, Clone, FusionRequirements) {
        var shouldClone;

        if (Clone == 1) {
            shouldClone = "&Clone=1";
        } else {
            shouldClone = "";
        }

        ShowPropertiesDialog("<%=AppRoot %>/mobilese/today/programs.asp?HWPM=0&ID=" + DisplayedID + shouldClone + "&Pulsar=" + FusionRequirements, "Product Properties", 980, 800);
    }
    function ClosePropertiesDialog(strID) {
        $("#iframeDialog").dialog("close");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function CloseIframeDialog() {
        $("#iframeDialog").dialog("close");
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
        window.location.replace("<%=AppRoot %>/scm/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&BID=" + ProductBrandID);
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
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/catFrame.asp?Mode=edit&PVID=" + node.pvid + "&FCID=" + node.fcid + "&BID=" + node.bid, "", "dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        document.location.reload();
    }

    function ARMO() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }
        node.style.color = "red";
        node.style.cursor = "hand";
        window.status = "PVID:" + node.pvid + "BID:" + node.bid + "AVID:" + node.avid;
    }


    function ARMOut() {
        var node = window.event.srcElement;
        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }
        node.style.color = "black";
    }

    function AROC() {
        var node = window.event.srcElement;
        if (node.nodeName.toLowerCase() !== 'td')
            return;

        while (node.nodeName.toLowerCase() != "tr") {
            node = node.parentElement;
        }
        ShowAvDetails(node.pvid, node.avid, node.bid);
    }

    function ARMD(ProductVersionID, AvDetailID) {
        if (event.button == 2) {
            RtClickMenu();
            return;
        }
    }

    function ShowAvDetails(ProductVersionID, AvDetailID, ProductBrandID) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/avFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&UserID=" + txtUser.value, "", "dialogWidth:900px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        document.location.reload();
    }

    var oPopup = window.createPopup();
    oPopup.document.createStyleSheet("<%=AppRoot %>/style/menu.css");
    var _avDetailID;
    var _productVersionID;
    var _productBrandID;

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

        oPopup.document.body.innerHTML = window.PopUpMenu.innerHTML;

        oPopup.show(lefter, topper, 150, 85, document.body);

        if (node.status == "A")
            oPopup.document.body.all["activate"].style.display = "none";
        else
            oPopup.document.body.all["obsolete"].style.display = "none";

        if (node.status == "H")
            oPopup.document.body.all["hide"].style.display = "none";
        else
            oPopup.document.body.all["unhide"].style.display = "none";

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

    function MenuSetObsolete() {
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'O');
    }

    function MenuSetActive() {
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'A');
    }

    function MenuSetHidden() {
        SetAvStatus(_productVersionID, _avDetailID, _productBrandID, 'H');
    }

    function MenuDelete() {
        DeleteAv(_productVersionID, _avDetailID, _productBrandID);
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
                    strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/avDelete.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                    var row = document.getElementById("AV" + AvDetailID);
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

    function MenuEditMrDates() {
        EditMrDates(_productVersionID, _productBrandID, _avDetailID);
    }

    function SetAvStatus(ProductVersionID, AvDetailID, ProductBrandID, Status) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/avSetStatus.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&Status=" + Status, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        var row = document.getElementById("AV" + AvDetailID);
        row.className = Status;
    }

    function DeleteAv(ProductVersionID, AvDetailID, ProductBrandID) {
        var strID;
        var response = confirm("Are you sure you want to delete this record?");
        if (response) {
            strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/avDelete.asp?PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
            var row = document.getElementById("AV" + AvDetailID);
            if ((row.className == 'O') || (hidLastPublishDt.value == '') || (hidBusinessID.value != 1))
                row.style.display = "none";
            else
                row.className = 'O';
        }
    }

    function ScmPublishRollback(ProductBrandID) {
        var strID;
        var response = confirm("You are about to permanently delete all records of the most recent SCM publish. \n Do you want to continue?");
        if (response) {
            strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/scmPublishRollback.aspx?PBID=" + ProductBrandID, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
            document.location.reload();
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
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/mrDatesFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&AVID=" + AvDetailID, "", "dialogWidth:325px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        //document.location.reload();
    }

    function ImportAv(ProductVersionID, ProductBrandID) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/pasteFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:410px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        document.location.reload();
    }

    function CopyScm(ProductVersionID, ProductBrandID) {
        var strID;
        strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/copyFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        document.location.reload();
    }

    function LinkAv(ProductVersionID, ProductBrandID, AvID) {
        var sMode = null;
        var strID = null;
        if (AvID == null) {
            sMode = "LinkFrom";
            strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/linkFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&Function=" + sMode, "", "dialogWidth:500px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        }
        else {
            sMode = "LinkTo";
            strID = window.parent.showModalDialog("<%=AppRoot %>/SCM/linkFrame.asp?PVID=" + ProductVersionID + "&BID=" + ProductBrandID + "&Function=" + sMode, "", "dialogWidth:500px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        }

        document.location.reload();

    }

    function PublishScm(ProductVersionID, ProductBrandID) {
        if (ValidateData(ProductBrandID, true) == 1) {
            var answer = confirm("Click OK to continue Publishing the SCM");
            if (answer)
                location.href = "/iPulsar/ExcelExport/SCM.aspx?BID=" + ProductBrandID + "&PVID=" + ProductVersionID + "&Publish=True";
        }
    }

    function ValidateData(ProductBrandID, chked) {
        if (chked) {
            var parameters = "function=ValidateData&PBID=" + ProductBrandID;
            var request = null;
            //Initialize the AJAX variable.
            if (window.XMLHttpRequest) {        //Are we working with mozilla
                request = new XMLHttpRequest(); //Yes -- this is mozilla.
            } else { //Not Mozilla, must be IE
                request = new ActiveXObject("Microsoft.XMLHTTP");
            } //End setup Ajax.
            request.open("POST", "<%=AppRoot %>/SCM/SCM_PM_Publish_ValidateData.asp", false);
            request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
            request.send(parameters);
            if (request.responseText == 'Success') {
                return 1;
            } else if (request.responseText.indexOf('Mass Prodcution First Customer Ship') > 0) {
                alert(request.responseText);
                return 1;
            } else {
                document.getElementById("chkPublish").checked = false;
                alert(request.responseText);
                return 0;
            }
        }
    }

    function clearAvHistory(ProductBrandID) {
        var answer = confirm("Are you sure you want to clear the SCM Change Log?");

        if (answer) {
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

    function clearAvHistoryCallback(returnstring) {
        alert("Change Log Cleared.");
    }


    //************************************************************************
    function ExportProjectMatrix() {
        $("#trCompare").show();
        $("#popup_submit").attr("disabled", false);
        $("#chkNewMatrix").attr("checked", false);
        $("#chkPublish").attr("checked", false);

        if ($("#selCompareDt option").length == 0) {
            $("#trCompare").hide();
            $("#chkNewMatrix").attr("checked", true);
        }

        modalDialog.open({ dialogTitle: 'Program Matrix Options', dialogDivID: 'modal_programmatrix', dialogHeight: 200, dialogWidth: 350, dialogResizable: false, dialogDraggable: true });
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
                    }
                    else {
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
        }
        else {
            alert("Unable to save this tab as the default.");
        }
    }

    function gochange_onmouseover() {
        window.event.srcElement.style.cursor = "hand";
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
        var url = "<%=AppRoot %>/SCM/DocKitDataFrame.asp?Mode=add&KMAT=" + KMAT;
        modalDialog.open({ dialogTitle: 'Check Doc Kit Status', dialogURL: '' + url + '', dialogHeight: 700, dialogWidth: 645, dialogResizable: true, dialogDraggable: true });

    }

    function FilterSCMByCategory(BusinessID, BID, UserID, PVID) {
        var Categories = document.getElementById("hidSCMCategories")

        var url = "<%=AppRoot %>/SCM/FilterByCategoryFrame.asp?BusinessID=" + BusinessID + "&BID=" + BID + "&UserID=" + UserID + "&PVID=" + PVID + "&Categories=" + Categories.value;

        var retValue;
        retValue = window.parent.showModalDialog(url, "", "dialogWidth:315px;dialogHeight:475px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        if (retValue != undefined) {
            window.location.replace("<%=AppRoot %>/scm/pmview.asp?ID=" + PVID + "&BID=" + BID + "&SCMCategories=" + retValue + "&Class=Arrow1");
        }
    }

    function FilterPMByCategory(BusinessID, BID, UserID, PVID) {
        var Categories = document.getElementById("hidPMCategories")

        var url = "<%=AppRoot %>/SCM/FilterByCategoryFrame.asp?BusinessID=" + BusinessID + "&BID=" + BID + "&UserID=" + UserID + "&PVID=" + PVID + "&Categories=" + Categories.value;

        var retValue;
        retValue = window.parent.showModalDialog(url, "", "dialogWidth:315px;dialogHeight:475px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        if (retValue != undefined) {
            window.location.replace("<%=AppRoot %>/scm/pmview.asp?ID=" + PVID + "&BID=" + BID + "&PMCategories=" + retValue + "&Class=Arrow1");
        }
    }


    function closeModalDialog(bReload) {
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

    <style type="text/css">
        #Table1 thead th {
        background-color:wheat;
        }

        </style>
</head>
<body onload="window_onload();">
    <%IF sFusionRequirements = true THEN %>
    <font size="4"><strong id="productNameTitle"><%= sProductName%> Information (Pulsar)</strong></font>
    <br />
    <br />
    <%ELSE%>
    <font size="4"><strong id="productNameTitle"><%= sProductName%> Information (Legacy)</strong></font>
    <br />
    <br />
    <%END IF%>
    <%If bIsPc Then Response.Write strNoticetable %>
    <%    if (clng(request("ID")) = 344 or clng(request("ID")) = 347 or clng(request("ID")) = 1107) and (not badministrator) then %>
    <td nowrap id="EditLink" style="display: none"></td>
    <%elseif ((m_IsToolPm )and trim(strProdType) = "2") then%>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <%IF sFusionRequirements = true THEN %>
        <font size="1"><a href="javascript:ShowProperties_Product(<%=PVID%>,0,1)">Edit Product</a></font>
        <%ELSE%>
        <font size="1"><a href="javascript:ShowProperties_Product(<%=PVID%>,0,0)">Edit Product</a></font>
        <%END IF%>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsActivePm and trim(strProdType) <> "2" then%>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <%IF sFusionRequirements = true THEN %>
        <font size="1"><a href="javascript:ShowProperties_Product(<%=PVID%>,0,1)">Edit Product</a></font>
        <%ELSE%>
        <font size="1"><a href="javascript:ShowProperties_Product(<%=PVID%>,0,0)">Edit Product</a></font>
        <%END IF%>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsCommodityPM and trim(strProdType) <> "2" then ' %>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsSCFactoryEngineer and trim(strProdType) <> "2" then %>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,2)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsAccessoryPM and trim(strProdType) <> "2" then %>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,3)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsHardwarePm and trim(strProdType) <> "2" then %>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,4)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%elseif m_IsTestLead and trim(strProdType) <> "2" then%>
    <%if trim(PVID) <> "-1" then%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1)">Edit Product</a></font>
        <font face="verdana" size="1" color="black"> | </font>
    </td>
    <%else%>
    <td nowrap id="EditLink" style="display: none"></td>
    <%end if%>
    <%else%>
    <td nowrap id="EditLink" style="display: none">
        <font size="1">
            <!--Contact the PM to edit product properties-->
        </font>
    </td>
    <%end if%>
    <span id="loadingProgress"></span>
    <span id="RFLink" style="display: none"><a href="javascript:RemoveFavorites(<%=PVID%>)">
        <font face="verdana" size="1">Remove From Favorites</font></a>|</span><span id="AFLink"
            style="display: none"><a href="javascript:AddFavorites(<%=PVID%>)"><font face="verdana"
                size="1">Add To Favorites</font></a> |</span><span id="StatusLink" style="display: none">
                    <a href="<%=AppRoot %>/Productstatus.asp?Product=<%=sDisplayedProductName%>&ID=<%=PVID%>">
                        <font face="verdana" size="1">Real-Time Status Report</font></a>|</span>
    <%if strDisplayedList <> CurrentUserDefaultTab and trim(strProdType) <> "2" then%>
    <span id="DefaultTabLink"><a href="javascript:SetDefaultDisplay('<%=strDisplayedList%>',<%=CurrentUserID%>)">
        <font face="verdana" size="1">Set Default List</font></a></span>
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
Set cmd = dw.CreateCommAndSP(cn, "spListBrands4Product")
dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, PVID
dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
Set rs = dw.ExecuteCommAndReturnRS(cmd)
	
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
                                SCM | <a href="#" onclick="ShowReport('pm');">Program Matrix</a>
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
	Do Until rs.EOF
		'Response.Write "<td><a href=""javascript:void(0)"">" & server.HTMLEncode(rs("schedule_name")) & "</a></td>"
		If Not bFirstWrite Then
			Response.Write "&nbsp;|&nbsp;"
		End If
			
		If (Request("BID") = "" And m_BrandID = "") Or (CLng(rs("ProductBrandID")) = CLng(Request("BID"))) Then
			m_BrandID = rs("ProductBrandID")			
			m_BrandName = rs("Name")
			m_ShortName = rs("streetname2") & " " & rs("SeriesSummary")
			m_LastPublishDt = rs("LastPublishDt") & ""
			If rs("LastPublishDt") & "" = "" Then
			    bUnpublished = True
			ElseIf DateDiff("h", m_LastPublishDt, Now()) <= 8 Then
			    bShowPublishRollback = True
			End If
			m_BusinessID = rs("BusinessID") & ""
			Response.Write server.HTMLEncode(m_BrandName)
		Else
			Response.Write "<a href=""javascript:BrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"
		End If
		bFirstWrite = False
		rs.MoveNext
	Loop
	
	rs.Close
	
	
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
                                Active | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
		Case "obsolete"
                                %>
                                <a href="#" onclick="ShowObsolete('active');">Active</a> | Obsolete | <a href="#"
                                    onclick="ShowObsolete('all');">All</a>
                                <%
        Case "all"              'all is no longer a default when no cookie is set yet
                                %>
                                <a href="#" onclick="ShowObsolete('active');">Active</a> | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | All
                                <%
		Case Else               'active is no longer a default when no cookie is set yet
                                %>
                                Active | <a href="#" onclick="ShowObsolete('obsolete');">Obsolete</a> | <a href="#"
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
	                            Select Case Request("SCMCategories")
		                            Case ""
                                %>
                                <a href="#" onclick="FilterSCMByCategory(<%=m_BusinessID%>,<%=m_BrandID%>,<%=m_CurrentUserId%>,<%=PVID%>);">By Category</a>
                                <%
		                            Case Else
		                              sSCMCategories = Request("SCMCategories")
                                %>
                                <a href="#" style="color: Red;" onclick="FilterSCMByCategory(<%=m_BusinessID%>,<%=m_BrandID%>,<%=m_CurrentUserId%>,<%=PVID%>);">By Category</a>
                                <%
	                            End Select
                                %>
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
        <a href="#" onclick="EditKMAT(<%=PVID%>, <%=m_BrandID%>);">Program Data</a> <span style="color: Red; font-weight: bold; cursor: pointer;">(Action Items Off)</span>        |
           <%End If%>
        <%End If%>
        <%If bIsPc or m_IsMarketingUser Then%>
        <a href="/iPulsar/ExcelExport/ScmWorkingTemplate.aspx?PBID=<%=m_BrandID%>">SCM Wrksht</a> | <a href="#" onclick="UploadScm(<%=m_BrandID%>);">Upload Wrksht</a>        |
        <%End If%>
        <a href="/iPulsar/ExcelExport/SCM.aspx?BID=<%=m_BrandID%>&PVID=<%=PVID%>">Export
                SCM</a> |
        <%If bIsPc or (m_IsSCMPublishCoordinator and m_BusinessID=2) Then%>
        <a href="#" onclick="PublishScm(<%=PVID%>, <%=m_BrandID%>);">Publish SCM</a> |
        <%End If%>
        <a href="#" onclick="ExportProjectMatrix(<%=m_BrandID%>)">Program Matrix</a> |
        <%If bUnpublished Then %>
        <a href="javascript:clearAvHistory(<%=m_BrandID%>)">Clear Change Log</a> |
        <%End If 'bUnpublished %>
        <a href="<%=AppRoot %>/SCM/ChangeLog.aspx?ID=<%=PVID%>&BID=<%=m_BrandID%>" target="New">Show Change Log</a> | <a href="<%=AppRoot %>/SCM/mrDatesView.asp?ID=<%=PVID%>&BID=<%=m_BrandID%>"
            target="New">Show MR Dates</a>
        <%  If bShowPublishRollback Then %>
        | <a href="javascript:ScmPublishRollback(<%=m_BrandID%>)">Rollback SCM Publish</a>
        <%  End If 'bShowPublishRollback %>
        <br />
        <br />
        <span style="color: Black; font-weight: bold; cursor: pointer;">Tools:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
        <a href="#" onclick="QuickSearch()">QuickSearch</a>
        <%If bIsPc Then%>
        | <a href="#" onclick="AddAV(<%=PVID%>, <%=m_BrandID%>, <%=m_CurrentUserId%>);">Add New AV</a>
        | <a href="#" onclick="LocalizeAVs(<%=PVID%>, '<%=strSeriesSummary%>', '<%=bMultipleSeriesNumbers%>', <%=m_BrandID%>, '<%=CurrentUserName%>','<%=sKmat%>');">Add Localized AVs</a>
        | <a href="/iPulsar/ExcelExport/HPRPN.aspx?BID=<%=m_BrandID%>">Export RPN</a>
        | <a href="#" onclick="UploadAvData(<%=m_BrandID%>,<%=PVID%>,'<%=CurrentUserName%>')">Import AV Data</a>
        | <a href="#" onclick="AddSAs()">Assign SAs</a>
        | <a href="#" onclick="StructureBOM('<%=CurrentUserName%>', <%=m_BrandID%>, '<%=sKmat%>', <%=m_BusinessID%>, <%=PVID%>)">Structure BOM</a>
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
        <% End if 'orphanCount
         End If
         If bIsPc or bIsPDM Then %>
        <%If ShowPhWebActionItems = "True" And PendingAvActions = "True" Then%>
        <a href="/iPulsar/ExcelExport/PendingPhWebActions.aspx?PVID=<%=PVID%>">Pending
            AV Actions</a> <span style="color: Red; font-weight: bold; cursor: pointer;">(Beta)</span> |
        <% End If 'ShowPhWebActionItems
         End If %>
        <a href="#" onclick="AvActionScorecard(<%=m_CurrentUserId%>);">AV Action Scorecard</a>
        <span style="color: Red; font-weight: bold; cursor: pointer;">(Beta)</span>
        <%If bIsPc Then%>
        | <a href="#" onclick="DocKitData('<%=sKmat%>');">Check Doc Kit Status</a>
        <%If UnpublishedAvsCount > 0 Then %>
        | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
            onclick="showUnpublishedAVs(<%=m_BrandID%>)">Unpublished AVs</span>
        <% Else %>
        | <a href="#" onclick="showUnpublishedAVs(<%=m_BrandID%>)">Unpublished AVs</a>
        <% End if 'Unpublished AVs %>
        <%End If
        If bIsPc or bIsPDM Then %>
        <% If AVsWithMissingDataCount > 0 Then %>
            | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
                onclick="showAVsWithMissingData(<%=m_BrandID%>,'<%=bIsPDM%>')">AVs Missing Corporate Data</span>
        <% Else %>
            | <a href="#" onclick="showAVsWithMissingData(<%=m_BrandID%>,'<%=bIsPDM%>')">AVs Missing Corporate Data</a>
        <% End if 'AVs Missing Corporate Data %>
        <span style="display: none"><a href="#" onclick="CopyScm(<%=PVID%>, <%=m_BrandID%>);">Copy Existing SCM</a> |</span>
        <% If AVsWithMissingDelRootCount > 0 Then %>
            | <span style="color: Red; font-weight: bold; text-decoration: underline; cursor: pointer;"
                onclick="showAVsWithMissingDelRoot(<%=m_BrandID%>,<%=PVID%>)">AVs Missing Deliverable Root</span>
        <% Else %>
            | <a href="#" onclick="showAVsWithMissingDelRoot(<%=m_BrandID%>,<%=PVID%>)">AVs Missing Deliverable Root</a>
        <% End if 'AVs Missing Deliverable Root %>
        <%End If%>
        <%If ShowReport = "scm" Or ShowReport = "" Then
        %>
        <p>
            <span style="font-size: x-small; color: Red; font-weight: bold;">The information on
                this page is working data only. <a href="<%=strScmPath%>" target="new">Click here</a>
                to go to the published SCM</span>
        </p>
        <% 
        End If
    End If 'ShowReport = "scm" Or ShowReport = "pm"
        %>
    </span>
    <p>
        <span style="font-size: x-small; font-weight: bold;">
            <%= m_ShortName%>
            <%
            Select Case ShowReport
                Case "scm"
                    Response.Write "&nbsp;-&nbsp;Supported Configuration Matrix <span id=""lblAvCount"" style=""font-size:xx-small; font-weight: normal;""></span><span style=""font-size:small; color: Red; font-weight: bold;"">&nbsp; - Global GBU View</span>"
                Case "pm"
                    Response.Write "&nbsp;-&nbsp;Program Matrix"
                Case Else
                    Response.Write "&nbsp;-&nbsp;Supported Configuration Matrix<span style=""font-size:small; color: Red; font-weight: bold;""> - Global GBU View<span>"
            End Select
            %></span>
        <div id="scmLoading">
            <span style="font: Bold x-small Tahoma; color: Red; text-decoration: blink">Loading
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

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectScmDetail")
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
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, "H"
	Case "obsolete"
		dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "O"
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
    Case "all" 'all is no longer a default when no cookie is set yet
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, ""
        dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, ""
	Case Else 'Active is now a default when no cookie is set yet
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
		dw.CreateParameter cmd, "@p_Status2", adChar, adParamInput, 1, "H"
End Select

'response.Write("<div>SCMCategories: " & request("SCMCategories") & "</div>")
If request("SCMCategories") <> "" Then
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, Request("SCMCategories")
Else
    dw.CreateParameter cmd, "@p_Categories", adVarchar, adParamInput, 500, ""
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
        <div id="GridViewContainer" class="GridViewContainer" style="width: 100%; height: 100%">
            <table id="Table1" class="Table" width="100%">
                <col width="95">
                <col width="220">
                <col width="150">
                <col width="150">
                <col width="20" align="center">
                <col width="75" align="center" />
                <col width="75" align="center" />
                <col width="75" align="center" />
                <col width="125">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="50">
                <col width="40" align="center">
                <col width="40" align="center">
                <col width="40" align="center">
                <col width="40" align="center">
                <col width="50" align="center">
                <col width="50" align="center">
                <col width="75" align="center" />
                <thead>
                    <tr class="FrozenHeader">
                        <th>AV No.</th>
                        <th>GPG Description</th>
                        <th>Marketing Description<br />
                            (40 Char GPSy)</th>
                        <th>Marketing Description<br />
                            (100 Char PMG)</th>
                        <th>Program Version</th>
                        <th>Select Availability<br />
                            (SA) Date </th>
                        <th>General Availability<br />
                            (GA) Date </th>
                        <th>End of Manufacturing<br />
                            (EM) Date</th>
                        <th>Configuration Rules</th>
                        <th>AVID</th>
                        <th>Group 1</th>
                        <th>Group 2</th>
                        <th>Group 3</th>
                        <th>Group 4</th>
                        <th>Group 5</th>
                        <th>IDS-SKUs</th>
                        <th>IDS-CTO</th>
                        <th>RCTO-SKUs</th>
                        <th>RCTO-CTO</th>
                        <% If (sProdVersionBSAMFlag = "True") Then %>
                        <th>BSAM SKUS</th>
                        <th>BSAM -B parts</th>
                        <% End If %>
                        <th>UPC</th>
                        <th>Weight<br />
                            (in oz)</th>
                        <th>Global Series Config Plan<br />
                            for End of<br />
                            Manufacturing (PE) Date</th>
                    </tr>
                </thead>
                <tbody>
                    <%
	Do Until rs.EOF
	   If (iFeatureCategoryID <> rs("FeatureCategoryID")) Then
			iFeatureCategoryID = rs("FeatureCategoryID")
			sFeatureCategory = rs("AvFeatureCategory")
                    %>
                    <tr class="FeatureCategory" pvid="<%=PVID%>" fcid="<%=iFeatureCategoryID%>" bid="<%=m_BrandID%>" onmouseover="return FRMO()" onmouseout="return FRMOut()" onclick="return FROC()">
                        <td colspan="2">
                            <span style="font-weight: bold;"><%=sFeatureCategory%></span>
                            <%If bIsPc Then%>
                            <br />
                            <input type="checkbox" id="chkAll<%=rs("FeatureCategoryID")%>" name="chkAll<%=rs("FeatureCategoryID")%>"
                                style="width: 16px; height: 16px" onclick="chkAll_onclick()">
                            <input type="button" value="Delete" id="btnDelete<%=rs("FeatureCategoryID")%>" name="btnDelete"
                                class="button2" onclick="DeleteMultipleAvs(<%=PVID %>,<%=m_BrandID %>)">
                            <%End If%>
                        </td>
                        <td><%= PrepForWeb(rs("CategoryMarketingDescription"))%></td>
                        <td colspan="5"></td>
                        <td><%= PrepForWeb(rs("CategoryRules"))%></td>
                        <% If (sProdVersionBSAMFlag = "True") Then %>
                        <td colspan="15"></td>
                        <% Else %>
                        <td colspan="13"></td>
                        <% End If %>
                    </tr>
                    <%
            Response.Flush
		End If
		If rs.Fields(0).Value & "" <> "" Then
                    %>
                    <tr class='<%=rs("Status")%>' id="AV<%=rs.Fields(0).Value%>" pvid="<%=PVID%>" avid="<%=rs.Fields(0).Value%>"
                        status="<%=rs("Status")%>" bid="<%=rs("ProductBrandID")%>" onmouseover="return ARMO()" onmouseout="return ARMOut()"
                        onmousedown="return ARMD(<%=PVID%>,<%=rs.Fields(0).Value%>)" onclick="return AROC()">
                        <td nowrap>
                            <%If bIsPc Then%>
                            <input class='<%=rs("FeatureCategoryID")%>' type='checkbox' id='chkAv<%=rs("FeatureCategoryID")%>'
                                name='chkAv<%=rs("AvDetailID")%>' value='<%=rs("FeatureCategoryID")%>' style="width: 16px; height: 16px"
                                onclick="return chkAv_onclick()">
                            <%End If%>
                            <%=PrepForWeb(rs("AvNo"))%></td>
                        <%if rs("GPGDescSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("GPGDescription"))%></td>
                        <%else%>
                        <td nowrap><%=PrepForWeb(rs("GPGDescription"))%></td>
                        <%end if %>
                        <%if rs("MktDescSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("MarketingDescription"))%></td>
                        <%else%>
                        <td nowrap><%=PrepForWeb(rs("MarketingDescription"))%></td>
                        <%end if %>
                        <%if rs("MktDescPMGUSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("MarketingDescriptionPMG"))%></td>
                        <%else%>
                        <td nowrap><%=PrepForWeb(rs("MarketingDescriptionPMG"))%></td>
                        <%end if %>
                        <td><%=PrepForWeb(sProgramVersion)%></td>
                        <%if rs("CplBlindSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("CPLBlindDt"))%></td>
                        <%else%>
                        <td><%=PrepForWeb(rs("CPLBlindDt"))%></td>
                        <%end if %>
                        <%if rs("GeneralAvailSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("RTPDate"))%></td>
                        <%else%>
                        <td><%=PrepForWeb(rs("GeneralAvailDt"))%></td>
                        <%end if %>
                        <%if rs("RasDiscoSysUpdate") = 0 then%>
                        <td style="color: Gray"><%=PrepForWeb(rs("RASDiscontinueDt"))%></td>
                        <%else%>
                        <td><%=PrepForWeb(rs("RASDiscontinueDt"))%></td>
                        <%end if %>
                        <td nowrap><%=PrepForWeb(rs("ConfigRules"))%></td>
                        <td nowrap><%=PrepForWeb(rs("AvId"))%></td>
                        <td nowrap><%=PrepForWeb(rs("Group1"))%></td>
                        <td nowrap><%=PrepForWeb(rs("Group2"))%></td>
                        <td nowrap><%=PrepForWeb(rs("Group3"))%></td>
                        <td nowrap><%=PrepForWeb(rs("Group4"))%></td>
                        <td nowrap><%=PrepForWeb(rs("Group5"))%></td>
                        <td><%=PrepForWeb(rs("IdsSkus_YN"))%></td>
                        <td><%=PrepForWeb(rs("IdsCto_YN"))%></td>
                        <td><%=PrepForWeb(rs("RctoSkus_YN"))%></td>
                        <td><%=PrepForWeb(rs("RctoCto_YN"))%></td>
                        <% if (sProdVersionBSAMFlag = "True") Then %>
                        <td><%=PrepForWeb(rs("BSAMSkus_YN"))%></td>
                        <td><%=PrepForWeb(rs("BSAMBparts_YN"))%></td>
                        <% End If %>
                        <td><%=PrepForWeb(rs("UPC"))%></td>
                        <td><%=PrepForWeb(rs("Weight"))%></td>
                        <td><%=PrepForWeb(rs("GSEndDt"))%></td>
                    </tr>
                    <%
                iAvCount = iAvCount + 1        
		End If
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
Set cmd = dw.CreateCommAndSP(cn, "usp_SelectProgramMatrixSS")
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
        This is not the Publish Matrix. Working Data Snapshot taken at
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
                <th>AV<br />Level 2</th>
                <th>SA<br />Level 3</th>
                <th>Component<br />Level 4</th>
                <th>Component<br />Level 5</th>
                <th>Component<br />Level 6</th>
                <th>Qty</th>
                <th>SAP Rev</th>
                <th>ZWAR = X<br />PRI/ALT</th>
                <th>ROHS</th>
                <th>UPC</th>              
                <th>IDS-SKUS</th>
                <th>IDS-CTO</th>
                <th>RCTO-SKUS</th>
                <th>RCTO-CTO</th>
                <% If (sProdVersionBSAMFlag = "True") Then %>
                <th>BSAM-SKUS</th>
                <th>BSAM-Bparts</th>
                <% End If %>
                <th>Config Rules</th>
                <th>AVID</th>
                <th>Group 1</th>
                <th>Group 2</th>
                <th>Group 3</th>
                <th>Group 4</th>
                <th>Group 5</th>
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
Loop
rs.Close
rsVer.Close

Sub PmDrawCategoryRow(rs, iVersions)
    If (sProdVersionBSAMFlag = "True") Then
        Response.Write "<tr class=""MatrixFeatureCategory""><td colspan=" & 28 + iVersions & ">"
    else
        Response.Write "<tr class=""MatrixFeatureCategory""><td colspan=" & 26 + iVersions & ">"
    end if
    Response.Write rs("FeatureCategory") & ""
    Response.Write "</td></tr>"
End Sub

Sub PmDrawAvRow(rs, rsVer, iVersions)
    sLastAv = Trim(rs("AvNo") & "")
    
    rsVer.Filter= "SortAvNo = '" & rs("AvNo") & "'"

    Response.Write "<tr class=""MatrixAvRow""><td>"
    Response.Write rs("AVManufacturingNotes") & "" 
    Response.Write "</td>"
    Do Until rsVer.EOF
        Response.Write "<td class=""cell"">"
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
        Response.Write rs("bsamskus") & ""
        Response.Write "</td><td>"
        Response.Write rs("bsambparts") & ""
        Response.Write "</td><td>"
    End If
    Response.Write rs("ConfigRules") & ""
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
    Response.Write "</td><tdcolspan=""14""></td></tr>"
    
    PmDrawComponentRow rs, iVersions
End If
End Sub

Sub PmDrawComponentRow(rs, iVersions)
Dim i

If Trim(rs("CompNo") & "") <> "" Then
    sLastComponent = Trim(rs("CompNo") & "")
    Response.Write "<tr><td>&nbsp;</td>"
    For i = 0 to iVersions
        Response.Write "<td class=""cell"">&nbsp;</td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription") & ""
    Response.Write "</td><td>&nbsp;</td><td>&nbsp;</td><td nowrap>"
    Response.Write rs("CompNo") & ""
    Response.Write "</td><td>&nbsp;</td><td>&nbsp;</td><td>"
    Response.Write rs("CompQuantity") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS") & ""
    Response.Write "</td><tdcolspan=""14""></td></tr>"
    PmDrawComponentRow_L5 rs, iVersions
End If
End Sub
Sub PmDrawComponentRow_L5(rs, iVersions)
Dim i

If Trim(rs("CompNo_L5") & "") <> "" Then
    sLastComponent_L5 = Trim(rs("CompNo_L5") & "")
    Response.Write "<tr><td>&nbsp;</td>"
    For i = 0 to iVersions
        Response.Write "<td class=""cell"">&nbsp;</td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription_L5") & ""
    Response.Write "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td nowrap>"
    Response.Write rs("CompNo_L5") & ""
    Response.Write "</td><td>&nbsp;</td><td>"
    Response.Write rs("CompQuantity_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR_L5") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS_L5") & ""
    Response.Write "</td><tdcolspan=""14""></td></tr>"
    PmDrawComponentRow_L6 rs, iVersions
End If
End Sub

Sub PmDrawComponentRow_L6(rs, iVersions)
Dim i

If Trim(rs("CompNo_L6") & "") <> "" Then
    sLastComponent_L6 = Trim(rs("CompNo_L6") & "")
    Response.Write "<tr><td>&nbsp;</td>"
    For i = 0 to iVersions
        Response.Write "<td class=""cell"">&nbsp;</td>"
    Next
    Response.Write "<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write rs("CompGPGDescription_L6") & ""
    Response.Write "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td nowrap>"
    Response.Write rs("CompNo_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompQuantity_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompRevisionLevel_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompZWAR_L6") & ""
    Response.Write "</td><td>"
    Response.Write rs("CompROHS_L6") & ""
    Response.Write "</td><tdcolspan=""14""></td></tr>"
End If
End Sub
                %>
            </tbody>
        </table>
    </div>
    <%End If%>

    <!-- Popups -->
    <div id="PopUpMenu" class="hidden">
        <ul id="menu">
            <li class="default"><a href="#" onclick="parent.location.href='javascript:MenuProperties();'">Properties</a></li>
            <li><a href="#" onclick="parent.location.href='javascript:MenuEditMrDates();'">Edit
                MR Dates</a></li>
            <li id="spacer">
                <hr width="95%">
            </li>
            <li id="delete"><a href="#" onclick="parent.location.href='javascript:MenuDelete();'">Delete This AV</a></li>
            <li id="Li1">
                <hr width="95%" />
            </li>
            <li id="obsolete"><a href="#" onclick="parent.location.href='javascript:MenuSetObsolete();'">Obsolete This AV</a></li>
            <li id="activate"><a href="#" onclick="parent.location.href='javascript:MenuSetActive();'">Activate This AV</a></li>
            <li id="Li2">
                <hr width="95%" />
            </li>
            <li id="hide"><a href="#" onclick="parent.location.href='javascript:MenuSetHidden();'">Hide This AV (SCM/PM)</a> </li>
            <li id="unhide"><a href="#" onclick="parent.location.href='javascript:MenuSetActive();'">Unhide This AV (SCM/PM)</a> </li>
        </ul>
    </div>

    <div id="modal_programmatrix" class="content-dialog hide">
        <div class="menu_form_body">
            <form method="post" action="/iPulsar/ExcelExport/ProgramMatrix.aspx?ProductBrandID=<%=m_BrandID %>" id="popup_form">
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
                        <th style="white-space: nowrap">Comparison:
                        </th>
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
                    <%If (bIsPc And ShowReport <> "spb") Then %>
                    <tr>
                        <th>Publish:
                        </th>
                        <td>
                            <input type="checkbox" name="chkPublish" title="Publish" onclick="chkPublish_Click(<%= m_BrandID %>, this.checked);" />
                        </td>
                    </tr>
                    <tr>
                        <th>Push to x-ROST:</th>
                        <td>
                            <input type="checkbox" name="chkXrost" id="chkXrost" title="XRost" onclick="pushToXrost(<%= m_BrandID %>, this.checked);" />
                        </td>
                    </tr>
                    <% end if %>
                    <tr>
                        <th>New&nbsp;Matrix:
                        </th>
                        <td>
                            <input type="checkbox" id="chkNewMatrix" name="chkNewMatrix" title="Publish" />
                        </td>
                    </tr>
                    <tr>
                        <th>&nbsp;
                        </th>
                        <td>
                            <input class="btn" type="submit" id="popup_submit" value="Export" onclick="closeModalDialog();" />
                        </td>
                    </tr>
                </table>
            </form>
        </div>
    </div>

    <div style="display: none;">
        <div id="iframeDialog" title="Coolbeans">
            <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
        </div>
    </div>
    <input type="hidden" id="txtClass" name="txtClass" value="<%=sClass%>" />
    <input type="hidden" id="txtID" name="txtID" value="<%=PVID%>" />
    <input type="hidden" id="txtFavs" name="txtFavs" value="<%=sFavs%>" />
    <input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=sFavCount%>" />
    <input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>" />
    <input type="hidden" id="hidLastPublishDt" value="<%=m_LastPublishDt %>" />
    <input type="hidden" id="hidBusinessID" value="<%=m_BusinessID %>" />
    <input type="hidden" id="hidAvCount" value="<%= iAvCount %>" />
    <input type="hidden" id="hidSCMCategories" value="<%= sSCMCategories %>" />
    <input type="hidden" id="hidPMCategories" value="<%= sPMCategories %>" />
    <input type="hidden" id="hidProductBrandID" value="<%= m_BrandID %>" />
</body>
</html>
<script type="text/javascript">
    $('#loadingProgress').hide();
    $('#loadingProgress').spin('medium', '#0096D6');

    $(window).on('beforeunload', function () {
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

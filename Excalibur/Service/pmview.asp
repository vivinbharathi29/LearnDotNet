<%@  language="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%
Server.ScriptTimeout = 480
Response.Clear
   
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
    Dim sClass
    Dim sSeries : sSeries = regEx.Replace(Request("Series"), "")
    Dim sGroupBy : sGroupBy = regEx.Replace(Request("GroupBy"), "")
    Dim sStatus : sStatus = regEx.Replace(Request("Status"), "")
    Dim sInterval : sInterval = regEx.Replace(Request("Interval"), "")
    Dim sRegion : sRegion = regEx.Replace(Request("Region"), "")
    Dim LinkAnchor : LinkAnchor = regEx.Replace(Request("Anchor"),"")
    Dim strSKU: strSKU=""
    dim noSKAVresults 

    if(Request("Class") = "") then
        sClass = "1"
    else
        sClass = regEx.Replace(Request("Class"), "")
    end if
%>
<%
Dim rs, dw, cn, cmd, strSql
Dim iFeatureCategoryID, sFeatureCategory
Dim iAvCount
Dim bFirstWrite
Dim bIsPc : bIsPc = False
Dim m_BrandID : m_BrandID = ""
Dim m_BrandName : m_BrandName = ""
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
Dim strDevCenter
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
Dim ServiceReport
Dim sProgramName
Dim strCallDataLastUpdated
Dim sFusionRequirements : sFusionRequirements = false
Dim sBrandDisplayed : sBrandDisplayed = ""

Dim AppRoot
AppRoot = Session("ApplicationRoot")

on error resume next

ServiceReport = Trim(Request.Cookies("ServiceReport"))
regEx.Pattern = "^rsl|spb|sku$"

If Not regEx.Test(ServiceReport) Then 
    ServiceReport = "spb"
End If


If ServiceReport="sku" Then
    Dim objRegExp
    
    Set objRegExp= New RegExp
    objRegExp.Global = False
    objRegExp.Pattern = "[^0-9a-zA-Z_#]"

    strSKU =  UCase(objRegExp.Replace(TRIM(Session("SKU")),""))

    Set objRegExp=Nothing

    Session("SKU")=""
End If

on error goto 0
If instr(CurrentUser,"\") > 0 Then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If

Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim m_ProductVersionID : m_ProductVersionID = PVID
Dim m_IsSpdmUser : m_IsSpdmUser = False
Dim m_CurrentUserId
Dim m_IsTestLead
Dim m_IsToolPm
Dim m_IsActivePm
Dim m_IsCommodityPM
Dim m_IsSCFactoryEngineer
Dim m_IsAccessoryPM
Dim m_IsHardwarePm
Dim m_IsGplm
Dim m_IsBomAnalyst
Dim m_IsRplm

Dim blnCanEditProduct : blnCanEditProduct = false
Dim blnPopFilterCategory
Dim blnApplyCatFilter
Dim strSelCatIDs

Dim strCatNames
Dim strCatIDs
Dim strCurrCatID

Dim strTruncQS : strTruncQS=""
Dim strNoneQS : strNoneQS=""
Dim intQSIdx

IF Request.QueryString.Count>0 THEN
    IF(LEN(TRIM(Request.QueryString("CatIDs"))))>0 THEN
        FOR intQSIdx=1 TO Request.QueryString.Count
            If UCase(Trim(Request.QueryString.Key(intQSIdx)))<>"CATIDS" Then
                If LEN(strTruncQS)=0 Then
                    strTruncQS=Request.QueryString.Key(intQSIdx) & "=" & Request.QueryString.Item(intQSIdx)
                else
                    strTruncQS=strTruncQS & "&" & Request.QueryString.Key(intQSIdx) & "=" & Request.QueryString.Item(intQSIdx)
                End If
            End If
        NEXT
        IF LEN(TRIM(strTruncQS))>0 THEN
            strNoneQS=strTruncQS
            strTruncQS=strTruncQS & "&"
        END IF
    ELSE
        strNoneQS=Request.QueryString
        strTruncQS=strNoneQS & "&"
    END IF
END IF



'##############################################################################	
'
' Create Security Object to get User Info
'
	
	bIsPc = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
    	m_IsGplm = Security.UserInRole(PVID, "GPLM")
    	m_IsBomAnalyst = Security.UserInRole(PVID, "SBA")
    	m_IsRplm = Security.UserInRole("", "RPLM")
    
	m_UserFullName = Security.CurrentUserFullName()
    	m_CurrentUserId = Security.CurrentUserId()
            	
	If m_IsSysAdmin Or m_IsGplm Or m_IsBomAnalyst Or m_IsRplm Then
		m_IsSpdmUser = True
	End If
	


Dim CurrentPartnerTypeID 
Dim blnIsOSSPUser: blnIsOSSPUser=false

'==================================================================================================================

	Set Security = Nothing
'##############################################################################	


'
' Setup the data connections
'
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "spGetUserInfo")
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

	CurrentPartnerTypeID=trim(rs("PartnerTypeID"))

	If(CurrentPartnerTypeID = "2" OR CurrentPartnerTypeID = 2) Then
		blnIsOSSPUser=true	
	End If

    'permission needed for Edit Product link on Pulsar products
    if rs("CanEditProduct") = 1 then
        blnCanEditProduct = true
    end if

End If
rs.Close

IF (CurrentUserPartner<>1) AND (CurrentPartnerTypeID<>2) THEN
'************************************************************************************************
' Verify the requested Product Version
'************************************************************************************************
Dim strUserPVList
Dim objPVComm
Dim objPVRS

Dim isAllowPartners
isAllowPartners = false

rs.Open "SELECT v.ID FROM ProductVersion v JOIN PartnerODMProductWhitelist w on v.PartnerID=w.ProductPartnerId WHERE w.UserPartnerId = " + CurrentUserPartner + " and v.ID=" + CStr(PVID) + ";",cn,adOpenForwardOnly
do while not rs.EOF
    isAllowPartners = true
    exit do
loop
rs.Close

If Not isAllowPartners Then

    Set objPVComm=dw.CreateCommandSQL(cn,"SELECT dbo.GET_USER_PVLIST(" & Trim(CurrentUserID) & ",NULL,0)")

    Set objPVRS=dw.ExecuteCommandReturnRS(objPVComm)

    If(Not objPVRS.EOF) Then
	    strUserPVList=objPVRS(0)
	
	    If(Instr("," & strUserPVList & ",","," & CStr(PVID) & ",")<=0) OR Len(trim(strUserPVList))=0 Then
		    Response.Write("<b>USER DOES NOT HAVE ACCESS TO THIS PRODUCT.</b>")

		    objPVRS.Close
		    Set objPVRS=Nothing
		    Set objPVComm=Nothing
		    Response.End

	    Else
		    objPVRS.Close
		    Set objPVRS=Nothing
		    Set objPVComm=Nothing

	    End If
    Else
	    Set objPVRS=Nothing
	    Set objPVComm=Nothing

	    Response.Write("<b>USER DOES NOT HAVE ACCESS TO ANY PRODUCTS.</b>")
	    Response.End
    End If

End If '''Not isAllowPartners

'************************************************************************************************
END IF  '''(CurrentUserPartner<>1) AND (CurrentPartnerTypeID<>2)

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

Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, PVID
Set rs = dw.ExecuteCommandReturnRS(cmd)

If (rs.EOF And rs.BOF) And PVID <> "-1" Then
	Response.Write "Unable to find the selected program.<br /><font size=1>ID=" & PVID & "</font>"
	Response.Write "<br /><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & PVID & ")""><font face=verdana size=1>Remove From Favorites</font></a>"
	Response.Write "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
	Response.Write "<span id=EditLink style=""Display:none""></span><span id=StatusLink style=""Display:none""></span><span id=menubar style=""Display:none""></span><span ID=Wait style=""Display:none""></span>"
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""1"">"
Else
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""0"">"
	sProductName = rs("Name") & " " & rs("Version") 
	sDisplayedProductName = rs("Name") & " " & rs("Version")
	sProgramVersion = rs("Version") & ""
	strDevCenter = trim(rs("DevCenter") & "")
	strCallDataLastUpdated = trim(rs("CallDataLastUpdated") & "")
    sFusionRequirements = rs("FusionRequirements")

	SEPMID = rs("sepmid")
	PMID = rs("PMID")
    if rs("SMID") & "" <> "" then
        PMID = PMID & "_" & rs("SMID")
    end if
    PMID = "_" & PMID & "_"
	strSCMPath = rs("SCMPath") & ""
    strProdType = rs("TypeID") & ""
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
	    myString = Replace(myString, vbCrLf, "<br />")
	    myString = Replace(myString, vbCr , "<br />")
	    myString = Replace(myString, vbLf, "<br />")
	    myString = Replace(myString, Chr(10), "<br />")
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
Set cmd = dw.CreateCommandSP(cn, "usp_GetServiceFamilyPn")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, m_ProductVersionID
Set rs = dw.ExecuteCommandReturnRS(cmd)

Dim sServiceFamilyPn
If Not rs.EOF Then 
    sServiceFamilyPn = rs("ServiceFamilyPn") & ""
End If
rs.Close

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!--<meta http-equiv="X-UA-Compatible" content="IE=8" />-->

    <title>Service PM View</title>
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/service/style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/service/sample.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/cupertino/jquery-ui-1.8.2.custom.css" rel="stylesheet" type="text/css" />
        
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="includes/client/json2.js"></script>
    <script type="text/javascript" src="includes/client/json_parse.js"></script>
    <script src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/popup.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/spin.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/jquery.spin.js" type="text/javascript"></script>
    <script src="/pulsar2/js/userfavorite.js" type="text/javascript"></script>

    <script type="text/javascript">
        <!--
        // JQuery Block
        //
        $(function() {
            $("#PublishDialog").dialog({
                modal: true,
                autoOpen: false,
                width: 400,
                height: 200
            });

            $("#QuickSearch").dialog({
                modal: true,
                autoOpen: false,
                width: 300,
                height: 100,
                buttons: {
                    "Go": function() {
                        PartQuickSearch();
                    }
                }
            });

            $("#txtQuickSearch").keypress(function(e) {
                if (e.keyCode == 13) {
                    PartQuickSearch();
                }
            });

            $('#btnRefreshRSL').click(function () {
                ServiceReport('rsl');
            });
        });

        function PartQuickSearch() {
            var searchValue = $("#txtQuickSearch").val();
            window.open("<%= AppRoot %>/MobileSE/Today/find.asp?Find=" + encodeURIComponent(searchValue) + "&Type=Part");
            $("#QuickSearch").dialog("close");
        }
        // -->
    </script>

    <script type="text/javascript">
        var _userFullName = '<%= m_UserFullName %>';
        var _serviceReport = '<%= ServiceReport %>';
        
        String.prototype.trim = function() {
            return this.replace(/^\s+|\s+$/g, "");
        }

        String.prototype.ltrim = function() {
            return this.replace(/^\s+/, "");
        }

        String.prototype.rtrim = function() {
            return this.replace(/\s+$/, "");
        }


        //
        // BEGIN HEADER SUPPORT
        //

        function window_onload() {
            var anchor = document.getElementById("hidAnchor").value;
            var strFavorites = "," + document.getElementById("txtFavs").value;
            var strID = document.getElementById("txtID").value;
            var found = strFavorites.indexOf(",P" + strID.trim() + ",");
            var editLink = document.getElementById("EditLink");
            var cloneLink = document.getElementById("CloneLink");

            var expireDate = new Date();
            expireDate.setMonth(expireDate.getMonth() + 12);
            document.cookie = "LastProductDisplayed=" + strID + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";

            try{

            if (txtClass.value == "") {
                editLink.style.display = "none";
                cloneLink.style.display = "none";
                RFLink.style.display = "none";
                AFLink.style.display = "none";
                StatusLink.style.display = "";
            }
            else if (found == -1) {
                editLink.style.display = "";
                cloneLink.style.display = "";
                RFLink.style.display = "none";
                AFLink.style.display = "";
            }
            else {
                editLink.style.display = "";
                cloneLink.style.display = "";
                RFLink.style.display = "";
                AFLink.style.display = "none";
            }

            editLink.style.display = "";
            cloneLink.style.display = "";

            }catch(err){}

            var lblAvCount = document.getElementById("lblAvCount");
            if (lblAvCount != null) {
                lblAvCount.innerHTML = "( " + hidAvCount.value + " AVs Displayed )";
            }
                        
            if (_serviceReport == "rsl" || _serviceReport == "spb" || _serviceReport == "sku")
                initializeDIVStyleSize();
            
            if(anchor.trim().length>0)
            {
                document.location.hash = "#" + anchor;
                var oDIV=document.getElementById("DIV1");

                if((oDIV!=undefined)&&(oDIV!=null))
                    oDIV.doScroll("scrollbarUp");
            }

            //Instantiate modalDialog load - PBI 26992 (Showmodaldialog to jquery dialog)
            modalDialog.load();
        }

        // Close product properties dialog when opened from Service tab - PBI 26992 (Showmodaldialog to jquery dialog)
        function ClosePropertiesDialog(strID) {
            modalDialog.cancel(false);

            if (typeof (strID) != "undefined") 
            {
                window.location.reload(true);
            }
        }

        // Close clone product dialog when opened from Service tab - PBI 26992 (Showmodaldialog to jquery dialog)
        function ClosePropertiesDialog_fromClone(strID) {
            modalDialog.cancel(false);

            if (typeof (strID) != "undefined") {
                if (strID == txtID.value) {
                    window.location.reload(true);
                } else {
                    window.location = "/Excalibur/pmview.asp?ID=" + strID + "&Class=" + txtClass.value;
                }
            }
        }

        function initializeDIVStyleSize(){
            
            if(document.getElementById("DIV1")!=undefined){
                if (window.DIV1.style.width == "")
                {
                    window.DIV1.style.width = window.screen.availWidth - window.DIV1.style.left - 50;	
                    window.DIV1.style.height = window.screen.availHeight - window.DIV1.style.top - 320;
                } else {
                    
                    if(!bCategoryFilter){
                        window.DIV1.style.width = window.frameElement.width - window.DIV1.style.left - 50;	
                        window.DIV1.style.height = window.frameElement.height - window.DIV1.style.top - 275;
                        
                    } else {
                        window.DIV1.style.width = window.frameElement.width - window.DIV1.style.left - 50;	
                        window.DIV1.style.height = window.frameElement.height - window.DIV1.style.top - 355;
                    }
                }
            }
            return;
        }

        //
        //  Added to implement CATEGORY FILTERS for RSL, SPB, and SKU views
        //
	    function populateFilterList(sIDs, sNames, sSelCatIDs)
	    {
		    var aIDs=sIDs.split('|');
		    var aNames=sNames.split('|');
		    var i;
		    var oList=document.getElementById("CategoryList");
            var sTmp="";

		    if(oList!=undefined)
		    {
                oList.options.length=0;

                if(aIDs.length==aNames.length) // Double check, should not be necessary
                {
			        for(i=0;i<aIDs.length;i++)
			        {
                        if(aNames[i].toString().trim().length>0) // Only add items that are not blank
                        {
				            oList.add(new Option(aNames[i], aIDs[i]));

                           if(sSelCatIDs!=null)
                           {
                                // Determine if the option added is selected, may be an indexOf or similar method in <select> element
                                sTmp="|"+aIDs[i].toString()+"|";

                                if(sSelCatIDs.indexOf(sTmp)>-1)
                                {
                                    oList.options[oList.options.length-1].selected=true;
                                }
                           }
                       }
			        }
                }
		    }
            
	    }


        function getListSelections(oList)
        {
            var sRetValue="";
            var i=0;

            if(oList.options.length>0)
            {
                for(i=0;i<oList.length;i++)
                {
                    if(oList.options[i].selected)
                    {
                        if(sRetValue.length==0)
                        {
                            sRetValue=oList.options[i].value.toString();
                        }else{
                            sRetValue+="|"+oList.options[i].value.toString();
                        }
                    }
                }
            }

            return sRetValue;
        }


        function applyCatFilter()
        {
            var i=0;
            var oList=document.getElementById("CategoryList");
            var sSelectedCatIDs=getListSelections(oList);
            
            // Reload the page with the Category Filter applied
            if(sSelectedCatIDs.length>0)
            {
                window.location="<%=AppRoot %>/pmview.asp?<%=strTruncQS%>CatIDs="+sSelectedCatIDs;
            }
            else if(oList.options.length>0)
            {
                alert("Please select 1 or more Categories with which to Filter data.");
                oList.focus();
            }
            else
            {
                alert("No Categories exist with which to Filter data.");
            }
        }

        function applyAVFilter()
        {       
            var avNo = document.getElementById("avtext").value;
            // get rid of querystring if 2nd time
            var urlstr = "<%=AppRoot %>/pmview.asp?<%=strTruncQS%>";
            urlstr = urlstr.split("av=")[0];
              
            if(avNo.length>0)
            {
                window.location = urlstr + "av=" + avNo;
            }
            else
            {
                alert("Please enter a AV No.");
            }
        }


        function applySKFilter()
        {
            var skNo = document.getElementById("sktext").value;

            // get rid of querystring if 2nd time
            var urlstr = "<%=AppRoot %>/pmview.asp?<%=strTruncQS%>";
            urlstr = urlstr.split("sk=")[0];
              
            if(skNo.length>0)
            {
                window.location = urlstr + "sk=" + skNo;
            }
            else
            {
                alert("Please enter a Spare Kit No.");
            }
        }

         function toggleAVFilter(oRdoBut)
         {
            // av
            // alert("ok");
            var ronAVFilterActions=document.getElementById("AVfilter");
            var ronFilterActions=document.getElementById("sparekitfilter");
            switch(oRdoBut.value)
            {
                case "0": // None
                    ronAVFilterActions.style.display = 'none';           
                    break;
               case "1": // Other
                    ronAVFilterActions.style.display = '';
                    // ronFilterActions.style.display = 'none';
                    break;
            }
            if((ronAVFilterActions)&&(oRdoBut.value=="0"))
            {
                // Reload the page without a Filter
                var urlstr = "<%=AppRoot %>/pmview.asp?<%=strNoneQS%>";
                urlstr = urlstr.split("&av=")[0];
                window.location = urlstr;
            }
            return;
        }

        function toggleSKAVFilter(oRdoBut)
        {
            //sparekit        
            var ronFilterActions=document.getElementById("sparekitfilter");
            var ronAVFilterActions=document.getElementById("AVfilter");
            switch(oRdoBut.value)
            {
                case "0": // None
                    ronFilterActions.style.display = 'none';           
                    break;
               case "1": // Other
                    ronFilterActions.style.display = '';
                    //ronAVFilterActions.style.display = 'none';            
                    break;
            }
            if((ronFilterActions)&&(oRdoBut.value=="0"))
            {
                // Reload the page without a Filter
                var urlstr = "<%=AppRoot %>/pmview.asp?<%=strNoneQS%>";
                urlstr = urlstr.split("&sk=")[0];
                window.location = urlstr;
            }
            return;
                  
        }
        function toggleCategoryFilter(oRdoBut)
        {
            var oFilterActions=document.getElementById("FilterActions");
            var oList=document.getElementById("CategoryList");
            
            switch(oRdoBut.value)
            {
              case "0": // None
                    oFilterActions.style.display = 'none';
                    break;

              case "1": // Other
                    if(oList.options.length>0)
                    {
                        oFilterActions.style.display = '';
                    }else{
                        alert("No Categories exist with which to Filter data.");
                        
                        // Reset action by user         
                        oRdoBut.checked=false;

                        var oNoneRdoBut=document.getElementById("FilterOptionsNone");
                        oNoneRdoBut.checked=true;

                        return;
                    }
                    break;
            }

            if((bCategoryFilter)&&(oRdoBut.value=="0"))
            {
                // Reload the page without a Category Filter
                window.location="<%=AppRoot %>/pmview.asp?<%=strNoneQS%>";
            }
            return;
        }

        //
        //  Batch Updating Support
        //
        function getSelectedRows(sName)
        {
            var aElements=document.getElementsByName(sName);
            var sSelRows="";
            var i;

            if(aElements.length>0)
            {
                for(i=0;i<aElements.length;i++)
                {
                    if(aElements[i].checked)
                    {
                        if(sSelRows.length==0)
                        {
                            sSelRows=aElements[i].getAttribute("id");
                        }else{
                            sSelRows+=","+aElements[i].getAttribute("id");
                        }
                    }
                }
            }

            return sSelRows;

        }


        function doBatchUpdate()
        {
            var sSelections=getSelectedRows("CBCatDtl");
            var aSelections=null;
            var strRetVal;
            var i;
            var sNodeAttrs="";
            var oCurrChkBox=null;
            var oCurrAttrNode=null;
            var sPVID=null;
            var sSFPN=null;
            var CategoryId=null;
            var sSKIDs="";

            if(sSelections.length>0)
            {
                aSelections=sSelections.split(",");

                // MAY NEED TO ADD VALIDATION WHEN CHECKBOX IS CLICKED TO CONFIRM THAT THE MINIMUM INFORMATION TO UPDATE RECORD IS PRESENT
                // Build a list (MAY EVENTUALLY USE AN "OBJECT" ARRAY) of ProductVersionID, DeliverableRootId, ServiceFamilyPn, SpareKitId, CategoryId
                for(i=0; i<aSelections.length; i++)
                {
                    oCurrChkBox=document.getElementById(aSelections[i]);

                    oCurrAttrNode=getRowAttributeNode(oCurrChkBox);

                    if((oCurrAttrNode!=undefined)&&(oCurrAttrNode!=null)&&(oCurrAttrNode!=oCurrChkBox))
                    {
                        sNodeAttrs="";

                        // ELIMINATE REDUNDANCIES (e.g. - PVID etal...only send the bare minimum)...COMMON IDENTIFIERS CAN BE PART OF QUERYSTRING

                        if(sPVID==null)
                        {
                            if (oCurrAttrNode.PVID != undefined) sPVID=oCurrAttrNode.PVID;
                        }

                        if(sSFPN==null)
                        {
                            if (oCurrAttrNode.SFPN != undefined) sSFPN=oCurrAttrNode.SFPN;
                        }

                        if(CategoryId==null)
                        {
                            if(oCurrAttrNode.CID != undefined) CategoryId=oCurrAttrNode.CID;
                        }

                        if (oCurrAttrNode.SKID != undefined)
                            sNodeAttrs +=  oCurrAttrNode.SKID;

                        if(sSKIDs.length==0)
                        {
                            sSKIDs=sNodeAttrs;
                        }else{
                            sSKIDs+=","+sNodeAttrs;
                        }

                    }
                    else
                    {
                        alert("Problem!");
                        break;
                    }
                }

                // create a pop-up dialog window 
                // CSR Level, 
                // Disposition, 
                // Warranty Labour Tier, 
                // Local Stock Advice, 
                // GEOS (4 checkboxes), 
                // First Service Dt., 
		        // Notes, and 
                // RSL Comments.
               
                strRetVal=window.showModalDialog("<%=AppRoot%>/Service/SpareKitDetailsBatch.asp?PVID="+sPVID+"&SFPN="+sSFPN+"&SKIDS="+sSKIDs, "", "dialogWidth:635px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");

                if((strRetVal=="cancel")||(strRetVal=="undefined")) return;

    	        var sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>";

                // Create Current Category location Anchor
	            if(sLocation.indexOf("&Anchor")>-1) {
		            sLocation = sLocation.split("&Anchor")[0] + "&Anchor=C" + CategoryId;
                } else {
		            sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>&Anchor=C"+CategoryId;
	            }
		
                // Retain Category Filter, if applicable
                if(bCategoryFilter)
                {
                    sLocation+="&CatIDs="+getListSelections(document.getElementById("CategoryList"));
                }

                window.location = sLocation;

            } else {
                alert("Please select 1 or more rows to update.");
            }
        }


        function getRowAttributeNode(oChkBox) 
	    {
            var node = oChkBox;

            //Find the row element to get the properties.
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

	        return node;
        }


        function toggleCBCatAll(oChkBox)
        {
            var sHdrIDPref="CBCatHdr";
            var sDtlIDPref="CBCatDtl";
            var iIdx=0;
            var oCurrDtlElement=null;
            var aCatHeaders=document.getElementsByName("CBCatHdr");
            var sHdrID=null;
            var sDtlID=null;

            if(aCatHeaders.length>0)
            {
                for(iIdx=0;iIdx<aCatHeaders.length;iIdx++)
                {

                    sHdrID=aCatHeaders[iIdx].getAttribute("id");

                    if(sHdrID!=undefined)
                    {
                        // Determine if any applicable Detail Rows exist
                        oCurrDtlElement=null;
                        sDtlID=sDtlIDPref+sHdrID.replace(sHdrIDPref,"")+"-0";
                        oCurrDtlElement=document.getElementById(sDtlID);

                        if((oCurrDtlElement!=undefined)&&(oCurrDtlElement!=null))
                        {
                            aCatHeaders[iIdx].checked=oChkBox.checked;
                            toggleCBCatHdr(aCatHeaders[iIdx]);
                        }
                    }
                }

            }
            else
            {
                alert("No Categories Exist!");
            }

        }


        function toggleCBCatHdr(oChkBox)
        {
            // Cancel event handling by 'rslCategoryRow_onclick'
            window.event.cancelBubble=true;

            var sHdrID=oChkBox.getAttribute("id");
            var sHdrIDPref="CBCatHdr";
            var sDtlIDPref="CBCatDtl";
            var iIdx=0;
            var oCurrDtlElement=null;
            var oAllElement=document.getElementById("CBCatAll");

            try
            {
                if(sHdrID!=undefined)
                {
                    sDtlIDPref+=sHdrID.replace(sHdrIDPref,"")+"-";

                    // Check|Uncheck all related CBCatDtl check boxes
                    oCurrDtlElement=document.getElementById(sDtlIDPref+iIdx.toString());

                    if((oCurrDtlElement==undefined)||(oCurrDtlElement==null))
                    {
                        // Not Applicable, i.e. - no detail records exist that can be flagged (probably due to lack of 'Part No.' OR 'Description' and therefore 'SKID')
                        // Force unchecking of this box
                        alert('CAN NOT FLAG THIS CATEGORY FOR BATCH UPDATES BECAUSE EITHER: \n(1) No Detail records exist at all\nOR\n(2) No Detail records exist that can be flagged for a Batch Update (perhaps due to the lack of a "Part Number" OR "Description")');
                        oChkBox.checked=false;
                        //oAllElement.checked=false;
                        
                    }else{

                        if(!oChkBox.checked) oAllElement.checked=false;

                        while((oCurrDtlElement!=undefined)&&(oCurrDtlElement!=null))
                        {
                            oCurrDtlElement.checked=oChkBox.checked;
                            iIdx++;
                            oCurrDtlElement=document.getElementById(sDtlIDPref+iIdx.toString());
                        }
                    }

                }
            }catch(Error)
            {
                alert(Error.Message);
            }

        }


        function toggleCBCatDtl(oChkBox)
        {
            // Cancel event handling by 'rslCategoryRow_onclick'
            window.event.cancelBubble=true;

            var sDtlID=oChkBox.getAttribute("id");
            var sDtlIDPref="CBCatDtl";
            var sHdrIDPref="CBCatHdr";
            var aTmp;
            var oHdrElement=null;
            var oAllElement=document.getElementById("CBCatAll");

            try
            {
                if((sDtlID!=undefined)&&(!oChkBox.checked))
                {
                    aTmp=sDtlID.split("-");

                    if(aTmp.length>1)
                    {
                        sHdrID=sHdrIDPref+aTmp[0].replace(sDtlIDPref,"");

                        // Uncheck related CBCatHdr checkbox
                        oHdrElement=document.getElementById(sHdrID);

                        if(oHdrElement.checked) 
                        {
                            oHdrElement.checked=false;
                        }

                        // Uncheck the CBCatAll checkbox
                        if(oAllElement.checked)
                        {
                            oAllElement.checked=false;
                        }
                    }
                }
            }catch(Error)
            {
                alert(Error.Message);
            }

        }


        function getCBCatAllTitle(oChkBox)
        {
            try
            {
                if(oChkBox.checked)
                {
                    oChkBox.setAttribute("title","UnSelect All applicable Categories for Batch Update");
                }else{
                    oChkBox.setAttribute("title","Select All applicable Categories for Batch Update");
                }
            }catch(Error)
            {
            }
        }


        function getCBCatHdrTitle(oChkBox, sCategoryName)
        {
            try
            {
                if(oChkBox.checked)
                {
                    oChkBox.setAttribute("title","UnSelect All [" + sCategoryName + "] for Batch Update");
                }else{
                    oChkBox.setAttribute("title","Select All [" + sCategoryName + "] for Batch Update");
                }
            }catch(Error)
            {
            }
        }


        function getCBCatDtlTitle(oChkBox)
        {
            try
            {
                if(oChkBox.checked)
                {
                    oChkBox.setAttribute("title","UnSelect for Batch Update");
                }else{
                    oChkBox.setAttribute("title","Select for Batch Update");
                }
            }catch(Error)
            {
            }
        }

        function QuickSearch(sItem) {
            $("#txtQuickSearch").val("");
            $("#QuickSearch").dialog("open");
        }

        function AdvancedSearch() 
	    {
            window.parent.open("<%=AppRoot%>/Service/ServiceAdvancedSearchReports.aspx","","toolbar=no,menubar=no,location=no,status=no,scrollbars=yes,resizable=yes","");

        }


        function deleteSkuAV(one,two,three,skno,mapid)
        {
         var win2 = window.parent.open("<%=AppRoot%>/Service/DeleteSKAV.aspx?brandID=" + one + "&sfPN=" + two + "&userID=" + three + "&four=" + skno,"","toolbar=no,menubar=no,location=no,status=no,scrollbars=yes,resizable=yes","");
        }
        function atest()
        {
        alert("test");
        }
      
        // **********************************************************************************************************************************************************************
        //
        //      SKU BOMs Client side functions
        //
        // **********************************************************************************************************************************************************************

        var sCurrentSKU = null;
        var oCurrentSKULink = null;
        var bProcessingSKURequest=false;

        function showSKUSpareKitsCallBack(sData) {

            var oSKUSKElem = document.getElementById("["+sCurrentSKU+"]");

            oSKUSKElem.innerHTML = sData;

            oCurrentSKULink.innerHTML = "Hide Spare Kits"; //"Hide";
            oCurrentSKULink.setAttribute("title", "Click to hide the Spare Kits for this SKU.");
	        oCurrentSKULink.setAttribute("loaded", "true");
	        oCurrentSKULink.setAttribute('mode', '1');

            bProcessingSKURequest = false;
        }


        function showSKUSpareKits(oElem) 
        {
            var sMode = oElem.getAttribute("mode");
            var sSKU = oElem.getAttribute("id");
	        var oSKURowElem = document.getElementById(sSKU+"Row");

	        if (sMode == "0") {
	            // Get the Spare Kits for the specified SKU

    	        if(oSKURowElem != null) {

	                oSKURowElem.style.Color = "red";

			if (!bProcessingSKURequest) oElem.style.Color = "red";
            	}

		        var sLoaded=oElem.getAttribute("loaded");

		        if(sLoaded=="true")
		        {
			        var oSKUSKElem=document.getElementById("["+sSKU+"]");

			        oSKUSKElem.style.display="";
            		oElem.innerHTML = "Hide Spare Kits"; //"Hide";
            		oElem.setAttribute("title", "Click to hide the Spare Kits for this SKU.");

            		oElem.setAttribute('mode', '1');
	    
		        }else{


           		    if (bProcessingSKURequest) {
           		        alert("Currently processing another request.\nTry again later.");
           		        return; // Exit as we can only process one request at a time for now
           		    }

			        bProcessingSKURequest = true;

	                oElem.innerHTML = "Retrieving Spare Kits...";
        	        oElem.setAttribute("title", "Retrieving Spare Kits...");

			        oCurrentSKULink = oElem;
	                sCurrentSKU = sSKU;

        	        jsrsExecute("<%=AppRoot %>/Service/rsBTOServiceSpare.asp", showSKUSpareKitsCallBack, "GetSKUSpareKits", Array(String(sCurrentSKU)));

		        }

            } else if (sMode == "1") {
                // Hide
                var oSKUSKElem=document.getElementById("["+sSKU+"]");
                oSKUSKElem.style.display = "none"; 

                oElem.setAttribute('mode', '0');
                oElem.innerHTML = "Show Spare Kits"; 
                oElem.setAttribute("title", "Click to display the Spare Kits for this SKU.");
            }
        }


        function gotoSKUHeader(oList) {
            var sSKU = oList.value.toString();
	    var iIdx=oList.selectedIndex+1;
	    var iLen=oList.options.length;
	    var oSKURowElem = document.getElementById(sSKU+"Row");

            if (oSKURowElem != null) {
                oSKURowElem.style.Color = "red";
            }

            document.location.hash = "#" + sSKU; // OUGHT TO HIGHLIGHT NAVIGATION

            var oDIV = document.getElementById("DIV1");

            if ((oDIV != undefined) && (oDIV != null)) {
		if(iIdx+2<iLen)
		{
                	oDIV.doScroll("scrollbarUp");
		}else{
			oDIV.doScroll("scrollbarDown");
		}
            }
        }

        // **********************************************************************************************************************************************************************

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

        function ShowProperties(DisplayedID, isCloning, FusionRequirements) {
            var strID;
            var shouldClone;

            if (isCloning == 1) {
                shouldClone = "&Clone=1";
            } else {
                shouldClone = "";
            }
            // Open Product Properties and Clone Product using jquery dialog when opened from Service tab - PBI 26992 (ShowModaldialog to jquery dialog)
            var url = "<%=AppRoot %>/mobilese/today/programs.asp?Commodity=0&ID=" + DisplayedID + shouldClone + "&Pulsar=" + FusionRequirements;
            modalDialog.open({dialogTitle:'Product Properties', dialogURL:''+url+'', dialogHeight:650, dialogWidth:1050, dialogResizable:true, dialogDraggable:true}); 
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

        function SelectTab(strStep, blnLoad) {
            var i;
            var expireDate = new Date();

            expireDate.setMonth(expireDate.getMonth() + 12);
            document.cookie = "PMTab=" + strStep + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";

            CurrentState = strStep;

            window.location.replace("<%=AppRoot %>/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>" + "&List=" + strStep);
        }

        function SetDefaultDisplay(strList, CurrentUserID) {

            if (window.confirm("Are you sure you want to make this display list the default display that you see each time you view a product?")) {
                //jsrsExecute("<%=AppRoot %>/DefaultProductTabRSUpdate.asp", myTabSetCallback, "UpdateTab", Array(strList, String(CurrentUserID)));
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

        //
        // END HEADER SUPPORT
        //

        function BrandLink_onClick(ProductBrandID) {
	        clearContent();
            window.location.replace("<%=AppRoot %>/Service/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&BID=" + ProductBrandID);
        }

        function SKUBOMBrandLink_onClick(ProductBrandID) {
            var oSPAN = document.getElementById("BTOSKUListSpan");
            if (oSPAN != null) oSPAN.innerHTML = "";
            var oDIV = document.getElementById("DIV1");
            oDIV.innerHTML = "<b>Processing request, please wait...</b>";                        
            window.location.replace("<%=AppRoot %>/pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&BID=" + ProductBrandID);

        }

        var oldColor;
        function HeaderMouseOver() {
            window.event.srcElement.style.cursor = "hand";
            oldColor = window.event.srcElement.style.color;
            window.event.srcElement.style.color = "red";
        }

        function HeaderMouseOut() {
            window.event.srcElement.style.color = oldColor;
        }

<% If (Not blnIsOSSPUser) Then %>

        function EcrProperties(strID, strType) {
            var strResult;
            strResult = window.showModalDialog("<%= AppRoot %>/service/action.asp?ID=" + strID + "&Type=" + strType, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;maximize:Yes;;center:Yes; help: No;resizable: Yes;status: No")
            if (typeof (strID) != "undefined") {
                window.location.reload(true);
            }
        }

<% End If %>

        function ServiceReportCallBack(sRetValue) 
        {
            window.location="<%=AppRoot %>/pmview.asp?<%=strNoneQS%>";
        }

        function ServiceReport(value)
        {
            var expireDate = new Date();
    	    expireDate.setTime (expireDate.getTime() + (365 * 24 * 60 * 60 * 1000)); // one year from now

	        clearContent();

            document.cookie = "ServiceReport=" + value + ";expires=" + expireDate.toGMTString();

	        window.location = "<%=AppRoot %>/service/pmview.asp?<%=strNoneQS%>";
	        location.reload();
        }

	function clearContent()
	{
	    var oDIVStatus=document.getElementById("SKUStatus");
	    var oDIV1=document.getElementById("DIV1");
	    var oSectionTitle=document.getElementById("theSectionTitle");
	    var oSectionList=document.getElementById("BTOSKUListSpan");
	    var oSSTS=document.getElementById("SnapShotTS");

	    if((oSectionList!=null)&&(oSectionList!=undefined)) 
		oSectionList.innerHTML="";
	
	    if((oSectionTitle!=null)&&(oSectionTitle!=undefined))
		oSectionTitle.innerHTML="";

	    if((oSSTS!=null)&&(oSSTS!=undefined))
		oSSTS.innerHTML="";
 		
	    if((oDIVStatus!=null)&&(oDIVStatus!=undefined))
		oDIVStatus.innerHTML="";

	    if((oDIV1!=null)&&(oDIV1!=undefined))
	    	oDIV1.innerHTML="<b>Processing request, please wait...</b>";
	}

        //
        // Scripts to support SPB
        //
<% If (Not blnIsOSSPUser) Then %>

        function SetServiceFamilyPn(PVID) {
            var strID = ""
            
            // Edit Service Family Details dialog - PBI 26992 (Showmodaldialog to jquery dialog)
            var url = "<%=AppRoot %>/service/spbSetServiceFamilyPn_Frame.asp?PVID=" + PVID;
            modalDialog.open({dialogTitle:'Edit Service Family Details', dialogURL:''+url+'', dialogHeight:600, dialogWidth:800, dialogResizable:true, dialogDraggable:true}); 
           
            // strID = window.parent.showModalDialog("<%=AppRoot %>/service/spbSetServiceFamilyPn_Frame.asp?PVID=" + PVID, "", "dialogWidth:800px;dialogHeight:600px;scroll:yes;edge:sunken;center:yes;help:no;resizable:yes;status:no");

            // Determine if the user updated the data OR if this is the SKU view (do not refresh if no updates were made or ServiceReport="sku" or "spb")
            //if ((strID == "false") || (_serviceReport == "spb") || (_serviceReport == "sku")) return;

            //window.location.reload();
        }

        // Edit Service Family Details return function - PBI 26992 (Showmodaldialog to jquery dialog)
        function SetServiceFamilyPn_return(strID) {

                if (strID == "false" || (_serviceReport == "spb") || (_serviceReport == "sku"))
                {
                    return;
                }
                else
                {
                    window.location.reload();
                }
        }


        function exportSpb(ServiceFamilyPn) {
            document.getElementById("trCompare").style.display = "";
            document.getElementById("selCompareDtDiv").innerHTML = "<b>Loading ...</b>";
            document.getElementById("popup_submit").disabled = true;
            document.getElementById("chkNewMatrix").checked = false;
            document.getElementById("chkPublish").checked = false;

            var popup_form = document.getElementById("popup_form");
            popup_form.action = "/iPulsar/ExcelExport/SparesBom.aspx";

            popup_show('popup', 'popup_drag', 'popup_exit', 'screen-center', 10, 10);

            jsrsExecute("<%=AppRoot %>/service/rsService.asp", getSpbPublishDatesCallback, "GetSpbPublishDates", String(ServiceFamilyPn));
        }


        function getSpbPublishDatesCallback(returnString) {
            var chkNewMatrix = document.getElementById("chkNewMatrix");
            var trCompare = document.getElementById("trCompare");
            document.getElementById("selCompareDtDiv").innerHTML = returnString;
            document.getElementById("popup_submit").disabled = false;

            if (returnString == "") {
                trCompare.style.display = "none";
                chkNewMatrix.checked = true;
            }

        }

<% End If %>

        var defaultStyleColor = null;

        function spbRow_onMouseOver() {
            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }
            defaultStyleColor = node.style.backgroundColor;
            node.style.backgroundColor = "#88FF88";
            node.style.cursor = "hand";
            window.status = "PVID:" + node.PVID + " SFPN:" + node.SFPN + " HPPN:" + node.HPPN;
        }

        function spbRow_onMouseOut() {
            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            node.style.backgroundColor = defaultStyleColor;
        }


        function spbPart_onClick() {
            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            ShowSpbPartDetails(node.PVID, node.HPPN, node.SFPN);
        }


        function ShowSpbPartDetails(ProductVersionID, HpPartNo, ServiceFamilyPn)
        {
            var strID;
            strID = window.parent.showModalDialog("<%=AppRoot %>/service/spbEditDetails_frame.asp?PVID=" + ProductVersionID + "&HPPN=" + HpPartNo + "&SFPN=" + ServiceFamilyPn, "", "dialogWidth:350px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
	        // Current data is a static snapshot, reloading will not do anything
            //document.location.reload();
        }

<% If (Not blnIsOSSPUser) Then %>
        function AddEcr(ProductVersionID)
        {
            var strID;
            strID = window.showModalDialog("<%=AppRoot %>/service/action.asp?ProdID=" + ProductVersionID + "&Type=7", "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
            if (typeof (strID) != "undefined") {
                window.location.reload(true);
            }
        }

        //
        // Change Request Tab
        //
        function ShowEcrStatus(value)
        {
            var expireDate = new Date();

            expireDate.setMonth(expireDate.getMonth() + 12);
            document.cookie = "EcrStatus=" + value + ";expires=" + expireDate.toGMTString() + ";";

            window.location.reload(true);
        }


<% End If%>

        function changerows_onmouseover()
        {
            if (window.event.srcElement.className == "cell") {
                window.event.srcElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.style.cursor = "hand";
            }
            else if (window.event.srcElement.className == "text") {
                window.event.srcElement.parentElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
            }
        }

        function changerows_onmouseout()
        {
            if (window.event.srcElement.className == "text")
                window.event.srcElement.parentElement.parentElement.style.color = "black";
            else if (window.event.srcElement.className == "cell")
                window.event.srcElement.parentElement.style.color = "black";
        }

<% If Not blnIsOSSPUser Then %>
        function ActionProperties(strID, strType)
        {
            var strResult;
            strResult = window.showModalDialog("<%=AppRoot %>/Service/action.asp?ID=" + strID + "&Type=" + strType, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;maximize:Yes;;center:Yes; help: No;resizable: Yes;status: No")
            if (typeof (strID) != "undefined") {
                window.location.reload(true);
            }
        }

        function ActionPrint(strID, strType)
        {
            var strResult;
            var NewTop;
            var NewLeft;

            NewLeft = (screen.width - 655) / 2
            NewTop = (screen.height - 650) / 2

            strResult = window.open("<%=AppRoot %>/Service/actionReport.asp?Action=0&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=yes,status=no")
        }

        function ActionMail(strID, strType) {
            var strResult;
            var NewTop;
            var NewLeft;

            NewLeft = (screen.width - 655) / 2
            NewTop = (screen.height - 650) / 2

            strResult = window.open("<%=AppRoot %>/mobilese/today/actionReport.asp?Action=1&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No")
        }

<% End If%>

        function changerows_onclick() {
            var strID;
            var strDisplay;


            if (window.event.srcElement.className == "text") {
                strDisplay = window.event.srcElement.parentElement.parentElement.className;
                window.event.srcElement.parentElement.parentElement.style.color = "black";
                strID = window.showModalDialog("<%=AppRoot %>/mobilese/today/action.asp?" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No")
            }
            else if (window.event.srcElement.className == "cell") {
                strDisplay = window.event.srcElement.parentElement.className;
                window.event.srcElement.parentElement.style.color = "black";
                strID = window.showModalDialog("<%=AppRoot %>/mobilese/today/action.asp?" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; maximize:Yes;help: No;resizable: Yes;status: No")
            }
            if (typeof (strID) != "undefined") {
                window.location.reload(true);
            }
        }

        //
        // Script for RSL
        //
        function rslRow_onmouseover() {
            var node = window.event.srcElement;
            var status = "";
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }
            defaultStyleColor = node.style.backgroundColor;
            node.style.backgroundColor = "#88FF88";
            node.style.cursor = "hand";
            if (node.PVID != undefined)
                status = status + " PVID:" + node.PVID;
            if (node.SFPN != undefined)
                status = status + " SFPN:" + node.SFPN;
            if (node.CID != undefined)
                status = status + " CID:" + node.CID;
            if (node.SKID != undefined)
                status = status + " SKID:" + node.SKID;
            if (node.DRID != undefined)
                status = status + " DRID:" + node.DRID;
            if (node.MAPID != undefined)
                status = status + " MAPID:" + node.MAPID;
            if (node.PBID != undefined)
                status = status + " PBID:" + node.PBID;
            window.status = status;
        }

        function rslRow_onmouseout() {
            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            node.style.backgroundColor = defaultStyleColor;
        }





        function rslRow_onclick() {
            var node = window.event.srcElement;

            //If we click on a hyperlink then don't open the properties window.
            if (node.nodeName.toLowerCase() == "a") return;
            
            //Find the row element to get the properties.
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            ShowRslDetails(node.PVID, node.DRID, node.SFPN, node.SKID, node.CID);
        }

        function mapRow_onclick() {
            var node = window.event.srcElement;

            //If we click on a hyperlink then don't open the properties window.
            if (node.nodeName.toLowerCase() == "a") return;

            //Find the row element to get the properties.
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            ShowMapDetails(node.PVID, node.SKID, node.MAPID, node.PBID, node.CID);
        }



        function ShowRslDetails(ProductVersionID, DeliverableRootId, ServiceFamilyPn, SpareKitId, CategoryId) {
            var strID;

            var boolServiceFamilyPnReady;

            <%
                rs.Open "SELECT TOP 1 sf.ServiceFamilyPn FROM ServiceFamilyDetails sf WHERE sf.ServiceFamilyPn IN (SELECT pv.ServiceFamilyPn FROM ProductVersion pv WHERE pv.ID = " + CStr(PVID) + ");",cn,adOpenForwardOnly
                if not(rs.eof and rs.bof) then
            %>
                    boolServiceFamilyPnReady = true; // JS Code
            <%
                else
            %>
                    boolServiceFamilyPnReady = false; // JS Code
            <%
                end if
                rs.Close
            %>

            if (!boolServiceFamilyPnReady){
                var msg = 'This Product does not have an associated Service Family Detail.\nPlease click on "Edit Service Family Details" to insert a Service Family Detail for this Product before adding Spare Kit.'
                alert(msg);
                return;
            }

            strID = window.parent.showModalDialog("<%=AppRoot %>/Service/SpareKitDetails.asp?PVID=" + ProductVersionID + "&DRID=" + DeliverableRootId + "&SFPN=" + ServiceFamilyPn + 
                "&SKID=" + SpareKitId + "&CID=" + CategoryId, "", "dialogWidth:635px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")

            // Determine if the user canceled the action 
            if ((strID == "cancel") || (strID == "undefined")) return;
            
    	    var sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>";

            // Create Current Category location Anchor
	        if(sLocation.indexOf("&Anchor")>-1)
	        {
		        sLocation = sLocation.split("&Anchor")[0] + "&Anchor=C" + CategoryId;
            }else{
		        sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>&Anchor=C"+CategoryId;
	        }
		
            // Retain Category Filter, if applicable
            if(bCategoryFilter)
            {
                sLocation+="&CatIDs="+getListSelections(document.getElementById("CategoryList"));
            }

    	    window.location=sLocation;
        }

        function ShowMapDetails(ProductVersionId, SpareKitId, SpareKitMapId, ProductBrandId, CategoryId) {
            var strID;
            var url = "<%=AppRoot %>/Service/SpareKitMap.asp?PVID=" + ProductVersionId + "&SKID=" + SpareKitId + "&MapId=" + SpareKitMapId + "&PBID=" + ProductBrandId;
            modalDialog.open({dialogTitle:'', dialogURL:''+url+'', dialogHeight:700, dialogWidth:900, dialogResizable:true, dialogDraggable:true}); 
            

            //save SpareKitId for return function: ---
            globalVariable.save(SpareKitId, 'spare_kit_map_id');

        }

        function ShowMapDetails_return(strID)
        {
            var SpareKitId = globalVariable.get('spare_kit_map_id');

            // Determine if the user canceled the action 
            if ((strID == "cancel") || (strID == "undefined") || (strID == undefined)) return;

            var sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>";

            // Create Current Category location Anchor
            if(sLocation.indexOf("&Anchor")>-1)
            {
                sLocation = sLocation.split("&Anchor")[0] + "&Anchor=" + SpareKitId;
            }else{
                sLocation="<%=AppRoot%>/pmview.asp?<%=strNoneQS%>&Anchor=" + SpareKitId;
            }
		
            // Retain Category Filter, if applicable
            if(bCategoryFilter)
            {
                sLocation+="&CatIDs="+getListSelections(document.getElementById("CategoryList"));
            }

            window.location=sLocation;
        }

<% If (Not blnIsOSSPUser) Then %>

        function rslCategoryRow_onclick() {
            var node = window.event.srcElement;

            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            ShowRslDetails(node.PVID, '', node.SFPN, '', node.CID);
        }

        function RslRow_OnMouseDown() {
            if (event.button == 2) {
                RtClickMenu();
                return;
            }
        }

        function mapRow_OnMouseDown() {
            if (event.button == 2) {
                mapRtClickMenu();
                return;
            }
        }





        function ShowSaManagement(ProdID, RootID) {
            strID = window.showModalDialog("<%=AppRoot %>/Deliverable/commodity/SubAssembly.asp?ProductID=" + ProdID + "&RootID=" + RootID, "", "dialogWidth:500px;dialogHeight:350px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        }

        var oPopup = window.createPopup();
        oPopup.document.createStyleSheet("<%=AppRoot %>/style/menu.css");

        var _productVersionID;
        var _deliverableRootID;
        var _serviceFamilyPn;
        var _spareKitID;
        var _categoryID;
        var _spareKitMapID;
        var _productBrandID;

        function RtClickMenu() {
            var lefter = event.clientX;
            var topper = event.clientY;
            var popupBody;

            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            _productVersionID = node.PVID;
            _deliverableRootID = node.DRID;
            _serviceFamilyPn = node.SFPN;
            _spareKitID = node.SKID;
            _categoryID = node.CID;

            oPopup.document.body.innerHTML = window.PopUpMenu.innerHTML;

            oPopup.show(lefter, topper, 150, 85, document.body);

            //Adjust window size
            if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
                NewHeight = oPopup.document.body.scrollHeight;
                NewWidth = oPopup.document.body.scrollWidth;
                oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
            }

        }

        function mapRtClickMenu() {
            // The variables "lefter" and "topper" store the X and Y coordinates
            // to use as parameter values for the following show method. In this
            // way, the popup displays near the location the user clicks. 
            var lefter = event.clientX;
            var topper = event.clientY;
            var popupBody;

            var node = window.event.srcElement;
            while (node.nodeName.toLowerCase() != "tr") {
                node = node.parentElement;
            }

            _productVersionID = node.PVID;
            _deliverableRootID = node.DRID;
            _serviceFamilyPn = node.SFPN;
            _spareKitID = node.SKID;
            _categoryID = node.CID;
            _productBrandID = node.PBID;
            _spareKitMapID = node.MAPID;
            
            oPopup.document.body.innerHTML = document.getElementById("MapContextMenu").innerHTML;

            oPopup.show(lefter, topper, 150, 85, document.body);

            //Adjust window size
            if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
                NewHeight = oPopup.document.body.scrollHeight;
                NewWidth = oPopup.document.body.scrollWidth;
                oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
            }

        }



        function MenuProperties() { }

        function MenuShowMapRowProperties() { }


        function MenuAddMapRow() {
            ShowMapDetails(_productVersionID, _spareKitID, 0, _productBrandID, _categoryID);
         }
  
        function MenuAdd(){
        	ShowRslDetails(_productVersionID, _deliverableRootID, _serviceFamilyPn, '', _categoryID);
     }

     function MenuShowSaProperties() {
         ShowSaManagement(_productVersionID, _deliverableRootID);
     }

        function MenuDelete() {
            DeleteSpareKit(_serviceFamilyPn, _spareKitID);
        }

        function DeleteSpareKit(ServiceFamilyPn, SpareKitId) {
            var strID;
            var response = confirm("Are you sure you want to delete this record?");
            //alert("ServiceFamilyPn:" + ServiceFamilyPn + " SpareKitId:" + SpareKitId);
            if (response) {
                jsrsExecute("<%=AppRoot %>/Service/rsService.asp", DeleteSpareKitCallback, "DeleteSpareKit", Array(ServiceFamilyPn, String(SpareKitId), _userFullName));
            }
        }

        function DeleteSpareKitCallback(returnString) {

            if ((returnString != "")&&(returnString.indexOf("ERROR")==-1)) {
                //alert(returnString);
                var row = document.getElementById("SK" + returnString);
                row.style.display = "none";
            }
            else {
                alert("Error occured while deleting this record.\n" + returnString);
            }
        }

        function exportRsl(ProductVersionId) {
            document.getElementById("trCompare").style.display = "";
            document.getElementById("selCompareDtDiv").innerHTML = "<b>Loading ...</b>";
            document.getElementById("popup_submit").disabled = true;
            document.getElementById("chkNewMatrix").checked = false;
            document.getElementById("chkPublish").checked = false;

            var popup_form = document.getElementById("popup_form");
            popup_form.action = "/iPulsar/ExcelExport/ServiceRsl.aspx";

            popup_show('popup', 'popup_drag', 'popup_exit', 'screen-center', 10, 10);

            jsrsExecute("<%=AppRoot %>/Service/rsService.asp", getRslPublishDatesCallback, "GetRslPublishDates", String(ProductVersionId));
        }

        function getRslPublishDatesCallback(returnString) {
            var chkNewMatrix = document.getElementById("chkNewMatrix");
            var trCompare = document.getElementById("trCompare");
            document.getElementById("selCompareDtDiv").innerHTML = returnString;
            document.getElementById("popup_submit").disabled = false;

            if (returnString == "") {
                trCompare.style.display = "none";
                chkNewMatrix.checked = true;
            }

        }
<% End If %>      
        //-->

        function batchUpdSkuAV (pbID, pvName, serviceFamilyPn){
           var url = "/Pulsar/SpareKit/BatchUpdateSpareKit?pbID=" + pbID + "&pvName=" + pvName + "&serviceFamilyPn="+serviceFamilyPn + "&kmat=" + $('#hidKmat').val();
           var WinTop = Math.floor((screen.height - 500) / 4);
           window.open(url, "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
            $('#btnRefreshRSL').show();
        }
    </script>

</head>
<body onload="window_onload();">
    <!-- PMView Header -->
    <% 
IF blnIsOSSPUser THEN
    %>
    <span id="productNameTitle" style="font: bold medium Verdana;">
        <%= sProductName%>
        SERVICE Information 
        <%IF sFusionRequirements = true THEN %>(Pulsar)<%ELSE%>(Legacy)<%END IF%>
    </span><br />
    <br />
    <%
ELSE
    %>
    <span id="productNameTitle" style="font: bold medium Verdana;">
        <%= sProductName%>
        Information
        <%IF sFusionRequirements = true THEN %>(Pulsar)<%ELSE%>(Legacy)<%END IF%>
    </span><br />
    <br />
    <%
END IF
    %>
    
    <%if (clng(request("ID")) = 344 or clng(request("ID")) = 347 or clng(request("ID")) = 1107)then %>
    <span nowrap id="EditLink" style="display: none"></span>
    <%'malichi 07/19/2016, Product Backlog Item 16765: Marketing role needs permissions to Edit Product (Pulsar product) %>
    <%else
        if sFusionRequirements = true then 'Pulsar Product %>
            <%if trim(PVID) <> "-1" and blnCanEditProduct then%>
            <span nowrap id="EditLink" style="display: none">
                <font size="1"><a href="javascript:ShowProperties(<%=PVID%>,0,1)">Edit Product</a></font><font face="verdana" size="1" color="black"> | </font>
            </span>
            <span nowrap id="CloneLink" style="display: none">
                <font size="1"><a href="javascript:ShowProperties(<%=PVID%>,1,1)">Clone Product</a></font><font face="verdana" size="1" color="black"> | </font>
            </span>
            <%else%>
            <span nowrap id="EditLink" style="display: none"></span>
            <span nowrap id="CloneLink" style="display: none"></span>
            <%end if%>
        <%else 'Legacy Product %>

        <%  if (not badministrator) then %>
        <span nowrap id="EditLink" style="display: none"></span>
        <%  elseif ((m_IsToolPm )and trim(strProdType) = "2") then%>
        <%      if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowProperties(<%=PVID%>)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowProperties(<%=PVID%>, 1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%      else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%      end if%>
        <%elseif m_IsActivePm and trim(strProdType) <> "2" then%>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowProperties(<%=PVID%>)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowProperties(<%=PVID%>, 1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%elseif m_IsCommodityPM and trim(strProdType) <> "2" then ' %>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1,1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%elseif m_IsSCFactoryEngineer and trim(strProdType) <> "2" then %>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,2)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1,1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%elseif m_IsAccessoryPM and trim(strProdType) <> "2" then %>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,3)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1,1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%elseif m_IsHardwarePm and trim(strProdType) <> "2" then %>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,4)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1,1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%elseif m_IsTestLead and trim(strProdType) <> "2" then%>
        <%if trim(PVID) <> "-1" then%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1)">Edit Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1"><a href="javascript:ShowCommodityProperties(<%=PVID%>,1,1)">Clone Product</a></font><font
                face="verdana" size="1" color="black"> | </font>
        </span>
        <%else%>
        <span nowrap id="EditLink" style="display: none"></span>
        <span nowrap id="CloneLink" style="display: none"></span>
        <%end if%>
        <%else%>
        <span nowrap id="EditLink" style="display: none">
            <font size="1">
                <!--Contact the PM to edit product properties-->
            </font>
        </span>
        <span nowrap id="CloneLink" style="display: none">
            <font size="1">
                <!--Contact the PM to edit product properties-->
            </font>
        </span>
        <%end if%>
        <%end if%>
    <%end if %>
    <span id="loadingProgress"></span>
    <span id="RFLink" style="display: none"><a href="javascript:RemoveFavorites(<%=PVID%>)">
        <font face="verdana" size="1">Remove From Favorites</font></a></span> <span id="AFLink"
            style="display: none"><a href="javascript:AddFavorites(<%=PVID%>)"><font face="verdana"
                size="1">Add To Favorites</font></a></span> <span id="StatusLink" style="display: none">
                    <a href="<%=AppRoot %>/Productstatus.asp?Product=<%=sDisplayedProductName%>&ID=<%=PVID%>">
                        <font face="verdana" size="1">Real-Time Status Report</font></a>&nbsp;|&nbsp;</span>
    <%if strDisplayedList <> CurrentUserDefaultTab and trim(strProdType) <> "2" and (Not blnIsOSSPUser) then%>
    <span id="DefaultTabLink">&nbsp;|&nbsp;<a href="javascript:SetDefaultDisplay('<%=strDisplayedList%>',<%=CurrentUserID%>)"><font
        face="verdana" size="1">Set Default List</font></a></span>
    <%end if%>
    <br />
    <br />
    <%
IF Not blnIsOSSPUser THEN
    %>
    <table id="Table1" class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0"
        cellpadding="2">
        <tr bgcolor="<%=strTitleColor%>">
            <td id="Td1" width="10">
                <font size="1" color="white">&nbsp;<a href="javascript:SelectTab('DCR',1)">Change&nbsp;Requests</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td2" width="10">
                <font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Action',1)">Action&nbsp;Items</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td7" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('OTS',1)">Observations</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td15" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Agency',1)">Certifications</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td8" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('PMR',1)">SMR</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td6" width="10" bgcolor="wheat">
                <font size="1" color="black">&nbsp;Service&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td9" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('General',1)">General</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
        </tr>
        <tr bgcolor="<%=strTitleColor%>">
            <td id="Td10" width="10">
                <font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Requirements',1)">Requirements</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td11" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Country',1)">Localization</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td12" width="10">
                <font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Local',1)">Images</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td13" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Deliverables',1)">Deliverables</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td14" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Schedule',1)">Schedule</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td16" width="10">
                <font size="1" color="black">&nbsp;<a href="javascript:SelectTab('SCM',1)">Supply&nbsp;Chain</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
            <td id="Td17" width="10">
                <font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Documents',1)">Documents</a>&nbsp;&nbsp;&nbsp;</font>
            </td>
        </tr>
    </table>
    <%
END IF
    %>
    <span id="AddCallsLink"><span id="CallDisplayFilters">
        <br />
        <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
            <tr>
                <td valign="top">
                    <table width="100%">
                        <tr>
                            <td valign="top">
                                <font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font>
                            </td>
                        </tr>
                    </table>
                </td>
                <td width="100%">
                    <table width="100%">
                        <tr>
                            <td>
                                <b>Report:</b>
                            </td>
                            <td width="100%">
                                <% 
	                                Select Case ServiceReport
		                                Case "rsl"
                                %>
                                RSL | <a href="#" onclick="ServiceReport('spb');">SPB</a> | <a href="#" onclick="ServiceReport('sku');">
                                    SKU BOMs</a><span style="color: Red; font-weight: bold; cursor: pointer;"></span>
                                <%
		                                Case "spb"
                                %>
                                <a href="#" onclick="ServiceReport('rsl');">RSL</a> | SPB | <a href="#" onclick="ServiceReport('sku');">
                                    SKU BOMs</a><span style="color: Red; font-weight: bold; cursor: pointer;"></span>
                                <%
		                                Case "ecr"
                                %>
                                <a href="#" onclick="ServiceReport('rsl');">RSL</a> | <a href="#" onclick="ServiceReport('spb');">
                                    SPB</a> | <a href="#" onclick="ServiceReport('sku');">SKU BOMs</a><span style="color: Red;
                                        font-weight: bold; cursor: pointer;"></span>
                                <%
		                                Case "sku"
                                %>
                                <a href="#" onclick="ServiceReport('rsl');">RSL</a> | <a href="#" onclick="ServiceReport('spb');">
                                    SPB</a> | SKU BOMs<span style="color: Red; font-weight: bold; cursor: pointer;"></span>
                                <%
		                                Case Else
                                %>
                                RSL | <a href="#" onclick="ServiceReport('spb');">SPB</a> | <a href="#" onclick="ServiceReport('sku');">
                                    SKU LookUp</a><span style="color: Red; font-weight: bold; cursor: pointer;"></span>
                                <%
	                                End Select
                                %>
                            </td>
                        </tr>
                        <%If ServiceReport = "rsl" Then%>
                        <% 
'
' Get the list of Brands for the product.
'
	Set cmd = dw.CreateCommAndSP(cn, "usp_GetBrands4Product")
	dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, PVID
	dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
	rs.Sort = "CombinedName, Name"
    bFirstWrite = True

	Response.Write "<tr><td nowrap><b>Sub Report:&nbsp;&nbsp;</b></td><td width=""100%"">"

    If (Request("BID") = "0" And m_BrandID = "") Then
        Response.Write "Spare Kits"
		m_BrandID = ""
		m_BrandName = ""
		m_BusinessID = ""
    Else
        Response.Write "<a href=""javascript:BrandLink_onClick(0)"">Spare Kits</a>"
    End If
	
    Do Until rs.EOF
        Response.Write "&nbsp;|&nbsp;"
			
		If (Request("BID") = "" And m_BrandID = "") or (CLng(rs("ProductBrandID")) = CLng(Request("BID"))) or (CLng(rs("CombinedProductBrandId")) = CLng(Request("BID"))) Then   'This part is when the Brand is currently being displayed so there is no link                                                           
       		m_BrandID = rs("ProductBrandID")			
            
            if sFusionRequirements = true then 'Pulsar
                 if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
                    if sBrandDisplayed <> rs("CombinedName") then
					    If Not bFirstWrite Then
					      Response.Write "&nbsp;|&nbsp;"
					    end if
                         m_BrandName = rs("CombinedName")
					     m_BrandID = rs("combinedProductBrandId")	
                         Response.Write server.HTMLEncode(m_BrandName)   
				    end if                                        
                else
                   if sBrandDisplayed <> rs("Name") then
					    If Not bFirstWrite Then
						    Response.Write "&nbsp;|&nbsp;"
					    end if					    
                         m_BrandName = rs("Name")
                         Response.Write server.HTMLEncode(m_BrandName )   
				    end if                          
                end if
            else 'Legacy                
                m_BrandName = rs("Name")
                Response.Write server.HTMLEncode(m_BrandName)                   
		    end if   

            m_BusinessID = rs("BusinessID") & ""		                   
            sBrandDisplayed = m_BrandName                  
		Else  ' below part is when the Brand is not currently being displayed so there needs to be a link on the other Brands
            if sFusionRequirements = true then 'Pulsar
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
            else 'Legacy
            	Response.Write "<a href=""javascript:BrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"				                                        
            end if			
		End If

		bFirstWrite = False
		rs.MoveNext
	Loop
	rs.Close
	Response.Write "</td></tr>"
    
End If 
                    %>
                    <%If ServiceReport = "sku" Then%>
                    <% 
'
' Get the list of Brands for the product.
'
	Set cmd = dw.CreateCommAndSP(cn, "usp_GetBrands4Product")
	dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, PVID
	dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
    rs.Sort = "CombinedName, Name"
    bFirstWrite = True
		
	Response.Write "<tr><td nowrap><b>Sub Report:&nbsp;&nbsp;</b></td><td width=""100%"">"

    If (Request("BID") = "0" And m_BrandID = "") Or (Request("BID") = "" And m_BrandID = "") Then
        'Response.Write "<b>Select a Brand -></b>&nbsp;"
		m_BrandID = ""
		m_BrandName = ""
		m_BusinessID = ""
    Else
       ' Response.Write "<a href=""javascript:BrandLink_onClick(0)"">All</a>"
       'Response.Write "<b>Select a Brand -></b>&nbsp;"
    End If
    
    Dim intIteration' As Integer
    intIteration=0

	Do Until rs.EOF

        If(intIteration>0) Then	
            Response.Write "&nbsp;|&nbsp;"
		End If

		If (Request("BID") = "" And m_BrandID = "") or (CLng(rs("ProductBrandID")) = CLng(Request("BID"))) or (CLng(rs("CombinedProductBrandId")) = CLng(Request("BID"))) Then   'This part is when the Brand is currently being displayed so there is no link                                                           
       
			m_BrandID = rs("ProductBrandID")		
            if sFusionRequirements = true then 'Pulsar
                if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
                    m_BrandName = rs("CombinedName")
                    m_BrandID = rs("CombinedProductBrandId")		            
		        else
                    m_BrandName = rs("Name") 'no combined name so display the name from the Brand table
                end if
            else 'Legacy
                m_BrandName = rs("Name")
		    end if

            if sBrandDisplayed <> m_BrandName then 'have to handle if it is a combined Brand name because we don't want to display the same name multiple times
				If Not bFirstWrite Then
				    Response.Write "&nbsp;|&nbsp;"
				end if
				Response.Write server.HTMLEncode(m_BrandName)
			end if	
                            		
            sBrandDisplayed = m_BrandName
			m_BusinessID = rs("BusinessID") & ""
		Else
            if sFusionRequirements = true then 'Pulsar
                 if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
                    if sBrandDisplayed <> rs("CombinedName") then
					    If Not bFirstWrite Then
					      Response.Write "&nbsp;|&nbsp;"
					    end if
					    Response.Write "<a href=""javascript:SKUBOMBrandLink_onClick(" & rs("CombinedProductBrandId") & ")"">" & server.HTMLEncode(rs("CombinedName")) & "</a>"
				    end if
                    
                    sBrandDisplayed = rs("CombinedName")    
                 else
                    if sBrandDisplayed <> rs("Name") then
					    If Not bFirstWrite Then
						    Response.Write "&nbsp;|&nbsp;"
					    end if					    
                         Response.Write "<a href=""javascript:SKUBOMBrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"
				    end if                            
                 
                    sBrandDisplayed = rs("Name")
                 end if
			else 'Legacy
                  Response.Write "<a href=""javascript:SKUBOMBrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"           
            end if                            
		End If
		bFirstWrite = False
        
        intIteration=intIteration+1

		rs.MoveNext
Response.Flush()
	Loop
	
	rs.Close
	Response.Write "</td></tr>"
    
End If 

'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'   
'   CATEGORY FILTER 
'
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
IF ServiceReport="rsl" OR ServiceReport="spb" THEN ' OR ServiceReport="sku" THEN

        blnPopFilterCategory=true
        strSelCatIDs=Request.QueryString("CatIDs")

        If Len(Trim(strSelCatIDs))=0 Then
            blnApplyCatFilter=false
            strSelCatIDs="null"
        Else
            blnApplyCatFilter=true
            strSelCatIDs="|" & strSelCatIDs & "|"
        End If


                        %>
                        <tr>
                            <td nowrap>
                                <b>Category Filter:</b>
                            </td>
                            <td width="100%">
                                <%
    '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
        If blnApplyCatFilter Then
                                %>
                                <input type="radio" id="FilterOptionsNone" name="FilterOptions" value="0" onclick="javascript:toggleCategoryFilter(this);" />None&nbsp;<input
                                    type="radio" id="FilterOptionsOther" name="FilterOptions" value="1" onclick="javascript:toggleCategoryFilter(this);"
                                    checked />Other
                            </td>
                            <%
        Else
                            %>
                            <input type="radio" id="FilterOptionsNone" name="FilterOptions" value="0" onclick="javascript:toggleCategoryFilter(this);"
                                checked />None&nbsp;<input type="radio" id="FilterOptionsOther" name="FilterOptions"
                                    value="1" onclick="javascript:toggleCategoryFilter(this);" />
                    Other
                </td>
                <%
        End If
    '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
                %>
            </tr>
            <%
        If blnApplyCatFilter Then
            %>
            <tr id="FilterActions">
                <%
        Else
                %>
                <tr id="FilterActions" style="display: none">
                    <%
        End If
                    %>
                    <td nowrap align="right">
                        <a href="javascript:applyCatFilter();">Apply</a>
                    </td>
                    <td width="100%">
                        <select id="CategoryList" title="Hold down the <CTRL> key to select multiple non-contiguous Categories"
                            multiple>
                        </select>
                    </td>
                </tr>
                <%
END IF
''' RON EDIT ---  display filters for Spare Kit # and AV #
 if m_BrandName <> "" then          
 

  %>     
<tr>
<td ><b>Spare Kit Filter:</b></td>
  
  <%if request.QueryString("sk") = "" then%>
  <td width="100%">
  
  <input type="radio" id="Radio1" name="FilterOptions2" value="0" onclick="javascript:toggleSKAVFilter(this);"  checked />None&nbsp;<input type="radio" id="Radio2" name="FilterOptions2" onclick="javascript:toggleSKAVFilter(this);" value="1" />Other
  </td>
</tr>
  <tr id="sparekitfilter" style="display: none">
    <td align="right"><a href="javascript:applySKFilter();">Apply</a></td><td><input id="sktext" type="text"/></td>
  </tr>
  <%else %>
  <td width="100%">
  
  <input type="radio" id="Radio1" name="FilterOptions2" value="0" onclick="javascript:toggleSKAVFilter(this);"  />None&nbsp;<input type="radio" id="Radio2" name="FilterOptions2" onclick="javascript:toggleSKAVFilter(this);" value="1" checked/>Other
  </td>
</tr>
  <tr id="sparekitfilter">
    <td align="right"><a href="javascript:applySKFilter();">Apply</a></td><td><input id="sktext" type="text" value = "<%=request.QueryString("sk") %>"/></td>
  </tr>

  <%end if %>

  

                <%end if
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************


If ServiceReport = "ecr" Then%>
                <tr>
                    <td>
                        <b>Status:</b>
                    </td>
                    <td width="100%" colspan="2">
                        <% 
	dim strEcrStatus
	strEcrStatus = ""
	on error resume next
	strEcrStatus = Request.Cookies("EcrStatus")
	on error goto 0
	Select Case strEcrStatus
		Case "open"
                        %>
                        Open | <a href="#" onclick="ShowEcrStatus('closed');">Closed</a> | <a href="#" onclick="ShowEcrStatus('all');">
                            All</a>
                        <%
		Case "closed"
                        %>
                        <a href="#" onclick="ShowEcrStatus('open');">Open</a> | Closed | <a href="#" onclick="ShowEcrStatus('all');">
                            All</a>
                        <%
		Case Else
                        %>
                        <a href="#" onclick="ShowEcrStatus('open');">Open</a> | <a href="#" onclick="ShowEcrStatus('closed');">
                            Closed</a> | All
                        <%
	End Select		
                        %>
                    </td>
                </tr>
                <%End If 'ServiceReport = "ecr"%>
               </table>
        </td> </tr> </table> </span>      
      
        <% If ServiceReport = "spb" Then %>
        <br />
        <font face="verdana" size="1"><b>Tools:&nbsp;</b></font> <span style="font-size: xx-small;
            font-family: Verdana"><a href="#" onclick="QuickSearch();">QuickSearch</a> | <a href="#"
                onclick="AdvancedSearch();">Advanced Search</a>
            <% If (Not blnIsOSSPUser) Then %>
            | <a href="#" onclick="SetServiceFamilyPn(<%=PVID%>);">Edit Service Family Details</a>
            | <a href="#" onclick="exportSpb('<%=sServiceFamilyPn %>');">Export SPB</a>
            <% End If %>
            <% 'End If %>
        </span>
        <br />
        <br />
        <span id="theSectionTitle" style="font-size: x-small; font-weight: bold;">
            <%= m_BrandName%>
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;-&nbsp;Service Program BOM"%></span>
        <%
'##########################################################
'#
'# Draw Service Program BOM
'#
'##########################################################

If sServiceFamilyPn = "" Then
    Response.Write "No Service Family Partnumber."
Else

Dim iVersions

Set cmd = dw.CreateCommandSPwTimeout(cn, "usp_SelectServiceProgramBomSS", 240)
dw.CreateParameter cmd, "@p_ServiceFamilyPn", adVarchar, adParamInput, 10, sServiceFamilyPn
Set rs = dw.ExecuteCommandReturnRS(cmd)

'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' ADDED TO IMPLEMENT CATEGORY FILTER
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
strCatNames=""
strCatIDs=""
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

If Not rs.EOF Then
        %>
        <span id="SnapShotTS" style="font: bold xx-small verdana; color: Red;">Snapshot taken
            at
            <%= rs("ExportTime") %></span>
        <% End If %>
        <div style="overflow-y: scroll; overflow-x: scroll;" id="DIV1">
            <table class="MatrixTable">
                <tr>
                    <th class="pinnedRow">
                        Part Type
                    </th>
                    <th class="pinnedRow">
                        Level
                    </th>
                    <th class="pinnedRow">
                        OSSP Orderable
                    </th>
                    <th class="pinnedRow">
                        Line Item
                    </th>
                    <th class="pinnedRow">
                        Spare Kit Pn
                    </th>
                    <th class="pinnedRow">
                        Rev
                    </th>
                    <th class="pinnedRow">
                        Cross Plant Status
                    </th>
                    <th class="pinnedRow">
                        Qty
                    </th>
                    <th class="pinnedRow">
                        Pri<br />
                        Alt<br />
                        Gen
                    </th>
                    <th class="pinnedRow">
                        HP SA or Component
                    </th>
                    <th class="pinnedRow">
                        HP Part Description
                    </th>
                    <th class="pinnedRow">
                        ODM Part No.
                    </th>
                    <th class="pinnedRow">
                        ODM Part Description
                    </th>
                    <th class="pinnedRow">
                        ODM Bulk Part No.
                    </th>
                    <th class="pinnedRow">
                        ODM Production MOQ
                    </th>
                    <th class="pinnedRow">
                        ODM Post-Production MOQ
                    </th>
                    <th class="pinnedRow">
                        Part Supplier (ODM/OEM)
                    </th>
                    <th class="pinnedRow">
                        Model / Mfg Pn
                    </th>
                    <th class="pinnedRow">
                        Comments
                    </th>
                    <%

                    %>
                </tr>
                <%


Dim sLastAv
Dim sLastSa
Dim sLastComponent
Dim sLastCategory


Do Until rs.EOF
    
    strCurrCatID=rs("SpsKitCatSortOrder") ' ADDED TO IMPLEMENT CATEGORY FILTER

    If Trim(rs("SpsKitPn") & "") <> "" Then

         '*****************************************************************************************************************************
         ' ADDED/ALTERED TO IMPLEMENT CATEGORY FILTER          
         '*****************************************************************************************************************************
         If blnPopFilterCategory Then
			    ' Accummulate Categories and populate list after this routine is done,  MAY CONSIDER AJAX CALL TO NEW PAGE
                IF(InStr(1, "|" & strCatIDs & "|", "|" & strCurrCatID & "|")=0) THEN
			        If LEN(TRIM(strCatIDs))=0 Then
				        strCatIDs=rs("SpsKitCatSortOrder")
				        strCatNames=rs("SpsKitDelCat")	
			        Else
				        strCatIDs=strCatIDs & "|" & rs("SpsKitCatSortOrder")
				        strCatNames=strCatNames & "|" & rs("SpsKitDelCat")	
			        End If
                END IF
		End If
        '*****************************************************************************************************************************
        IF (blnApplyCatFilter AND InStr(1, strSelCatIDs, "|" & strCurrCatID & "|")>0) OR (NOT blnApplyCatFilter) THEN

            If sLastAv <> Trim(rs("SpsKitPn") & "") Then
                sLastAv = Trim(rs("SpsKitPn") & "")
                If Trim(rs("SaPn") & "") <> "" Then sLastSa = Trim(rs("SaPn"))
                If Trim(rs("PartPn") & "") <> "" Then sLastComponent = Trim(rs("PartPn"))
                DrawKitRow rs ', rsVer, iVersions
            ElseIf sLastSa <> Trim(rs("SaPn")) AND Trim(rs("SaPn") & "") <> "" Then
                sLastSa = Trim(rs("SaPn") & "")
                If Trim(rs("PartPn") & "") <> "" Then sLastComponent = Trim(rs("PartPn"))
                DrawSaRow rs ', iVersions 
            ElseIf sLastComponent <> Trim(rs("PartPn") & "") AND Trim(rs("PartPn") & "") <> "" Then
                sLastComponent = Trim(rs("PartPn") & "")
                DrawPartRow rs ', iVersions
            End If

        END IF
        '*****************************************************************************************************************************

    End If



    rs.MoveNext
Loop

rs.Close
End If ' Not Rs.Eof

Sub DrawKitRow(rs) ', rsVer, iVersions)
  If ValidRow(rs, "kit") Then
    'rsVer.Filter= "SpsKitPn = '" & rs("SpsKitPn") & "'"
'If (Not blnIsOSSPUser) Then 
    Response.Write "<tr class=""MatrixSpareKitRow"" PVID=""" & PVID & """ SKID=""" & rs("ID") & """ HPPN=""" & rs("SpsKitPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" onclick=""return spbPart_onClick()"" >"
'Else
'	Response.Write "<tr class=""MatrixSpareKitRow"" PVID=""" & PVID & """ SKID=""" & rs("ID") & """ HPPN=""" & rs("SpsKitPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" >"
'End If

    Response.Write "<td nowrap>" & rs("SpsKitDelCat") & "</td>"
    Response.Write "<td>2</td>"
    If rs("SpsKitOsspOrderable") = True Then
        Response.Write "<td>Y</td>"
    Else
        Response.Write "<td>FG</td>"
    End If
    Response.Write "<td></td>"
    Response.Write "<td nowrap>" & rs("SpsKitPn") & "</td>"
    Response.Write "<td>" & rs("SpsKitRev") & "</td>"
    Response.Write "<td>" & rs("SpsXplantStatus") & "</td>"
    Response.Write "<td></td>"
    Response.Write "<td></td>"
    Response.Write "<td></td>"
    Response.Write "<td>" & rs("SpsKitDescription") & "</td>"
	Response.Write "<td>" & rs("SpsOdmPartNo") & "</td>"
	Response.Write "<td>" & rs("SpsOdmPartDescription") & "</td>"
	Response.Write "<td>" & rs("SpsOdmBulkPartNo") & "</td>"
	Response.Write "<td>" & rs("SpsOdmProductionMoq") & "</td>"
	Response.Write "<td>" & rs("SpsOdmPostProductionMoq") & "</td>"
	Response.Write "<td>" & rs("SpsSupplier") & "</td>"
    Response.Write "<td></td>"
    Response.Write "<td>" & rs("SpsComments") & "</td>"
    'Do Until rsVer.EOF
    '    If rsVer("VersionSupported") & "" <> "" Then
    '        Response.Write "<td>x</td>"
    '    Else
    '        Response.Write "<td></td>"
    '    End If
    '    rsVer.MoveNext
    'Loop
    Response.Write "</tr>"

    DrawSaRow rs ', iVersions
  End If
End Sub

Sub DrawSaRow(rs)
    Dim i
    
    If rs("SaPn") & "" <> "" And ValidRow(rs, "sa") Then
        If rs("SaMatlType") & "" = "HALB" Then
	    'If Not blnIsOSSPUser Then
            	Response.Write "<tr class=""MatrixAvRow"" PVID=""" & PVID & """ HPPN=""" & rs("SaPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" onclick=""return spbPart_onClick()"" >"
	    'Else
            '	Response.Write "<tr class=""MatrixAvRow"" PVID=""" & PVID & """ HPPN=""" & rs("SaPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()""  >"
	    'End If
        Else
	   ' If Not blnIsOSSPUser Then
	            Response.Write "<tr PVID=""" & PVID & """ HPPN=""" & rs("SaPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" onclick=""return spbPart_onClick()"" >"
	    'Else
	    '        Response.Write "<tr PVID=""" & PVID & """ HPPN=""" & rs("SaPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" >"

	    'End If
        End If
        Response.Write "<td style=""color:gray;"">" & rs("SpsKitDelCat") & "</td>"
        Response.Write "<td>3</td>"
        If rs("SaOsspOrderable") & "" Then
            Response.Write "<td>Y</td>"
        Else
            Response.Write "<td>N</td>"
        End If
        Response.Write "<td>" & rs("SaBomItemNo") & "</td>"
        Response.Write "<td style=""color:gray;"" nowrap>" & rs("SpsKitPn") & "</td>"
        Response.Write "<td>" & rs("SaRev") & "</td>"
        Response.Write "<td>" & rs("SaXplantStatus") & "</td>"
        Response.Write "<td>" & rs("SaQty") & "</td>"
        Response.Write "<td>" & rs("SaPriAlt") & "</td>"
        Response.Write "<td nowrap>" & rs("SaPn") & "</td>"
        Response.Write "<td>" & rs("SaDescription") & "</td>"
        Response.Write "<td>" & rs("SaOdmPartNo") & "</td>"
        Response.Write "<td>" & rs("SaOdmPartDescription") & "</td>"
        Response.Write "<td>" & rs("SaOdmBulkPartNo") & "</td>"
        Response.Write "<td>" & rs("SaOdmProductionMoq") & "</td>"
        Response.Write "<td>" & rs("SaOdmPostProductionMoq") & "</td>"
        Response.Write "<td>" & rs("SaSupplier") & "</td>"
        Response.Write "<td>" & rs("SaModel") & "</td>"
        Response.Write "<td>" & rs("SaComments") & "</td>"
     '   For i = 1 To iVersions
     '       Response.Write "<td></td>"
     '   Next
        Response.Write "</tr>"

        DrawPartRow rs' , iVersions
    End If
End Sub

Sub DrawPartRow(rs)
    Dim i
    
    If rs("PartPn") & "" <> "" And ValidRow(rs, "part") Then
	'If Not blnIsOSSPUser Then
	        Response.Write "<tr PVID=""" & PVID & """ HPPN=""" & rs("PartPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" onclick=""return spbPart_onClick()"" ><td style=""color:gray;"">" & rs("SpsKitDelCat") & "</td>"
	'Else
	'        Response.Write "<tr PVID=""" & PVID & """ HPPN=""" & rs("PartPn") & """ SFPN=""" & sServiceFamilyPn & """ onmouseover=""return spbRow_onMouseOver()"" onmouseout=""return spbRow_onMouseOut()"" ><td style=""color:gray;"">" & rs("SpsKitDelCat") & "</td>"
	'End If
        Response.Write "<td>4</td>"
        If rs("PartOsspOrderable") & "" Then
            Response.Write "<td>Y</td>"
        Else
            Response.Write "<td>N</td>"
        End If
        Response.Write "<td>" & rs("PartBomItemNo") & "</td>"
        Response.Write "<td style=""color:gray;"" nowrap>" & rs("SpsKitPn") & "</td>"
        Response.Write "<td>" & rs("PartRev") & "</td>"
        Response.Write "<td>" & rs("PartXplantStatus") & "</td>"
        Response.Write "<td>" & rs("PartQty") & "</td>"
        Response.Write "<td>" & rs("PartPriAlt") & "</td>"
        Response.Write "<td nowrap>" & rs("PartPn") & "</td>"
        Response.Write "<td>" & rs("PartDescription") & "</td>"
        Response.Write "<td>" & rs("PartOdmPartNo") & "</td>"
        Response.Write "<td>" & rs("PartOdmPartDescription") & "</td>"
        Response.Write "<td>" & rs("PartOdmBulkPartNo") & "</td>"
        Response.Write "<td>" & rs("PartOdmProductionMoq") & "</td>"
        Response.Write "<td>" & rs("PartOdmPostProductionMoq") & "</td>"
        Response.Write "<td>" & rs("PartSupplier") & "</td>"
        Response.Write "<td>" & rs("PartModel") & "</td>"
		    Response.Write "<td>" & rs("PartComments") & "</td>"
      '  For i = 1 To iVersions
      '      Response.Write "<td></td>"
      '  Next
        Response.Write "</tr>"
    End If
End Sub
                %>
            </table>
        </div>
        <% End If 'ServiceReport = "sbp" %>
        <% If ServiceReport = "ecr" Then%>
        <br />
        <span style="font-size: xx-small; font-family: Verdana"><a href="#" onclick="AddEcr(<%=PVID%>);">
            New ECR</a> </span>
        <br />
        <br />
        <span style="font-size: x-small; font-weight: bold;">
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;-&nbsp;Service ECR" %>
        </span>
        <%
'##########################################################
'#
'# Draw ECR
'#
'##########################################################

on error resume next
Dim strCookie
Dim strStatusTxt
strCookie = ""
strCookie = Request.Cookies("EcrStatus")
on error goto 0

SELECT CASE strCookie
    CASE "all"
        strEcrStatus = 0
        strStatusTxt = " "
    CASE "open"
        strEcrStatus = 1
        strStatusTxt = " open "
    CASE "closed"
        strEcrStatus = 2
        strStatusTxt = " closed "
    CASE ELSE
        strEcrStatus = 1
        strStatusTxt = " open "
END SELECT

Set cmd = dw.CreateCommandSPwTimeout(cn, "spListActionItems", 240)
dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 10, PVID
dw.CreateParameter cmd, "@Type", adInteger, adParamInput, 10, 7
dw.CreateParameter cmd, "@Status", adInteger, adParamInput, 10, strEcrStatus
Set rs = dw.ExecuteCommandReturnRS(cmd)

Dim ColumnCount
ColumnCount = 0
Dim strDcrStatus
Dim strTarget
Dim strActual
Dim strAvailableForTest
Dim strStatus
If not(rs.EOF and rs.BOF) then%>
        <table id="TableECR" class="MatrixTable">
            <thead>
                <tr>
                    <th onclick="SortTable('TableDCR', 0,1,2);" nowrap width="50">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Number</span>
                    </th>
                    <%if trim(PVID) = "-1" then
		ColumnCount = ColumnCount + 1
                    %>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount%> ,0,2);" nowrap width="120">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Product</span>
                    </th>
                    <%end if%>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,0,2);" nowrap width="80">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Submitter</span>
                    </th>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+2%> ,0,2);" nowrap width="80">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Status</span>
                    </th>
                    <%  if false then %>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+3%> ,2,2);" nowrap width="80">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Target Date</span>
                    </th>
                    <%
        ColumnCount = ColumnCount +3
    else
        ColumnCount = ColumnCount + 2
    end if 

	if trim(strDcrStatus) <> "1" then 'strStatusID <> "1" then	%>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,2,2);" nowrap width="80">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Approved</span>
                    </th>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+2%> ,2,2);" nowrap width="80">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Available</span>
                    </th>
                    <%
		ColumnCount = ColumnCount +2		
    end if%>
                    <th onclick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,0,2);" nowrap width="100%">
                        <span onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">
                            Summary</span>
                    </th>
                </tr>
            </thead>
            <tbody>
                <%do while not rs.EOF  %>
                <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="changerows" onmouseover="return changerows_onmouseover()"
                    onmouseout="return changerows_onmouseout()" onclick="return changerows_onclick()"
                    oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
                    <%
	if isnull(rs("TargetDate")) then
		strTarget = ""
	else
		'if DateDiff("d",date,rs("TargetDate")) > 0 then
			strTarget = formatdatetime(rs("TargetDate"),2)
		'else
		'	strTarget = "<font color=red>" & formatdatetime(rs("TargetDate"),2) & "</font>"
		'end if

	end if
	if isnull(rs("ActualDate")) then
		strActual = "&nbsp;"
	else
		'if DateDiff("d",date,rs("TargetDate")) > 0 then
			strActual = formatdatetime(rs("ActualDate"),2)
		'else
		'	strTarget = "<font color=red>" & formatdatetime(rs("TargetDate"),2) & "</font>"
		'end if

	end if	
	if isnull(rs("AvailableForTest")) then
		strAvailableForTest = "&nbsp;"
	else
		strAvailableForTest = formatdatetime(rs("AvailableForTest"),2)
	end if	

	Select case rs("status")
	case 1
		strStatus = "Open"
	case 2
		strStatus = "Closed"	
	case 3
		strStatus = "Need Info"	
	case 4
		'if strStatusID = "0" and (not isnull(rs("ECNDate")))then
		'	strStatus = "ECN&nbsp;Complete"
		'else
			strStatus = "Approved"
		'end if
	case 5
		strStatus = "Disapproved"
	case 6
		strStatus = "Investigating"
	case else
		strStatus = "N/A"
	end select

	ItemsDisplayed = ItemsDisplayed + 1
                    %>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=rs("ID") & ""%></font>
                    </td>
                    <%if trim(PVID) = "-1" then%>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=rs("Product") & ""%></font>
                    </td>
                    <%end if%>
                    <td nowrap valign="top" class="cell">
                        <font size="1" class="text">
                            <%=rs("Submitter") & ""%></font>
                    </td>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=strStatus%></font>
                    </td>
                    <%if false then %>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=strTarget%></font>
                    </td>
                    <%end if %>
                    <%if trim(strDcrStatus) <> "1" then 'if strStatusID <> "1" then%>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=strActual%></font>
                    </td>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=strAvailableForTest%></font>
                    </td>
                    <%end if%>
                    <td valign="top" class="cell">
                        <font size="1" class="text">
                            <%=rs("summary") & "&nbsp;"%></font>
                    </td>
                </tr>
                <%	rs.MoveNext
	loop
                %>
            </tbody>
        </table>
        <%else%>
        <table id="TableDCR" style="display: none" cellspacing="1" cellpadding="1" width="100%"
            border="0">
            <tr>
                <td>
                    <font face="Verdana" size="2">No" & strStatusTxt & "ECRs found.</font>
                </td>
            </tr>
        </table>
        <%end if

'Action Items
rs.Close
        %>
        <% End If 'ServiceReport = "ecr" %>
        <% If ServiceReport = "rsl" Then
'##########################################################
'#
'# Draw RSL
'#
'##########################################################

        %>
        <br />
        <font face="verdana" size="1"><b>Tools:&nbsp;</b></font> <span style="font-size: xx-small;
            font-family: Verdana"><a href="#" onclick="QuickSearch();">QuickSearch</a> | <a href="#"
                onclick="AdvancedSearch();">Advanced Search</a>
            <% If(Not blnIsOSSPUser) Then %>
            | <a href="#" onclick="SetServiceFamilyPn(<%=PVID%>);">Edit Service Family Details</a>
            | 
            <% End If %>
			
            <% If (m_IsSysAdmin Or m_IsSpdmUser) and (Not blnIsOSSPUser) Then %>
				<% If m_BrandId = "" Then %>
				<a href="#" onclick="doBatchUpdate(<%=PVID%>);">Batch Update</a> |
				<% End If %>
            <% End If %>
			
            <% If(Not blnIsOSSPUser) Then %>
            <a href="#" onclick="exportRsl('<%=PVID %>');">Export RSL</a>
            <% End If  %>
			
			<%
            if m_BrandName <> "" then
                response.Write"| <a href=""#"" onclick=""deleteSkuAV('" &  m_BrandId & "', '" & sServiceFamilyPn & "' , ' " & CurrentUserID & "', '', '');"">Remove AV/Spare Kit Mapping(s)</a>"
                response.Write"| <a href=""#"" onclick=""batchUpdSkuAV('" &  m_BrandId & "', '" & sProductName & "', '" & sServiceFamilyPn & "');"">Batch Update AV/Spare Kit Mapping(s)</a>"
                response.Write"<button type=""button"" id=""btnRefreshRSL"" style=""background-color: #cccccc; color:#000000; font-size: 9px; font-weight:bolder; display:none;"">Refresh the Grid</button>"
            end if
			%>
        </span>
        <br />
        <br />
        <% If m_BrandId = "" Then %>
        <span id="theSectionTitle" style="font-size: x-small; font-weight: bold;">
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;-&nbsp;Recommended Spares List" %>
        </span>
        <div style="overflow-y: scroll; overflow-x: scroll;" id="DIV1" on>
            <table class="MatrixTable">
                <thead>
                    <tr style="height: 20px">
                        <% 
If (m_IsSysAdmin Or m_IsSpdmUser) And (Not blnIsOSSPUser) Then
                        %>
                        <th class="pinnedRow" style="text-align: left; vertical-align: top">
                            <input type="checkbox" id="CBCatAll" name="CBCatAll" onmouseover="getCBCatAllTitle(this)"
                                onclick="toggleCBCatAll(this)" />&nbsp;&nbsp;Status
                        </th>
                        <% 
Else
                        %>
                        <th class="pinnedRow">
                            Status
                        </th>
                        <%
End If
                        %>
                        <th class="pinnedRow">
                            Spare Kit #
                        </th>
                        <th class="pinnedRow">
                            GPG Description
                        </th>
                        <th class="pinnedRow">
                            Status
                        </th>
                        <th class="pinnedRow">
                            Service SA #
                        </th>
                        <th class="pinnedRow">
                            Root ID
                        </th>
                        <th class="pinnedRow">
                            Excalibur Description
                        </th>
                        <th class="pinnedRow">
                            Notes
                        </th>
                        <th class="pinnedRow">
                            Comments
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <%
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' ADDED TO IMPLEMENT CATEGORY FILTER
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
strCatNames=""
strCatIDs=""
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' ADDED TO IMPLEMENT BATCH UPDATE
'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
Dim intCatDtlId

'******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************


Dim sRslLastCategory

Set cmd = dw.CreateCommandSPwTimeout(cn, "usp_SelectRslSpareKits", 240)
dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 10, PVID
Set rs = dw.ExecuteCommandReturnRS(cmd)

If rs.EOF And rs.BOF Then
    Response.Write "<tr><td colspan=""9"">No Rows Returned</td></tr>"
Else
    Do Until rs.eof
    
        strCurrCatID=rs("CategoryId") ' ADDED TO IMPLEMENT CATEGORY FILTER

        If sRslLastCategory <> rs("CategoryName") Then

            intCatDtlId=-1

            '************************************************************************************************************************************************************************************************************************************************************************
            ' ADDED/ALTERED TO IMPLEMENT CATEGORY FILTER
            '************************************************************************************************************************************************************************************************************************************************************************
		    If blnPopFilterCategory Then
			    ' Accummulate Categories and populate list after this routine is done,  MAY CONSIDER AJAX CALL TO NEW PAGE
                IF(InStr(1, "|" & strCatIDs & "|", "|" & strCurrCatID & "|")=0) THEN
			        If LEN(TRIM(strCatIDs))=0 Then
				        strCatIDs=rs("CategoryId")
				        strCatNames=rs("CategoryName")	
			        Else
				        strCatIDs=strCatIDs & "|" & rs("CategoryId")
				        strCatNames=strCatNames & "|" & rs("CategoryName")	
			        End If
                END IF
		    End If

            IF (blnApplyCatFilter AND InStr(1, strSelCatIDs, "|" & strCurrCatID & "|")>0) OR (NOT blnApplyCatFilter) THEN
            '************************************************************************************************************************************************************************************************************************************************************************
                sRslLastCategory = rs("CategoryName") & ""

                Response.Write "<tr PVID=""" & PVID & """ SFPN=""" & sServiceFamilyPn & """ CID=""" & rs("CategoryID") & """ onmouseover=""rslRow_onmouseover()"" onmouseout=""rslRow_onmouseout()"""
                If (m_IsSysAdmin Or m_IsSpdmUser) and (Not blnIsOSSPUser) Then
                    Response.Write " onclick=""rslCategoryRow_onclick()"""
                    ' ADD CHECKBOX TO SELECT ALL SUBORDINATE DETAIL RECORDS...WOULD PROBABLY BE BETTER IF NO "HEADER" CHECK BOX EXIST WHEN NO APPLICABLE "DETAIL" RECORDS EXIST (WOULD REQUIRE "LOOK AHEAD" FUNCTIONALITY)
                    Response.Write " class=""MatrixAvRow"" style=""height:20px""><td colspan=""9""><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left""><input type=""checkbox"" Name=""CBCatHdr"" id=""CBCatHdr" & rs("CategoryId") & """ onclick=""toggleCBCatHdr(this)"" onmouseover=""getCBCatHdrTitle(this,'" & rs("CategoryName") & "')"" >&nbsp;&nbsp;" & rs("CategoryName") & "</span><span style=""float:right; text-align:right; color:Red; font-weight:bold; font-style:italic; vertical-align:middle;"">Click header to add new spare kit row.</span></td></tr>"
                    'Response.Write " class=""MatrixAvRow"" style=""height:20px""><td colspan=""9""><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left""><input type=""checkbox"" Name=""CBCatHdr"" id=""CBCatHdr" & rs("CategoryId") & """ onclick=""toggleCBCatHdr(this)"" onmouseover=""getCBCatHdrTitle(this,'" & rs("CategoryName") & "')"" >&nbsp;&nbsp;" & rs("CategoryName") & "</span><span style=""float:right; text-align:right; color:Red; font-weight:bold; font-style:italic; vertical-align:middle;"">Click header to add new spare kit row.</span></td></tr>"
                Else
                    ' CONSIDER REMOVING "Click header to add new spare kit row." text as this is not enabled here
                    Response.Write " class=""MatrixAvRow"" style=""height:20px""><td colspan=""9""><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span><span style=""float:right; text-align:right; color:Red; font-weight:bold; font-style:italic; vertical-align:middle;""></span></td></tr>"
                End If
                'DISABLED BY PR - Response.Write " class=""MatrixAvRow"" style=""height:20px""><td colspan=""9""><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span><span style=""float:right; text-align:right; color:Red; font-weight:bold; font-style:italic; vertical-align:middle;"">Click header to add new spare kit row.</span></td></tr>"
            '************************************************************************************************************************************************************************************************************************************************************************
            END IF
            '************************************************************************************************************************************************************************************************************************************************************************
        End If
        
        '************************************************************************************************************************************************************************************************************************************************************************
        ' ADDED/ALTERED TO IMPLEMENT CATEGORY FILTER
        '************************************************************************************************************************************************************************************************************************************************************************
        IF (blnApplyCatFilter AND InStr(1, strSelCatIDs, "|" & strCurrCatID & "|")>0) OR (NOT blnApplyCatFilter) THEN
        '************************************************************************************************************************************************************************************************************************************************************************

            response.Write "<tr PVID=""" & PVID & """ DRID=""" & rs("DeliverableRootID") & """ ID=""SK" & rs("SpareKitId") & """ SKID=""" & rs("SpareKitId") & """ SFPN=""" & sServiceFamilyPn & """ CID=""" & rs("CategoryID") & """ onmouseover=""rslRow_onmouseover()"" onmouseout=""rslRow_onmouseout()"""
            If (m_IsSysAdmin Or m_IsSpdmUser) Or blnIsOSSPUser Then 'and (Not blnIsOSSPUser) Then
                response.Write " onclick=""rslRow_onclick()"""
                ' ADD CHECKBOX TO SELECT INDIVIDUAL DETAIL RECORD (SEE CODE BLOCK BELOW)
            End If
            response.Write " onmousedown=""RslRow_OnMouseDown()"""
            If NOT (CLng(rs("StatusId")) = 3 Or CLng(rs("StatusId")) = 0) Then
                response.Write " style=""background-color:MistyRose;"""
            End If
            response.Write " ><td>"

            '******************************************************************************************************************************************************************
            ' ADDED TO IMPLEMENT BATCH UPDATES
            '******************************************************************************************************************************************************************
            If (m_IsSysAdmin Or m_IsSpdmUser) AND Not IsNull(rs("SpareKitId")) And (Not blnIsOSSPUser)  Then
                ' ADD CHECKBOX TO SELECT INDIVIDUAL DETAIL RECORD
                intCatDtlId=intCatDtlId+1
                Response.Write("<input type=""checkbox"" Name=""CBCatDtl"" id=""CBCatDtl" & rs("CategoryId") & "-" & CStr(intCatDtlId) & """ onclick=""toggleCBCatDtl(this)"" onmouseover=""getCBCatDtlTitle(this)"">&nbsp;&nbsp;")
            End If
            '******************************************************************************************************************************************************************

            response.Write rs("Status") 
            response.Write "</td><td nowrap>"
            response.Write rs("SpareKitNo") 
            response.Write "</td><td>"
            response.Write rs("GPGDescription") 
            response.Write "</td><td style=""text-align:center;"
            If UCase(rs("CrossPlantStatus")&"") = "C6" Or UCase(rs("CrossPlantStatus")&"") = "C1" Then
                response.Write "background-color:mistyrose;"
            End If
            response.Write""">"
            response.Write rs("CrossPlantStatus")&""
            response.Write "</td><td>"
            response.Write rs("ServiceSubassembly") 

	    If Not blnIsOSSPUser Then
            	response.Write "</td><td><a href=""" & AppRoot & "\dmview.asp?ID=" & rs("DeliverableRootID") & """ target=""_blank"">"
            	response.Write rs("DeliverableRootID") 
            	response.Write "</a></td><td>"
	    Else
            	response.Write "</td><td>"
            	response.Write rs("DeliverableRootID") 
            	response.Write "</a></td><td>"
	    End If

            response.Write rs("ExcaliburDescription") 
            response.Write "</td><td>"

            ' Convert carriage return and line feed combinations to HTML equivalents
            if(Not IsNull(rs("Notes"))) then response.Write(replace(rs("Notes"), vbcrlf, "<BR/>"))

            response.Write "</td><td>"

            ' Convert carriage return and line feed combinations to HTML equivalents
            If(Not IsNull(rs("Comments"))) then response.Write(replace(rs("Comments"), vbcrlf, "<BR/>"))
            
            response.Write "</td></tr>"

        '************************************************************************************************************************************************************************************************************************************************************************
        END IF
        '************************************************************************************************************************************************************************************************************************************************************************
        
        rs.MoveNext
    Loop
    rs.close


End If

                    %>
                </tbody>
            </table>
        </div>
        <!-- MAY ADD OTHER WRAPPERS BELOW, BUT FOR NOW... -->
        <% Else 'm_BrandID <> ""%>
        <%

'
' Get Kmat for Current Brand
'
Set cmd = dw.CreateCommandSQL(cn, "SELECT KMAT From Product_Brand (NOLOCK) WHERE ID=" & m_BrandID)
Set rs = dw.ExecuteCommandReturnRS(cmd)
If Not rs.EOF Then
    sKmat = rs("KMAT") & ""
End If
rs.Close

'Response.Write m_BrandId & "<br>"
'Response.Write "'" & sServiceFamilyPn & "'<br>"
'Response.Flush
        %>
        <span id="theSectionTitle" style="font-size: x-small; font-weight: bold;">
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;(" & m_BrandName &  ")&nbsp;-&nbsp;Recommended Spares List AV Mapping" %>
        </span>
        <%If sKmat = "" Then %>
        <table class="MatrixTable">
            <tr>
                <td width="100%" colspan="2">
                    <span style="color: Red; text-align: center; font-weight: bold;">KMAT is not saved in
                        Program Data.<br />
                        Mapping Information will not be available.</span>
                </td>
            </tr>
        </table>
        <%ElseIf Trim(sServiceFamilyPn) = "" Then%>
        <table class="MatrixTable">
            <tr>
                <td width="100%" colspan="2">
                    <span style="color: Red; text-align: center; font-weight: bold;">Service Family Pn is
                        not saved in Program Data.<br />
                        Mapping Information will not be available.</span>
                </td>
            </tr>
        </table>
        <%
Else
    Set cmd = dw.CreateCommandSP(cn, "usp_SelectSpareKitAvMappings")
    dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, sServiceFamilyPn
    dw.CreateParameter cmd, "@p_ProductBrandId", adInteger, adParamInput, 10, m_BrandId
    Set rs = dw.ExecuteCommandReturnRS(cmd)
              '  response.write(m_BrandId & " " & sServiceFamilyPn)
    If rs.EOF And rs.BOF Then
        %>
        <table class="MatrixTable">
            <%
        Response.Write "<tr><td colspan=""7"">No Rows Returned</td></tr>"
            %>
        </table>
        <%
    Else

        '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
        ' ADDED TO IMPLEMENT CATEGORY FILTER
        '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
        strCatNames=""
        strCatIDs=""
        '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

        %>
        <div style="overflow-y: scroll; overflow-x: scroll;" id="DIV1">
            <table class="MatrixTable">
                <%
        dim i
        response.Write "<thead><tr style=""height:20px""><th class=""pinnedRow"">Spare Kit #</th><th class=""pinnedRow"">GPG Description</th>"
        If rs.Fields.Count > 7 Then
            For i = 7 to rs.Fields.Count -1
                response.Write "<th class=""pinnedRow"">" & rs.Fields(i).name & "</th>"
            Next
        End If
        Response.Write "</tr></thead><tbody>"
    
      dim rr
      dim avCount 
        Do Until rs.eof
           'for rr = 1 to rs.Fields.Count - 1
         '  response.Write(rs.Fields.Item(rr))
          ' next
            'response.Write(rs.Fields.Count)
           ' response.Write(rs.Fields.Item(27))
            strCurrCatID=rs("CategoryId") ' ADDED TO IMPLEMENT CATEGORY FILTER

            If sRslLastCategory <> rs("CategoryName") Then
                '************************************************************************************************************************************************************************************************************************************************************************
                ' ADDED/ALTERED TO IMPLEMENT CATEGORY FILTER
                '************************************************************************************************************************************************************************************************************************************************************************
		        If blnPopFilterCategory Then
			        ' Accummulate Categories and populate list after this routine is done,  MAY CONSIDER AJAX CALL TO NEW PAGE
                    IF(InStr(1, "|" & strCatIDs & "|", "|" & strCurrCatID & "|")=0) THEN
			            If LEN(TRIM(strCatIDs))=0 Then
				            strCatIDs=rs("CategoryId")
				            strCatNames=rs("CategoryName")	
			            Else
				            strCatIDs=strCatIDs & "|" & rs("CategoryId")
				            strCatNames=strCatNames & "|" & rs("CategoryName")	
			            End If
                    END IF
		        End If
                '************************************************************************************************************************************************************************************************************************************************************************

                IF (blnApplyCatFilter AND InStr(1, strSelCatIDs, "|" & strCurrCatID & "|")>0) OR (NOT blnApplyCatFilter) THEN
                '************************************************************************************************************************************************************************************************************************************************************************
                    sRslLastCategory = rs("CategoryName") & ""
                     '' RON EDIT - AV / SKU FILTERS
                    ' response.write(rs("SpareKitMappingId"))
                     'response.write(rs("SpareKitNo") & " - sparekitno " & m_BrandId & " - brand id " &  sServiceFamilyPn & " - service fam pn")
              if request.querystring("sk") = "" and  request.querystring("av") = "" then
             ' response.write( sServiceFamilyPn & " - service fam pn")

                    response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span> <a href=""#"" onclick=""deleteSkuAV('" &  m_BrandId & "', '" & sServiceFamilyPn & "' , '" & rs("CategoryName") & "', '', '" & rs("SpareKitNo") & "');""></td></tr>"              
                    'response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span> <a href=""#"" onclick=""alert('" & sServiceFamilyPn & "' );"">.</a> </td></tr>"              

              end if 

            ' if request.querystring("sk") <> "" then
            ' response.write(rs("SpareKitNo") & "<br>")
               ' if ucase(rs("SpareKitNo")) = ucase(request.querystring("sk")) then
                    
                       ' response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span></td></tr>"              
               '         response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span> <a href=""#"" onclick=""deleteSkuAV('" &  m_BrandId & "', '" & sServiceFamilyPn & "' , '" & rs("CategoryName") & "', '" & rs("SpareKitNo") & "');"">Remove AV Mapping(s)</td></tr>"              

               ' end if        
             ' end if

              if request.querystring("av") <> "" then
                    if ucase(currentAV) = ucase(request.querystring("av")) then
                        response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span></td></tr>"              
                    end if        
              end if
              '************************************************************************************************************************************************************************************************************************************************************************
                END IF
                '************************************************************************************************************************************************************************************************************************************************************************
            End If

            '************************************************************************************************************************************************************************************************************************************************************************
            ' ADDED/ALTERED TO IMPLEMENT CATEGORY FILTER
            '************************************************************************************************************************************************************************************************************************************************************************
           'dim noresults as boolean = false
           noSKAVresults = "0"
            IF (blnApplyCatFilter AND InStr(1, strSelCatIDs, "|" & strCurrCatID & "|")>0) OR (NOT blnApplyCatFilter) THEN
            '************************************************************************************************************************************************************************************************************************************************************************
              '' RON EDIT - AV / SKU FILTERS
              if request.querystring("sk") = "" and  request.querystring("av") = "" then           
                response.Write "<tr PVID=""" & PVID & """ PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ MAPID=""" & rs("SpareKitMappingId") & """ SKID=""" & rs("SpareKitId") & """  CID=""" & rs("CategoryID") & """ onmouseover=""rslRow_onmouseover()"" onmouseout=""rslRow_onmouseout()"""
                If (m_IsSysAdmin Or m_IsSpdmUser) Or blnIsOSSPUser Then ' and (Not blnIsOSSPUser) Then
                    Response.Write " onclick=""mapRow_onclick()"""
                End If
                Response.Write " onmousedown=""mapRow_OnMouseDown()"" ><td nowrap><a name='"
                response.Write rs("SpareKitId") & ""
                response.Write "'/>"
                response.Write rs("SpareKitNo") 
                response.Write "</td><td>"
                response.Write rs("Description") 
                response.Write "</td>"
                If rs.Fields.Count > 7 Then
                    For i = 7 to rs.Fields.Count -1
                        Response.Write "<td>" & rs(i) & "" & "</td>"
                    Next
                End If
                response.Write "</tr>"
                noSKAVresults = "0"
            end if
           
            if request.querystring("sk") <> "" then
                if ucase(rs("SpareKitNo")) = ucase(request.querystring("sk")) then
               ' response.write("X")
               if avCount < 1 then
                 response.Write "<tr PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ class=""MatrixAvRow"" style=""height:20px""><td colspan=""" & 2 + rs.Fields.Count - 7 & """><a name=""C" & rs("CategoryId") & """ /><span style=""font: bold x-small verdana; float:left"">" & rs("CategoryName") & "</span> <a href=""#"" onclick=""deleteSkuAV('" &  m_BrandId & "', '" & sServiceFamilyPn & "' , '" & rs("CategoryName") & "', '', '" & rs("SpareKitNo") & "');"">Remove AV Mapping(s)</td></tr>"              
               end if
                    response.Write "<tr PVID=""" & PVID & """ PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ MAPID=""" & rs("SpareKitMappingId") & """ SKID=""" & rs("SpareKitId") & """  CID=""" & rs("CategoryID") & """ onmouseover=""rslRow_onmouseover()"" onmouseout=""rslRow_onmouseout()"""
                    If (m_IsSysAdmin Or m_IsSpdmUser) Or blnIsOSSPUser Then ' and (Not blnIsOSSPUser) Then
                        Response.Write " onclick=""mapRow_onclick()"""
                    End If
                    Response.Write " onmousedown=""mapRow_OnMouseDown()"" ><td nowrap><a name='"
                    response.Write rs("SpareKitId") & ""
                    response.Write "'/>"
                    response.Write rs("SpareKitNo") 
                    response.Write "</td><td>"
                    response.Write rs("Description") 
                    response.Write "</td>"
                    If rs.Fields.Count > 7 Then
                        For i = 7 to rs.Fields.Count -1
                            Response.Write "<td>" & rs(i) & "" & "</td>"
                        Next
                    End If
                    response.Write "</tr>"
                    noSKAVresults = "0"
                    avCount = avCount + 1
                 else
                   ' response.write("Spare Kit / AV not found")
                   
                   noSKAVresults = "1"
                 end if
            end if
            if request.querystring("av") <> "" then
           ' response.write("X")
                if ucase(currentAV) = ucase(request.querystring("av")) then
                'response.write("X")
                    response.Write "<tr PVID=""" & PVID & """ PBID=""" & m_BrandId & """ SFPN=""" & sServiceFamilyPn & """ MAPID=""" & rs("SpareKitMappingId") & """ SKID=""" & rs("SpareKitId") & """  CID=""" & rs("CategoryID") & """ onmouseover=""rslRow_onmouseover()"" onmouseout=""rslRow_onmouseout()"""
                    If (m_IsSysAdmin Or m_IsSpdmUser) Or blnIsOSSPUser Then ' and (Not blnIsOSSPUser) Then
                        Response.Write " onclick=""mapRow_onclick()"""
                    End If
                    Response.Write " onmousedown=""mapRow_OnMouseDown()"" ><td nowrap><a name='"
                    response.Write rs("SpareKitId") & ""
                    response.Write "'/>"
                    response.Write rs("SpareKitNo") 
                    response.Write "</td><td>"
                    response.Write rs("Description") 
                    response.Write "</td>"
                    If rs.Fields.Count > 7 Then
                        For i = 7 to rs.Fields.Count -1
                            Response.Write "<td>" & rs(i) & "" & "</td>"
                        Next
                    End If
                    response.Write "</tr>"
                    noSKAVresults = "0"
                 else
                   ' response.write("Spare Kit / AV not found")
                   
                   noSKAVresults = "1"
                 end if
            end if
            '************************************************************************************************************************************************************************************************************************************************************************
            END IF
            '************************************************************************************************************************************************************************************************************************************************************************
            'next 
            rs.MoveNext
            
        Loop
        rs.close
         
        if  noSKAVresults = "1" then
        'response.Write("<tr><td colspan=24>Spare Kit / AV not found. Select None to reset dispaly<td></tr>")
        end if
                %>
            </table>
        </div>


        <%
    End If 'rs.EOF And rs.BOF

End If 'sKmat = ""
        %>
        <!-- /tbody -->
        <!-- /table -->
        <% End If 'm_BrandId = "" %>
        <% End If 'ServiceReport = "rsl" %>
        <%
If ServiceReport="sku" Then
'##########################################################
'#
'# Draw SKU Sparekit Report
'#
'##########################################################
    '
    ' Retrieve and Display the list of Applicable Spare kits for the specified SKU
    ' DEVELOP SEPARATE FUNCTIONALITY FOR SKU FILTER  (i.e. - variables, storage, functions, etc.)
    '
    '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
    ' ADDED TO IMPLEMENT CATEGORY FILTER ---> MAY CHANGE TO SKU FILTER
    '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
    strCatNames=""
    strCatIDs=""
    '******************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
    Dim sSKLastCategory
    Dim strLastSKU

    Dim strSKUDesc
    Dim strPFVBDesc
    Dim strSKUHeader
    Dim strSKUSKDetail
    Dim intReturnCode
    Dim strReturnDesc
    Dim intNumSKUSKRecs
    Dim blnIncludeAVs: blnIncludeAVs=true
    Dim intIncludeAVs

    Dim dtSPStartTime
    Dim dtSPEndTime
    Dim strDuration

    Dim strBrandSKUList
    strBrandSKUList="<select id='brandSKUs' onchange='gotoSKUHeader(this)'><option value='0'>-- Select SKU --</option>"

    If blnIncludeAVs Then
	    intIncludeAVs=1
    Else
	    intIncludeAVs=0
    End If

    intNumSKUSKRecs=0
    strPFVBDesc=""
    strSKUSKDetail=""

    IF LEN(m_BrandId)>0 THEN

        dtSPStartTime=Now()

        '***************************************************************************************************************************************************************************************
        '   XSL TRANSFORMATIONS - May be able to execute sp once, load into a raw xml stream or string, then apply different templates to the same root data set for table and drop down list
        '***************************************************************************************************************************************************************************************
        Set rs = Server.CreateObject("ADODB.RecordSet")
        Set dw = New DataWrapper
        Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

        Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductSKUs")
        dw.CreateParameter cmd, "@ProductVersionId", adInteger, adParamInput, 10, CSTR(PVID)
        dw.CreateParameter cmd, "@ProductBrandId", adInteger, adParamInput, 10, CSTR(m_BrandId)
        Set rs = dw.ExecuteCommandReturnRS(cmd)

        strBrandSKUList="<select id='brandSKUs' onchange='gotoSKUHeader(this)'><option value='0'>-- Select SKU --</option>"
        
        If NOT rs.EOF Then
            strSKUSKDetail="<table style='font-size: xx-small; font-family: Verdana; ' width='100%'>"

            Do While Not rs.EOF
                strSKUSKDetail = strSKUSKDetail + "<tr style='height:20px; background-color: #EEE8AA; font-size:xx-small;'><td style='border-style: none'>"
                strSKUSKDetail = strSKUSKDetail + "  <table><tr>"
                strSKUSKDetail = strSKUSKDetail + "     <td colspan='9'><a name='" + rs("SKU") + "' />"
                strSKUSKDetail = strSKUSKDetail + "          <span style='font: bold x-small verdana; float:left' id='" + rs("SKU") + "Row'>" + rs("SKU") + " - " + rs("Name") + "</span>"
                strSKUSKDetail = strSKUSKDetail + "     </td>"
                strSKUSKDetail = strSKUSKDetail + "     <td colspan='2' style='text-decoration: underline; font-family: Verdana; font-size: xx-small; font-weight: normal; color: #0000FF; cursor:pointer;'>"
                strSKUSKDetail = strSKUSKDetail + "          <a title='Click to Perform a QuickSearch on this SKU' onclick='javascript:QuickSearch(""" + rs("SKU") + """)'>Show SKU BOM</a>"
                strSKUSKDetail = strSKUSKDetail + "     </td>"
                strSKUSKDetail = strSKUSKDetail + "     <td width='1px'></td><td colspan='2' style='text-decoration: underline; font-family: Verdana; font-size: xx-small; font-weight: normal; color: #0000FF; cursor:pointer;'>"
                strSKUSKDetail = strSKUSKDetail + "          <a title='Click to display the Spare Kits for this SKU.' mode='0' id='" + rs("SKU") + "' onclick='javascript:showSKUSpareKits(this);'>Show Spare Kits</a>"
                strSKUSKDetail = strSKUSKDetail + "     </td>"
                strSKUSKDetail = strSKUSKDetail + "  </tr></table>"
                strSKUSKDetail = strSKUSKDetail + "</td></tr>"
                strSKUSKDetail = strSKUSKDetail + "<tr><td id='[" + rs("SKU") + "]' /></tr>"

                strBrandSKUList=strBrandSKUList+"<option value='" + rs("SKU") + "'>" + rs("SKU") + "</option>"
                rs.MoveNext()
            Loop

            strSKUSKDetail=strSKUSKDetail+"</table>"
        End If

        strBrandSKUList=strBrandSKUList+"</select>"
        rs.Close  
            
        dtSPEndTime=Now()
        strDuration="EXECUTION TIME: " & CSTR(DateDiff("s",dtSPStartTime,dtSPEndTime)) & " second(s)"
        intReturnCode=0        

    ELSE
        intReturnCode=-10
        strReturnDesc="Click a Product Brand to display its SKUs."
    END IF

        %>
        <br />
        <font face="verdana" size="1"><b>Tools:&nbsp;</b></font> <span style="font-size: xx-small;
            font-family: Verdana"><a href="#" onclick="QuickSearch();">QuickSearch</a> | <a href="#"
                onclick="AdvancedSearch();">Advanced Search</a>
            <% If(Not blnIsOSSPUser) Then %>
            | <a href="#" onclick="exportSpb('<%=sServiceFamilyPn %>');">Export SPB</a> | <a
                href="#" onclick="SetServiceFamilyPn(<%=PVID%>);">Edit Service Family Details</a>
            <% End If %>

        </span>
        <br />
        <br />
        <% ' NEED TO ADD/ALTER CODE BELOW SO AS TO DISPLAY FULL PRODUCT BRAND DESC/NAME
IF ServiceReport="sku" THEN
    If m_BrandId = "" Then %>
        <span id="theSectionTitle" style="font-size: x-small; font-weight: bold;">
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;-&nbsp;BTO SKUS" %>
        </span>
        <%Else %>
        <span id="BTOSKUListSpan" style="font-size: x-small; font-weight: bold;">
            <% sProgramName = Left(sProductName, InStr(sProductName, ".")) & "x" %>
            <%= sProgramName & "&nbsp;(" & m_BrandName &  ")&nbsp;-&nbsp;BTO SKUS" %>
            &nbsp;<%=strBrandSKUList %>
        </span>
        <%End If 
ELSEIF LEN(strPFVBDesc)>0 THEN
        %>
        <span id="theSectionTitle" style="font-size: x-small; font-weight: bold;">
            <%= strPFVBDesc & "&nbsp;-&nbsp;BTO SKUS" %>
        </span>
        <%
END IF
        %>
        <div id="processingStatus">
        </div>
        <hr />
        <div style="overflow-y: scroll; overflow-x: scroll;" id="DIV1">
            <% 
    If intReturnCode=0 Then
        Response.Write(strSKUSKDetail)
    ElseIf intReturnCode=-10 Then
        Response.Write(strReturnDesc)
	End If
            %>
        </div>
        <div id="SKUStatus">
            <%
If intReturnCode=0 Then
    If blnApplyCatFilter Then
	   'Response.Write(strReturnDesc & "&nbsp;&nbsp;" & CStr(intNumSKUSKRecs) & " filtered record(s).")
    Else
       'Response.Write(strReturnDesc & "&nbsp;&nbsp;" & CStr(intNumSKUSKRecs) & " total record(s).")
    End If
ElseIf intReturnCode<>-10 Then
	Response.Write(strReturnDesc)
End If

Response.Write("<br/>" & strDuration)
            %>
        </div>
        <%

'***********************************************************************************************************************************************************************
End If
        %>
        <input type="hidden" id="txtClass" name="txtClass" value="<%=sClass%>" />
        <input type="hidden" id="txtID" name="txtID" value="<%=PVID%>" />
        <input type="hidden" id="txtFavs" name="txtFavs" value="<%=sFavs%>" />
        <input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=sFavCount%>" />
        <input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>" />
        <input type="hidden" id="hidLastPublishDt" value="<%=m_LastPublishDt %>" />
        <input type="hidden" id="hidBusinessID" value="<%=m_BusinessID %>" />
        <input type="hidden" id="hidAvCount" value="<%= iAvCount %>" />
        <input type="hidden" id="hidAnchor" value="<%= LinkAnchor %>" />
        <input type="hidden" id="hidKmat" value ="<%=sKmat %>" />
        <!--
--
--  Popups
--
-->
        <div id="PopUpMenu" class="hidden">
            <ul id="menu">
                <% If Not blnIsOSSPUser Then %>
                <li class="default"><a href="#" onclick="parent.location.href='javascript:MenuProperties();'">
                    Properties</a></li>
                <li id="add"><a href="#" onclick="parent.location.href='javascript:MenuAdd();'">Add
                    Kit</a></li>
                <li id="Li1"><a href="#" onclick="parent.location.href='javascript:MenuShowSaProperties();'">
                    Edit SA Numbers</a></li>
                <li id="spacer">
                    <hr width="95%">
                </li>
                <li id="delete"><a href="#" onclick="parent.location.href='javascript:MenuDelete();'">
                    Delete This Record</a></li>
                <% End If %>
            </ul>
        </div>
        <div id="MapContextMenu" class="hidden">
            <ul id="menu">
                <% If Not blnIsOSSPUser Then %>
                <li class="default"><a href="#" onclick="parent.location.href='javascript:MenuShowMapRowProperties();'">
                    Properties</a></li>
                <li id="add"><a href="#" onclick="parent.location.href='javascript:MenuAddMapRow();'">
                    Add Mapping</a></li>
                <% End If %>
            </ul>
        </div>
        <div class="sample_popup" id="popup" style="visibility: hidden; display: none;">
            <div class="menu_form_header" id="popup_drag">
                <img class="menu_form_exit" id="popup_exit" src="<%=AppRoot %>/images/form_exit.png" />
                <% If ServiceReport = "spb" Then %>
                &nbsp;&nbsp;&nbsp;Service Program Bom Options
                <% Else %>
                &nbsp;&nbsp;&nbsp;RSL Options
                <% End If 'ServiceReport = "spb" %>
            </div>
            <div class="menu_form_body">
                <form method="post" action="" id="popup_form">
                <input type="hidden" name="ProductBrandID" value="<%= m_BrandID %>" />
                <input type="hidden" name="ServiceFamilyPn" value="<%= sServiceFamilyPn %>" />
                <input type="hidden" name="ProductVersionId" value="<%= PVID %>" />
                <table>
                    <tr id="trCompare">
                        <th style="white-space: nowrap">
                            Comparison:
                        </th>
                        <td>
                            <div id="selCompareDtDiv">
                                <select class="form" name="selCompareDt" id="selCompareDt">
                                    <option></option>
                                </select>
                            </div>
                        </td>
                    </tr>
                    <tr <% If Not ((m_IsSpdmUser)) Then Response.Write "style=""display: none""" %>>
                        <th>
                            Publish:
                        </th>
                        <td>
                            <input type="checkbox" name="chkPublish" title="Publish" />
                        </td>
                    </tr>
                    <tr>
                        <th>
                            New&nbsp;Matrix:
                        </th>
                        <td>
                            <input type="checkbox" id="chkNewMatrix" name="chkNewMatrix" title="Publish" />
                        </td>
                    </tr>
                    <tr>
                        <th>
                            &nbsp;
                        </th>
                        <td>
                            <input class="btn" type="submit" id="popup_submit" value="Export" onclick="popup_exit('popup');" />
                        </td>
                    </tr>
                </table>
                </form>
            </div>
        </div>
        <div class="ui-helper-hidden">
            <div id="PublishDialog">
            </div>
            <div id="QuickSearch" title="Part Number Quick Search">
                <table>
                    <tr>
                        <td nowrap valign="top">
                            &nbsp;&nbsp;&nbsp;<font face="Verdana" size="2"><strong>Part Number:</strong></font>
                        </td>
                        <td>
                            <input id="txtQuickSearch" />
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <input type="hidden" id="ReportType" value="<%=ServiceReport %>" />
        <%
    '********************************************************************************************************************************************************
    ' ADDED TO IMPLEMENT CATEGORY FILTER
    '********************************************************************************************************************************************************
    If blnPopFilterCategory Then
        %>

        <script type="text/javascript" language="javascript">
<%
	    If ((strSelCatIDs<>"null") And (blnApplyCatFilter)) Then
        	Response.Write(vbTab & vbTab & "var bCategoryFilter=true;" & vbcrlf)
	    	Response.Write(vbTab & vbTab & "populateFilterList('" & strCatIDs & "', '" & strCatNames & "', '" & strSelCatIDs & "');" & vbcrlf)
    	   Else
	        Response.Write(vbTab & vbTab & "var bCategoryFilter=false;" & vbcrlf)
        	Response.Write(vbTab & vbTab & "populateFilterList('" & strCatIDs & "', '" & strCatNames & "', null);" & vbcrlf)
    	   End If
%>
        </script>

        <%
    Else
        %>

        <script type="text/javascript" language="javascript">
<%
	        Response.Write(vbTab & vbTab & "var bCategoryFilter=false;" & vbcrlf)
%>
        </script>

        <%
    End If
    '********************************************************************************************************************************************************
        %>
        <%
    If ServiceReport="sku" Then
        If intReturnCode=0 Then
        %>

        <script type="text/javascript" language="javascript">            var oSKUList = document.getElementById("brandSKUs"); var iTotalSKUs = oSKUList.options.length - 1; var oSKUStatus = document.getElementById("SKUStatus"); oSKUStatus.innerHTML = iTotalSKUs.toString() + " total record(s).<br/>" + oSKUStatus.innerHTML;</script>

        <%
        End If
    End If
        %>
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
	'Response.Cookies("LastProductDisplayed") = PVID

function FormatPeriodName(strName, strPeriod)
	dim DatePartArray
	dim strYear
	dim strDate
	dim strDOW
	dim MonthArray
	
	MonthArray = split(",Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")
	
	DatePartArray= split(strName,",")
	 
	if (trim(DatePartArray(0)) = "" and not isnumeric(DatePartArray(0)))or (trim(DatePartArray(1)) = "" and not isnumeric(DatePartArray(1))) then
		FormatPeriodName = strName
	elseif trim(strPeriod) = "1" then
		FormatPeriodName = MonthArray(DatePartArray(0)) & "&nbsp;" & right(DatePartArray(1),2)
	elseif trim(strPeriod) = "2" then
		strDate = DateAdd("ww",DatePartArray(0)-1,"1/1/" & DatePartArray(1))
		strDOW = weekday(strDate)-1
		strDate = dateadd("d",-strDOW,strDate)
		FormatPeriodName = Day(strDate) & "&nbsp;" & MonthArray(MOnth(strDate))
	elseif trim(strPeriod) = "3" then
		FormatPeriodName = DatePartArray(0) & "Q" & right(DatePartArray(1),2)
	else
		FormatPeriodName = strName
	end if
end function

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
%>
<%@ Language=VBScript %>
<!-- #include file = "./includes/Security.asp" -->
<!-- #include file="./Agency/AgencyPivot.asp" -->
<!-- #include file="./includes/cookies.inc" -->
<%
    Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
    Response.CodePage = 65001
    Response.Charset="UTF-8"
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim PVID : PVID = regEx.Replace(Request.QueryString("ID"), "")

    regEx.Pattern = "[^0-9a-zA-Z_ ]"
    Dim sList : sList = regEx.Replace(Request("List"), "")
    
    Dim sClass
    Dim sSeries : sSeries = regEx.Replace(Request("Series"), "")
    Dim sGroupBy : sGroupBy = regEx.Replace(Request("GroupBy"), "")
    Dim sStatus : sStatus = regEx.Replace(Request("Status"), "")
    Dim sInterval : sInterval = regEx.Replace(Request("Interval"), "")
    Dim sRegion : sRegion = regEx.Replace(Request("Region"), "")
    Dim AppRoot : AppRoot = Session("ApplicationRoot")
    Dim NumReleases : NumReleases = 0
    Dim reportCategory : reportCategory = "Certification"
    if(Request("Class") = "") then
        sClass = "1"
    else
        sClass = regEx.Replace(Request("Class"), "")
    end if

    Dim securityObj
    Set securityObj = New ExcaliburSecurity

    	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if

	end function

	function FormatSystemID(strValue)
		dim RowArray
		dim Row
		dim strOutput
		dim IDArray

		if instr(strValue,"^")=0 and instr(strValue,"|")=0 then
			FormatSystemID = strValue
		else
			strOutput = ""
			RowArray = split(strValue,"|")
			for each Row in Rowarray
				if instr(Row,"^")=0 then
					strOutput = strOutput & ", " & Row
				else
					IDArray = split(Row,"^")
					strOutput = strOutput & ", " & IDArray(0)
					if Ubound(IDArray) > 0 then
						if trim(IDArray(1)) <> "" and trim(IDArray(1)) <> "&nbsp;" then
							strOutput = strOutput & "&nbsp;(" & replace(IDArray(1)," ","&nbsp;") & ")"
						end if
					end if
				end if
			next
			if strOutput = "" then
				FormatSystemID = "&nbsp;"
			else
				FormatSystemID = mid(strOutput,3)
			end if

		end if
	end function

	function GetKeyValue(strString, strKey)
		dim strPair
		dim PairArray
		dim strValue
		dim KeyArray
		PairArray = split(strString,"&")
		strValue = ""
		for each strPair in PairArray
			if instr(strPair,"=") > 0 then
				KeyArray = Split(strPair,"=")
				if lcase(trim(KeyArray(0))) = lcase(trim(strKey)) then
					strValue = KeyArray(1)
					exit for
				end if
			end if
		next
		GetKeyValue = strValue

	end function

    function SetKeyValue(strString, strKey, strOldValue, strNewValue)
        dim strDefault
		if (strOldValue <> "") And (strNewValue <> "") then
            strDefault = Replace(strString, strKey & "=" & strOldValue, strKey & "=" & strNewValue)
        else 
            if strString <> "" then
                strDefault = strKey & "=" & strNewValue & "&" & strString
            elseif strNewValue <> "" then
                strDefault = strKey & "=" & strNewValue
            else 
                strDefault = strKey & "=SW"
            end if
        end if
		SetKeyValue = strDefault
	end function

	function AddURLParameter(strURL,strKey,strValue)

		AddURLParameter = strURL
		dim PairArray
		dim strPair
		dim blnFound
		dim ValueArray
		dim strbuffer

		PairArray = split(strURL,"&")
		blnFound = false
		for Each strPair in PairArray
			strbuffer = ""
			if lcase(left(strPair, len(strKEY))) = lcase(strKey) then
				blnFound = true
				if trim(strValue) = "" then
					strbuffer = replace(strURL,"&" & strPair,"")
					strbuffer = replace(strbuffer,strPair & "&","")
					strbuffer = replace(strbuffer,strPair,"")
					AddURLParameter = strbuffer
				else
					AddURLParameter = replace(strURL,strPair,strKey & "=" & strValue)
				end if
				exit for
			end if
		next

		if (Not blnFound) and trim(strValue) <> "" then
			if len(strURL) = 0 then
				AddURLParameter = strKey & "=" & strValue
			else
				AddURLParameter = strURL & "&" & strKey & "=" & strValue
			end if
		end if

	end function


	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
%>
<!DOCTYPE html>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
<head>
    <!--meta http-equiv="X-UA-Compatible" content="IE=8" /-->

<title>Excalibur</title>
<meta name="VI60_DefaultClientScript" content="JavaScript" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<link href="style/wizard style.css" type="text/css" rel="stylesheet" />
<link href="style/Excalibur.css" type="text/css" rel="stylesheet" />
<link href="style/bubble.css" type="text/css" rel="stylesheet" />
<!-- #include file="includes/bundleConfig.inc" -->
<script type="text/javascript" src="includes/client/json2.js"></script>
<script type="text/javascript" src="includes/client/json_parse.js"></script>
<script type="text/javascript" language="javascript" src="_ScriptLibrary/jsrsClient.js"></script>
<script type="text/javascript" language="javascript" src="_ScriptLibrary/sort.js"></script>
<script src="/Pulsar/Scripts/spin/spin.js" type="text/javascript"></script>
<script src="/Pulsar/Scripts/spin/jquery.spin.js" type="text/javascript"></script>
<script src="/pulsar2/js/userfavorite.js" type="text/javascript"></script>
    
<script type="text/javascript">
    $(function () {
        jQuery(window).scroll(function () {
            jQuery('#modal_dialog').dialog('option', 'position', 'center');
        });

        $("#iframeDialog").dialog({
            modal: true,
            autoOpen: false,
            width: 900,
            height: 800,
            close: function () {
                $("#modalDialog").attr("src", "about:blank");
            },
            resizable: false

        });
    });

    function ShowIframeDialog() {
        $("#iframeDialog iframe").attr("width", $("#iframeDialog").dialog("option", "width") - 50);
        $("#iframeDialog iframe").attr("height", $("#iframeDialog").dialog("option", "height") - 50);
        $("#iframeDialog iframe").attr("src", "../Agency/Agency.asp?ID=5000");
        $("#iframeDialog").dialog("option", "title", "Agency Status");
        $("#iframeDialog").dialog("open");
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

    function ShowMarketingNameDialog(BrandID, ExistingName, NameType, ProductBrandID, Series) {
        var DlgWidth = 700;
        var DlgHeight = 350;
        var DialogName = ""
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        $("#divOpenMarketingNameUpdate").dialog({ width: DlgWidth, height: DlgHeight });
        $("#ifOpenMarketingNameUpdate").attr("width", "98%");
        $("#ifOpenMarketingNameUpdate").attr("height", "92%");
        $("#ifOpenMarketingNameUpdate").attr("src", '<%=AppRoot %>/UpdateMarketingNames.asp?BID=' + BrandID + '&PBID=' + ProductBrandID + '&Name=' + ExistingName + '&Type=' + NameType + '&Series=' + Series);
        /* PER EFREN'S REQUEST - DO NOT REMOVE
        if (NameType == 4 && ExistingName == "") {
            DialogName = "Add BTO Service Tag Name";
        }
        if (NameType == 4 && ExistingName != "") {
            DialogName = "Edit BTO Service Tag Name";
        }*/

        if (NameType == 5 && ExistingName == "") {
            DialogName = "Add CTO Model Number";
        }
        if (NameType == 5 && ExistingName != "") {
            DialogName = "Edit CTO Model Number";
        }

        if (NameType == 6 && ExistingName == "") {
            DialogName = "Add Short Name";
        }
        if (NameType == 6 && ExistingName != "") {
            DialogName = "Edit Short Name";
        }

        if (NameType == 7 && ExistingName == "") {
            DialogName = "Add HP Brand Name (Service Tag up)";
        }
        if (NameType == 7 && ExistingName != "") {
            DialogName = "Edit HP Brand Name (Service Tag up)";
        }

        if (NameType == 8 && ExistingName == "") {
            DialogName = "Add Model Number (Service Tag down)";
        }
        if (NameType == 8 && ExistingName != "") {
            DialogName = "Edit Model Number (Service Tag down)";
        }

        if (NameType == 9 && ExistingName == "") {
            DialogName = "Add BIOS Branding";
        }
        if (NameType == 9 && ExistingName != "") {
            DialogName = "Edit BIOS Branding";
        }

        $("#divOpenMarketingNameUpdate").dialog("option", "title", DialogName);
        $("#divOpenMarketingNameUpdate").dialog("open");
    }

    function CloseMarketingNameDialog() {
        $("#divOpenMarketingNameUpdate").dialog("close");
    }

    function addZero(i) {
        if (i < 10) {
            i = "0" + i;
        }
        return i;
    }

    function GetTime() {
        var nowDate = new Date();
        var status = nowDate.getHours() < 12 ? "AM" : "PM";
        var dateTime = nowDate.getMonth() + 1 + "/" + nowDate.getDate() + "/" + nowDate.getFullYear() + " " + (nowDate.getHours() < 12 ? addZero(nowDate.getHours()) : nowDate.getHours() - 12) + ":" + addZero(nowDate.getMinutes()) + ":" + addZero(nowDate.getSeconds()) + " " + status;
        return dateTime;
    }

    function ClosePropertiesDialog(strID) {
        $("#iframeDialog").dialog("close");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function ClosePropertiesDialog_fromClone(strID) {

        $("#iframeDialog").dialog("close");
        if (typeof (strID) != "undefined") {
            if (strID == txtID.value) {
                window.location.reload(true);
            } else {
                window.location = "/Excalibur/pmview.asp?ID=" + strID + "&Class=" + txtClass.value;
            }
        }
    }

    function CloseIframeDialog() {
        $("#iframeDialog").dialog("close");
    }

    function adjustWidth(percent) {
        return document.documentElement.offsetWidth * (percent / 100);
    }

    function adjustHeight(percent) {
        return (document.documentElement.offsetHeight * (percent / 100));
    }
</script>

<script type="text/javascript">
<!--
    /*
    *  Top Link Scripts
    */

    function ShowProperties(DisplayedID, Clone, FusionRequirements) {
        var shouldClone;

        if (Clone == 1) {
            shouldClone = "&Clone=1";
        } else {
            shouldClone = "";
        }

        ShowPropertiesDialog("mobilese/today/programs.asp?HWPM=0&ID=" + DisplayedID + shouldClone + "&Pulsar=" + FusionRequirements, "Product Properties", 1200, 1000);
    }

    function ShowCommodityProperties(DisplayedID, Type, Clone) {
        var strID;
        var shouldClone;

        if (Clone == 1) {
            shouldClone = "&Clone=1";
        } else {
            shouldClone = "";
        }

        if (Type == 1)
            ShowPropertiesDialog("mobilese/today/programs.asp?Commodity=1&ID=" + DisplayedID + shouldClone, "Product Properties", 700, 550);
        else if (Type == 2)
            ShowPropertiesDialog("mobilese/today/programs.asp?FactoryEngineer=1&ID=" + DisplayedID + shouldClone, "Product Properties", 600, 450);
        else if (Type == 3)
            ShowPropertiesDialog("mobilese/today/programs.asp?Accessory=1&ID=" + DisplayedID + shouldClone, "Product Properties", 600, 450);
        else if (Type == 4)
            ShowPropertiesDialog("mobilese/today/programs.asp?HWPM=1&ID=" + DisplayedID + shouldClone, "Product Properties", 600, 450);

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
            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1");
            RFLink.style.display = "none";
            AFLink.style.display = "";
        }
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
            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1&ID=" + AddingID);
            RFLink.style.display = "";
            AFLink.style.display = "none";
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

    function SIAssignments(PVID) {
        modalDialog.open({ dialogTitle: 'SI Assignments', dialogURL: '<%=AppRoot %>/SIAssignmentsFrame.asp?PVID=' + PVID + '', dialogHeight: (500), dialogWidth: (GetWindowSize('width') + 100), dialogResizable: true, dialogDraggable: true });
    }

//-->
</script>
<script type="text/javascript" language="javascript">
<!--
    var oPopup = window.createPopup();
    var SelectedRow;
    var gstrSoftpaqPath = "";
    var gstrCVAPath = "";
    var gstrTextFilePath = "";

    //
    // Common Functions
    //
    function UpdateUserAccess() {
        window.location.href = "UpdateUserAccess.asp";
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

    function SendDelEmail(VersionID) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("query/DelVerDetailSendEmail.asp?ID=" + VersionID);
    }

    function contextMenu(RootID, TypeID) {
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;

        if (window.event.srcElement.className == "text") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                    SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
        else if (window.event.srcElement.className == "cell") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement)
                    SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }

        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionPrint(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Print&nbsp;Preview...</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionMail(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email...</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionProperties(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionHistory(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;History...</SPAN></FONT></DIV>";
        popupBody = popupBody + "</DIV>";

        oPopup.document.body.innerHTML = popupBody;

        oPopup.show(lefter, topper, 130, 115, document.body);
    }

    //
    // Localization Tab
    //

    function countryrows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }


    function AddChange(ID) {        
        window.open("/pulsarplus/product/product/GetTodayAction/0/3/"+ID+"/0/Excalibur", "blank", "scrollbars=yes,toolbar=no,menubar=no,resizable=yes,width=1200,height=800")
    }

    function changerows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function changerows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";

    }

    function ActionProperties(strID, strType) {
        var strResult;
        modalDialog.open({ dialogTitle: 'Action Properties', dialogURL: 'mobilese/today/action.asp?ID=' + strID + '&Type=' + strType + '', dialogHeight: 800, dialogWidth: 900, dialogResizable: true, dialogDraggable: true });
    }

    function ActionPrint(strID, strType) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("mobilese/today/actionReport.asp?Action=0&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=yes,status=no")
    }

    function DcrPddExport(ID) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("actions/pdd_export.asp?ID=" + ID + "&PDDExport=True", null, "left=" + NewLeft + ",top=" + NewTop + ",width=655,height=650,resizable=yes,menubar=no,toolbar=no,status=no")
    }

    function ActionMail(strID, strType) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("mobilese/today/actionReport.asp?Action=1&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No")
    }

    function ActionHistory(strID, strType) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("mobilese/today/actionHistory.asp?Action=1&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No")
    }

    function changerows_onclick() {
        var strID;
        var strDisplay;


        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
			var params=strDisplay.split("&");
			var issueId=params[0].split("=");
			var typeId=params[1].split("=");			
            //modalDialog.open({ dialogTitle: 'Change Properties', dialogURL: 'mobilese/today/action.asp?' + strDisplay, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });			
			window.open("/pulsarplus/product/product/TodayAction/"+issueId[1]+"/"+typeId[1]+"/0/0/Excalibur", "blank", "scrollbars=yes,toolbar=no,menubar=no,resizable=yes,width=1200,height=800")
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
			var params=strDisplay.split("&");
			var issueId=params[0].split("=");
			var typeId=params[1].split("=");			
            //modalDialog.open({ dialogTitle: 'Change Properties', dialogURL: 'mobilese/today/action.asp?' + strDisplay, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
            window.open("/pulsarplus/product/product/TodayAction/"+issueId[1]+"/"+typeId[1]+"/0/0/Excalibur", "blank", "scrollbars=yes,toolbar=no,menubar=no,resizable=yes,width=1200,height=800")
        }

    }
    //
    // Agency Tab
    //
    function AgencyPddExport(ID) {

        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        strResult = window.open("agency/pdd_export.asp?ID=" + ID + "&PDDExport=True", "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,resizable=yes,menubar=no,toolbar=no,status=no")
    }

    function OTSrows_onclick(strID) {
        var strID;
        var strDisplay;

        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            ShowOTSDetails(strID);
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            ShowOTSDetails(strID);
        }
    }

    //
    // Action Items Tab
    //
    function actionrows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function actionrows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";

    }

    function actionrows_onclick() {
        var strID;
        var strDisplay;


        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            modalDialog.open({ dialogTitle: 'Action Properties', dialogURL: 'mobilese/today/action.asp?' + strDisplay + '', dialogHeight: 800, dialogWidth: 900, dialogResizable: true, dialogDraggable: true });

        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            modalDialog.open({ dialogTitle: 'Action Properties', dialogURL: 'mobilese/today/action.asp?' + strDisplay + '', dialogHeight: 800, dialogWidth: 900, dialogResizable: true, dialogDraggable: true });
        }
    }

    //
    // Issues Tab
    //
    function issuerows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function issuerows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";

    }

    function issuerows_onclick() {
        var strID;
        var strDisplay;


        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            ShowPropertiesDialog("mobilese/today/action.asp?" + strDisplay, "Issue Properties", 800, 900);
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            ShowPropertiesDialog("mobilese/today/action.asp?" + strDisplay, "Issue Properties", 800, 900);
        }
    }

    //
    // Status Tab
    //
    function statusrows_onclick() {
        var strID;
        var strDisplay;


        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            ShowPropertiesDialog("mobilese/today/action.asp?" + strDisplay, "Status Properties", 800, 900);
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            ShowPropertiesDialog("mobilese/today/action.asp?" + strDisplay, "Status Properties", 800, 900);
        }
    }

    //
    // Schedule Tab
    //
    function scheduleLink_onClick(ScheduleID) {
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ScheduleID=" + ScheduleID
    }

    function schedulerows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function brandLink_onClick(ProductBrandID, strStep) {
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ProductBrandID=" + ProductBrandID + "&List=" + strStep;
    }

    function releaseLink_onClick(ProductRelease, strStep) {
        var ProductBrandID;
        ProductBrandID = txtProductBrandID.value;
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ProductBrandID=" + ProductBrandID + "&ProductRelease=" + ProductRelease + "&List=" + strStep;
    }

    function deliverableReleaseLink_onClick(ProductRelease, strStep) {
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ProductRelease=" + ProductRelease + "&List=" + strStep;
    }

    function imageReleaseLink_onClick(ProductReleaseID, tabName) {
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ImageTool=<%=request("ImageTool")%>&ImageActiveType=<%=request("ImageActiveType")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>&ProductReleaseID=" + ProductReleaseID + "&List=" + tabName;
    }

    function imageOSReleaseLink_onClick(ProductOSReleaseID, tabName) {
        window.location.href = "pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ImageTool=<%=request("ImageTool")%>&ImageActiveType=<%=request("ImageActiveType")%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=" + ProductOSReleaseID + "&List=" + tabName;
    }

    function schedulerows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";
    }

    function requirementrows_onmouseover(strID) {
        window.document.getElementById(strID).style.color = "red";
        window.document.getElementById(strID).style.cursor = "hand";
        return;

        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function requirementrows_onmouseout(strID) {
        window.document.getElementById(strID).style.color = "black";
        return;

        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";

    }

    function schedulerows_onclick(isPulsarProduct, ScheduleID) {
        var strID;
        var strDisplay;
        var sAction = 'Edit';

        if (UserHasPermission('System.Admin') !== true && UserHasPermission('Schedule.Edit') !== true) {
            sAction = 'View';
        }

        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            if (strDisplay.indexOf("ScheduleID") < 0) {
                strDisplay += "&ScheduleID=" + ScheduleID;
            }
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            if (isPulsarProduct == "1")
                ShowPropertiesDialog("schedule/schedule_Pulsar.asp?" + strDisplay + "&action=" + sAction, "Schedule Properties", 800, 700);
            else
                ShowPropertiesDialog("schedule/schedule.asp?" + strDisplay + "&action=" + sAction, "Schedule Properties", 800, 700);
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            if (strDisplay.indexOf("ScheduleID") < 0) {
                strDisplay += "&ScheduleID=" + ScheduleID;
            }
            window.event.srcElement.parentElement.style.color = "black";
            if (isPulsarProduct == "1")
                ShowPropertiesDialog("schedule/schedule_Pulsar.asp?" + strDisplay + "&action=" + sAction, "Schedule Properties", 800, 700);
            else
                ShowPropertiesDialog("schedule/schedule.asp?" + strDisplay + "&action=" + sAction, "Schedule Properties", 800, 700);
        }
    }

    function countryrows_onclick() {
        var strID;
        var strDisplay;
        var sFusionRequirements = document.getElementById("inpFusionRequirements").value;
        var ProductVersionID = $("#txtID").val();
        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            ShowPropertiesDialog("countries/localizations.asp?" + strDisplay + "&FusionRequirements=" + sFusionRequirements + "&pvID=" + ProductVersionID, "Select Localizations", GetWindowSize('width'), GetWindowSize('height'));

        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            ShowPropertiesDialog("countries/localizations.asp?" + strDisplay + "&FusionRequirements=" + sFusionRequirements + "&pvID=" + ProductVersionID, "Select Localizations", GetWindowSize('width'), GetWindowSize('height'));

        }
    }

    function requirementrows_onclick(strID) {

        var strResults = "";
        var strLeft = (screen.width - 720) / 2;
        var strTop = (screen.height - 680) / 2;
        var RowID = String(strID).substr(3, String(strID).indexOf("&") - 3);
        var DelCell;
        if (typeof (window.event.srcElement.href) != "undefined")
            return;

        modalDialog.open({ dialogTitle: 'Update Requirement', dialogURL: 'requirements/requirement.asp?' + strID + '', dialogHeight: 840, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });

        //save Product and Version ID for results function: ---
        globalVariable.save(RowID, 'row_id');
        globalVariable.save(strID, 'requirement_id');
    }

    function requirementrowsResults() {
        strResults = modalDialog.getArgument('requirement_save_array');
        strResults = JSON.parse(strResults);

        var iRowID = globalVariable.get('row_id');

        if (typeof (strResults) != "undefined") {
            if (strResults[0].replace(/ /g, "").replace(/&nbsp;/g, "") == "" || strResults[0].replace(/ /g, "").replace(/&nbsp;/g, "") == String.fromCharCode(160))
                document.all("SpecCell" + iRowID).innerHTML = "See PDD for requirements";
            else
                document.all("SpecCell" + iRowID).innerHTML = strResults[0]; //"See PDD for requirements"; //

            DelCell = document.getElementById("DellCell" + iRowID)
            if (DelCell != null)
                document.all("DellCell" + iRowID).innerHTML = strResults[1];
        }
        return;
    }

    function imagerows_onmouseover() {
        if (window.event.srcElement.className == "cell") {
            window.event.srcElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.style.cursor = "hand";
        }
        else if (window.event.srcElement.className == "text") {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function imagerows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";

    }

    function imagerows_onclick() {
        var strID;
        var strDisplay;


        if (window.event.srcElement.className == "text") {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            window.event.srcElement.parentElement.parentElement.style.color = "black";
            strID = window.showModalDialog("image/image.asp?ProdID=" + txtID.value + "&" + strDisplay, "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        }
        else if (window.event.srcElement.className == "cell") {
            strDisplay = window.event.srcElement.parentElement.className;
            window.event.srcElement.parentElement.style.color = "black";
            strID = window.showModalDialog("image/image.asp?ProdID=" + txtID.value + "&" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        }
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function DisplayImagePulsar(ImageID) {
        modalDialog.open({ dialogTitle: 'Image Properties', dialogURL: 'image/fusion/Image_Pulsar.asp?ProdID=' + txtID.value + '&ID=' + ImageID + '', dialogHeight: 700, dialogWidth: 905, dialogResizable: true, dialogDraggable: true });
    }

    function DisplayImageFusion(ImageID) {
        modalDialog.open({ dialogTitle: 'Image Properties', dialogURL: 'image/fusion/image.asp?ProdID=' + txtID.value + '&ID=' + ImageID + '', dialogHeight: 700, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });
    }

    function DisplayImage(ImageID) {
        modalDialog.open({ dialogTitle: 'Image Properties', dialogURL: 'image/image.asp?ProdID=' + txtID.value + '&ID=' + ImageID + '', dialogHeight: 700, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });
    }

    function DisplaySingleImage(ImageID) {
        var strID;
        var NewTop = (window.screen.Height - (500 + 200)) / 2;
        var NewLeft = (window.screen.Width - 600) / 2;
        strID = window.open("image/localization.asp?ImageID=" + ImageID + "&PINTest=0&ProdID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=600,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }

    function DisplaySingleImageFusion(ImageID) {
        var strID;
        var NewTop = (window.screen.Height - (500 + 200)) / 2;
        var NewLeft = (window.screen.Width - 800) / 2;
        strID = window.open("image/fusion/localization.asp?ImageID=" + ImageID + "&PINTest=0&ProdID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=800,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }
    function DisplaySingleImagePulsar(ImageID) {
        var strID;
        var NewTop = (window.screen.Height - (500 + 200)) / 2;
        var NewLeft = (window.screen.Width - 800) / 2;
        strID = window.open("image/fusion/localization_Pulsar.asp?ImageID=" + ImageID + "&PINTest=0&ProdID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=800,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }

    function ValidateSingleImage(ImageID) {
        var strID;
        var NewTop = (window.screen.Height - (500 + 200)) / 2;
        var NewLeft = (window.screen.Width - 600) / 2;
        strID = window.open("image/CompareImageChoose.asp?ImageDefinitionID=" + ImageID + "&PINTest=0&ProdID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=600,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }

    function ValidateSingleImageFusion(ImageID) {
        var strID;

        var NewTop = Math.floor((screen.height - 700) / 2);
        var NewLeft = Math.floor((screen.width - 600) / 2);;
        strID = window.open("image/Fusion/CompareFusionImage.asp?ImageDefinitionID=" + ImageID + "&PINTest=0&ProductID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=600,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }
    function ValidateSingleImagePulsar(ImageID) {
        var strID;

        var NewTop = Math.floor((screen.height - 700) / 2);
        var NewLeft = Math.floor((screen.width - 600) / 2);;
        strID = window.open("image/Fusion/CompareImage_Pulsar.asp?ImageDefinitionID=" + ImageID + "&PINTest=0&ProductID=" + txtID.value, "_blank", "left=" + NewLeft + ",top=" + NewTop + ",width=600,height=500,menubar=yes,location=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
    }

    function AddToolSchedule(ID) {
        var strID;
        strID = window.showModalDialog("actions/schedule.asp?ProductID=" + ID, "", "dialogWidth:655px;dialogHeight:450px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function roadmaprows_onclick(ID) {
        var strResult;
        strResult = window.showModalDialog("actions/schedule.asp?ID=" + ID, "", "dialogWidth:655px;dialogHeight:520px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

    function toolactionrows_onclick(ID, strType) {
        var strResult;
        strResult = window.showModalDialog("actions/action.asp?Type=" + strType + "&ID=" + ID, "", "dialogWidth:655px;dialogHeight:550px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

    function toolactionrowDetails_onclick(ID) {
        var strResult;
        strResult = window.open("Query/ActionReport.asp?txtFunction=2&txtNumbers=" + ID, "_blank");
    }

    function AddAction(ID) {
        var strID;
        strID = ShowPropertiesDialog("mobilese/today/action.asp?ProdID=" + ID + "&Type=2", "Add Task", 655, 650);
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function AddTask(ID) {
        modalDialog.open({ dialogTitle: 'Add Task', dialogURL: 'Actions/Action.asp?ID=0&Working=1&ProdID=' + ID + '', dialogHeight: 655, dialogWidth: 650, dialogResizable: true, dialogDraggable: true });
    }

    function AddToolAction(ID, strWorking, strType) {
        var strID;
        strID = window.showModalDialog("actions/action.asp?ID=0&Working=" + strWorking + "&ProdID=" + ID + "&Type=" + strType, "", "dialogWidth:655px;dialogHeight:550px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function AddOpportunity(ID) {
        var strID;
        strID = window.showModalDialog("mobilese/today/action.asp?ProdID=" + ID + "&Type=5", "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function AddStatus(ID) {
        var strID;
        strID = window.showModalDialog("mobilese/today/action.asp?ProdID=" + ID + "&Type=4", "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function AddIssue(ID) {
        var strID;
        strID = window.showModalDialog("mobilese/today/action.asp?ProdID=" + ID + "&Type=1", "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    function DisplayStatus(strStatus) {

        var expireDate = new Date();

        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "PMStatus=" + strStatus + ";expires=" + expireDate.toGMTString() + ";";

        window.location.reload(true);
    }

    function AddImage() {
        modalDialog.open({ dialogTitle: 'Add Image', dialogURL: 'image/image.asp?ProdID=' + txtID.value + '', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });
    }

    function AddImageFusion() {
        modalDialog.open({ dialogTitle: 'Add Image', dialogURL: 'image/fusion/image.asp?ProdID=' + txtID.value + '', dialogHeight: 650, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });
    }

    function AddImagePulsar() {
        modalDialog.open({ dialogTitle: 'Add Image', dialogURL: 'image/fusion/Image_Pulsar.asp?ProdID=' + txtID.value + '', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });
    }

    function RevImage() {
        var strID;

        strID = window.showModalDialog("image/RevImages.asp?ProductID=" + txtID.value, "", "dialogWidth:840px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
        if (typeof (strID) != "undefined") {
            if (strID == 1) {
                window.location.reload(true);
            }
        }
    }
    function EditDriveDefinitions() {
        var strID;

        strID = window.showModalDialog("image/ImageDriveDefinition.asp", "", "dialogWidth:600px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    function OemReadyVerify() {
        var strID;

        strID = window.open("image/WhqlVerify.asp?ProdID=" + txtID.value);
    }

    function ImportImage() {
        modalDialog.open({ dialogTitle: 'Import Images', dialogURL: 'image/import.asp?ProdID=' + txtID.value + '', dialogHeight: 650, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });

    }

    function ImportImageFusion() {
        modalDialog.open({ dialogTitle: 'Import Images', dialogURL: 'image/fusion/import.asp?ProdID=' + txtID.value + '', dialogHeight: 650, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });

    }

    function ImportImagePulsar(PRLList) {
        ShowPropertiesDialog("image/fusion/Import_Pulsar.asp?ProdID=" + txtID.value + '&ProductReleaseID=<%=Request("ProductReleaseID")%>&ProductOSReleaseID=<%=Request("ProductOSReleaseID")%>', "Import OS Definition", 800, 700);
    }

    function CopyImage(ImageID) {
        modalDialog.open({ dialogTitle: 'Copy Image', dialogURL: 'image/image.asp?CopyID=' + ImageID + '&ProdID=' + txtID.value + '', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });
    }

    function CopyImageFusion(ImageID) {
        modalDialog.open({ dialogTitle: 'Copy Image', dialogURL: 'image/fusion/image.asp?CopyID=' + ImageID + '&ProdID=' + txtID.value + '&CopyTarget=0', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });

    }

    function CopyImagePulsar(ImageID) {
        modalDialog.open({ dialogTitle: 'Copy Image', dialogURL: 'image/fusion/image_Pulsar.asp?CopyID=' + ImageID + '&ProdID=' + txtID.value + '&CopyTarget=0', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });
    }

    //******************************************************
    //Description:  Enable Copy with Targeting for Legacy Products
    //Function:     CopyWithTarget_Fusion()
    //Modified:     Harris, Valerie (5/31/2016) - PBI 19513/ Task 20989
    //******************************************************
    function CopyWithTarget_Fusion(ImageID) {
        modalDialog.open({ dialogTitle: 'Copy with Targeting', dialogURL: 'image/fusion/image.asp?CopyID=' + ImageID + '&ProdID=' + txtID.value + '&CopyTarget=1', dialogHeight: 650, dialogWidth: 655, dialogResizable: true, dialogDraggable: true });

    }

    //******************************************************
    //Description:  Copy Pulsar Image with Component's Targeted Data
    //Function:     CopyWithTarget_Pulsar()
    //Modified:     Harris, Valerie (3/14/2016) - PBI 17835/ Task 18059
    //******************************************************
    function CopyWithTarget_Pulsar(ImageID) {
        modalDialog.open({ dialogTitle: 'Copy with Targeting', dialogURL: 'image/fusion/image_Pulsar.asp?CopyID=' + ImageID + '&ProdID=' + txtID.value + '&CopyTarget=1', dialogHeight: 700, dialogWidth: 705, dialogResizable: true, dialogDraggable: true });
    }

    function AddVersion_onmouseover() {
        window.event.srcElement.style.color = "red";
        window.event.srcElement.style.cursor = "hand";
    }

    function AddVersion_onmouseout() {
        window.event.srcElement.style.color = "blue";
    }

    function AddVersion_onclick() {
        window.showModalDialog("WizardFrames.asp", "", "dialogWidth:800px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    function window_onload() {
        var Found;
        var strFavorites;
        var SelectedTab;
        if (txtFavs != undefined) {
            strFavorites = "," + txtFavs.value;
            Found = strFavorites.indexOf(",P" + trim(txtID.value) + ",");
        }

        if (txtClass.value == "") {
            RFLink.style.display = "none";
            AFLink.style.display = "none";
            StatusLink.style.display = "";
        }

        if (Found == -1) {
            RFLink.style.display = "none";
            AFLink.style.display = "";
        }
        else {
            RFLink.style.display = "";
            AFLink.style.display = "none";
        }
        Wait.style.display = "none";
        FavLinksTable.style.display = "";

        if (txtError.value != "1") {
            document.getElementById('menubar').style.display = "";

            if (txtProductType.value != "2") {
                SelectedTab = txtDisplayedList.value;
                var blnLoad = 0;

                if (window.location != undefined) {
                    var queryString = window.location.search.replace("?", "");
                    var pieces = queryString.split("&");
                    if (pieces.length > 1) {
                        for (var i = 0; i <= pieces.length - 1; i++) {
                            if (pieces[i] != undefined && pieces[i].indexOf("List=") > -1) {
                                var val = pieces[i];
                                SelectedTab = val.replace("List=", "");
                                if (SelectedTab == "Calls") {
                                    blnLoad = 1;
                                }
                            }
                        }
                    }
                }

                if (SelectedTab == "SCM")
                    SelectTab("SCM", blnLoad);
                else if (SelectedTab == "OTS")
                    SelectTab("OTS", blnLoad);
                else if (SelectedTab == "Issue")
                    SelectTab("Issue", blnLoad);
                else if (SelectedTab == "Action")
                    SelectTab("Action", blnLoad);
                else if (SelectedTab == "Documents")
                    SelectTab("Documents", blnLoad);
                else if (SelectedTab == "Deliverables")
                    SelectTab("Deliverables", blnLoad);
                else if (SelectedTab == "Agency")
                    SelectTab("Agency", blnLoad);
                else if (SelectedTab == "Schedule")
                    SelectTab("Schedule", blnLoad);
                else if (SelectedTab == "Status")
                    SelectTab("Status", blnLoad);
                else if (SelectedTab == "Local")
                    SelectTab("Local", blnLoad);
                else if (SelectedTab == "General")
                    SelectTab("General", blnLoad);
                else if (SelectedTab == "Country")
                    SelectTab("Country", blnLoad);
                else if (SelectedTab == "Requirements") {
                    SelectTab("Requirements", blnLoad);
                    resizeIframe();
                }
                else if (SelectedTab == "PMR")
                    SelectTab("PMR", blnLoad);
                else if (SelectedTab == "Opportunity")
                    SelectTab("Opportunity", blnLoad);
                else if (SelectedTab == "Calls")
                    SelectTab("Calls", blnLoad);
                else if (SelectedTab == "Country")
                    SelectTab("Country", blnLoad);
                else
                    SelectTab("DCR", blnLoad);
            }
        }
        var expireDate = new Date();
        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "LastProductDisplayed=" + trim(txtID.value) + ";expires=" + expireDate.toGMTString() + ";path=<%=AppRoot %>/";

        //Instantiate modalDialog load
        modalDialog.load();
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


    function View_onmouseover() {
        window.event.srcElement.style.color = "red";
        window.event.srcElement.style.cursor = "hand";
    }

    function View_onmouseout() {
        window.event.srcElement.style.color = "blue";
    }

    function View_onclick() {
        window.showModalDialog("WizardFrames.asp?Type=1", "", "dialogWidth:800px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
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
        //Reset all tabs

        document.all("CellGeneralb").style.display = "none";
        document.all("CellGeneral").style.display = "";

        document.all("CellOTSb").style.display = "none";
        document.all("CellOTS").style.display = "";

        document.all("CellCountryb").style.display = "none";
        document.all("CellCountry").style.display = "";

        document.all("CellPMRb").style.display = "none";
        document.all("CellPMR").style.display = "";

        document.all("CellDCRb").style.display = "none";
        document.all("CellDCR").style.display = "";
        document.all("CellActionb").style.display = "none";
        document.all("CellAction").style.display = "";
        document.all("CellOpportunityb").style.display = "none";
        document.all("CellOpportunity").style.display = "none";
        document.all("CellCallsb").style.display = "none";
        document.all("CellCalls").style.display = "";

        document.all("CellDocumentsb").style.display = "none";
        document.all("CellDocuments").style.display = "";

        document.all("CellDeliverablesb").style.display = "none";
        document.all("CellDeliverables").style.display = "";
        document.all("CellScheduleb").style.display = "none";
        document.all("CellLocalb").style.display = "none";
        document.all("CellLocal").style.display = "";

        document.all("CellScheduleb").style.display = "none";
        document.all("CellSchedule").style.display = "";

        document.all("CellRequirementsb").style.display = "none";
        document.all("CellRequirements").style.display = "";

        document.all("CellAgencyb").style.display = "none";
        document.all("CellAgency").style.display = "";

        document.all("CellSCMb").style.display = "none";
        document.all("CellSCM").style.display = "";

        //Highight the selected tab
        document.all("Cell" + strStep).style.display = "none";
        document.all("Cell" + strStep + "b").style.display = "";

        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "PMTab=" + strStep + ";expires=" + expireDate.toGMTString() + ";";

        CurrentState = strStep;

        if (strStep == "Calls") {
            window.location = "service/pmview.asp?ID=" + txtID.value + "&Class=" + txtClass.value + "&List=" + strStep;
            return;
        }

        if (blnLoad == 1) {
            window.location = "pmview.asp?ID=" + txtID.value + "&Class=" + txtClass.value + "&List=" + strStep;
        }
        else {
            if (strStep == "OTS") {
                TableOTS.style.display = "";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableDocuments.style.display = "none";
                TableIssue.style.display = "none";
                TableLocal.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                AddOTSLink.style.display = "";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "DCR") {
                TableOTS.style.display = "none";
                TableIssue.style.display = "none";
                TableDCR.style.display = "";
                TableAgency.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
                DCRFilters.style.display = "";
            }
            else if (strStep == "Agency") {
                TableOTS.style.display = "none";
                TableIssue.style.display = "none";
                TableAgency.style.display = "";
                TableDCR.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddAgencyLink.style.display = "";
                AddActionLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "PMR") {
                TableOTS.style.display = "none";
                TableIssue.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";

<%
                    dim smrLink
                smrLink = Application("Release_Houston_ServerName") + "/SMR/softpaq/ProductWidget/" + PVID
                    %>

                    TablePMR.style.display = "";
                if ($("#newSMRLoading").is(':visible')) {
                    $("#newSMRiFrame").load(function () {
                        $("#newSMRLoading").css('display', 'none');
                    }).attr('src', '<%=smrLink %>');
                }
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Country") {
                TableOTS.style.display = "none";
                TableIssue.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Schedule") {
                TableSchedule.style.display = "";
                TableOTS.style.display = "none";
                TableIssue.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableLocal.style.display = "none";
                TableDocuments.style.display = "none";
                TableStatus.style.display = "none";
                TableAction.style.display = "none";
                AddAgencyLink.style.display = "none";
                TableOpportunity.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Issue") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Action") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableIssue.style.display = "none";
                TableAgency.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "";
                AddAgencyLink.style.display = "none";
                TableLocal.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Opportunity") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "";
                TableLocal.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Status") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableStatus.style.display = "";
                TableSchedule.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                TableLocal.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "";
            }
            else if (strStep == "Deliverables") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableStatus.style.display = "none";
                TableLocal.style.display = "none";
                TableDeliverables.style.display = "";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddDeliverablesLink.style.display = "";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
                if (typeof (window.parent.frames["TitleWindow"]) != "undefined") {
                    if (typeof (window.parent.frames["TitleWindow"].SavedValue) != "undefined") {
                        if (window.parent.frames["TitleWindow"].SavedValue.value != "") {
                            window.document.body.scrollTop = window.parent.frames["TitleWindow"].SavedValue.value;
                            window.parent.frames["TitleWindow"].SavedValue.value = "";
                        }
                    }
                }
            }

            else if (strStep == "Local") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableSchedule.style.display = "none";
                TableOpportunity.style.display = "none";
                TableLocal.style.display = "";
                TableStatus.style.display = "none";
                TableDeliverables.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                AddLocalLink.style.display = "";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "Requirements") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableLocal.style.display = "none";
                TableStatus.style.display = "none";
                TableDeliverables.style.display = "none";
                TableRequirements.style.display = "";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "";
                AddDeliverablesLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else if (strStep == "General") {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableLocal.style.display = "none";
                TableStatus.style.display = "none";
                TableDeliverables.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "";
                AddRequirementsLink.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
            else {
                TableOTS.style.display = "none";
                TableDCR.style.display = "none";
                TableAgency.style.display = "none";
                TableIssue.style.display = "none";
                TableDocuments.style.display = "";
                TableStatus.style.display = "none";
                TableLocal.style.display = "none";
                TableAction.style.display = "none";
                TableOpportunity.style.display = "none";
                TableSchedule.style.display = "none";
                TableRequirements.style.display = "none";
                TablePMR.style.display = "none";
                TableCountry.style.display = "none";
                TableGeneral.style.display = "none";
                AddCountryLink.style.display = "none";
                AddGeneralLink.style.display = "none";
                AddRequirementsLink.style.display = "none";
                AddChangeLink.style.display = "none";
                AddActionLink.style.display = "none";
                AddIssueLink.style.display = "none";
                AddOpportunityLink.style.display = "none";
                AddOTSLink.style.display = "none";
                AddAgencyLink.style.display = "none";
                AddScheduleLink.style.display = "none";
                TableDeliverables.style.display = "none";
                AddDeliverablesLink.style.display = "none";
                AddLocalLink.style.display = "none";
                AddStatusLink.style.display = "none";
            }
        }
    }

    function Export(intType) {
        if (intType == 1)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableIssue.innerHTML + "</TABLE>";
        else if (intType == 2)
            ExportForm.txtData.value = "<TABLE  BORDER=1>" + TableAction.innerHTML + "</TABLE>";
        else if (intType == 3) {
            var url = "/Pulsar/Product/ChangeRequestProduct?userId=" + document.getElementById("txtUser").value + "&prodId=<%=PVID%>&statusFilter=" + hidDCRFilterStatus.value;
            modalDialog.open({
                dialogTitle: 'Export List to Excel',
                dialogURL: url,
                dialogHeight: 550,
                dialogWidth: 850,
                dialogResizable: true, dialogDraggable: true
            });
            return;
        }
        else if (intType == 4)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableOTS.innerHTML + "</TABLE>";
        else if (intType == 5)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableDeliverables.innerHTML + "</TABLE>";
        else if (intType == 6)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableLocal.innerHTML + "</TABLE>";
        else if (intType == 7)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableStatus.innerHTML + "</TABLE>";
        else if (intType == 8)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableRequirements.innerHTML + "</TABLE>";
        else if (intType == 9) {
            var divs = document.getElementsByName("schedule_tooltip")
            for (var i = 0; i < divs.length; i = i + 1) {
                divs[i].innerHTML = "";
            }
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableSchedule.innerHTML + "</TABLE>";
        }
        else if (intType == 10)
            ExportForm.txtData.value = "<TABLE BORDER=1></TABLE>";
        else if (intType == 11)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableGeneral.innerHTML + "</TABLE>";
        else if (intType == 12)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableCountry.innerHTML + "</TABLE>";
        else if (intType == 13)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableOpportunity.innerHTML + "</TABLE>";
        else if (intType == 14)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableAgency.innerHTML + "</TABLE><TABLE><TR><TD>* Country added after POR by DCR</TD></TR></TABLE>";
        else if (intType == 15)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + ToolTable.innerHTML + "</TABLE>";

        ExportForm.submit();
    }

    function ExportWord(intType) {
        if (intType == 1)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableIssue.innerHTML + "</TABLE>";
        else if (intType == 2)
            ExportWordForm.txtBody.value = "<TABLE  BORDER=1>" + TableAction.innerHTML + "</TABLE>";
        else if (intType == 3)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableDCR.innerHTML + "</TABLE>";
        else if (intType == 4)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableOTS.innerHTML + "</TABLE>";
        else if (intType == 5)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableDeliverables.innerHTML + "</TABLE>";
        else if (intType == 6)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableLocal.innerHTML + "</TABLE>";
        else if (intType == 7)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableStatus.innerHTML + "</TABLE>";
        else if (intType == 8)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableRequirements.innerHTML + "</TABLE>";
        else if (intType == 9) {
            var divs = document.getElementsByName("schedule_tooltip")
            for (var i = 0; i < divs.length; i = i + 1) {
                divs[i].innerHTML = "";
            }

            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableSchedule.innerHTML + "</TABLE>";
        }
        else if (intType == 10)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1></TABLE>";
        else if (intType == 11)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableGeneral.innerHTML + "</TABLE>";
        else if (intType == 12)
            ExportWordForm.txtBody.value = "<TABLE BORDER=1>" + TableCountry.innerHTML + "</TABLE>";
        else if (intType == 13)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableOpportunity.innerHTML + "</TABLE>";
        else if (intType == 14)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + TableAgency.innerHTML + "</TABLE>";
        else if (intType == 15)
            ExportForm.txtData.value = "<TABLE BORDER=1>" + ToolTable.innerHTML + "</TABLE>";

        ExportWordForm.submit();
    }

    function setDcrFilterType(type) {
        var expireDate = new Date();
        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "DCRFilterType=" + type + ";expires=" + expireDate.toGMTString() + ";";
        window.location.reload(true);
    }

    function setDcrFilterStatus(status) {
        var expireDate = new Date();
        expireDate.setMonth(expireDate.getMonth() + 12);
        document.cookie = "DCRFilterStatus=" + status + ";expires=" + expireDate.toGMTString() + ";";
        window.location.reload(true);
    }

    function Details(intType) {
        DetailsForm.Query.value = "spListActionItemsDetails " + txtID.value + "," + intType + "," + txtDisplayedStatus.value;
        DetailsForm.submit();
    }

    function ShowOptions(intType) {
        var strID;
        var strQuery = ""
        if (intType == 3) {
            strQuery = "spListActionItemsDetails " + txtID.value + "," + intType + "," + hidDCRFilterStatus.value + "," + hidDCRFilterType.value + "," + hidDCRFilterScr.value;
            modalDialog.open({ dialogTitle: 'Export Details to Excel', dialogURL: 'exportoptions.asp?ActionType=' + intType + '&ID=' + txtID.value + '&Type=' + intType + '&Status=' + hidDCRFilterStatus.value + '&Bios=' + hidDCRFilterType.value + '&Scr=' + hidDCRFilterScr.value + '', dialogHeight: 580, dialogWidth: 470, dialogArguments: strQuery, dialogArgumentsName: 'export_option_query', dialogResizable: true, dialogDraggable: true });
        }
        else {
            strQuery = "spListActionItemsDetails " + txtID.value + "," + intType + "," + txtDisplayedStatus.value;
            modalDialog.open({ dialogTitle: 'Export Details to Excel', dialogURL: 'exportoptions.asp?ActionType=' + intType + '&ID=' + txtID.value + '', dialogHeight: 580, dialogWidth: 470, dialogArguments: strQuery, dialogArgumentsName: 'export_option_query', dialogResizable: true, dialogDraggable: true });
        }
    }

    function ModifyMilestoneList(ProductVersionID, ScheduleID) {
        modalDialog.open({ dialogTitle: 'Add/Remove Items', dialogURL: 'Schedule/Schedule.asp?PVID=' + ProductVersionID + '&ScheduleID=' + ScheduleID + '', dialogHeight: 700, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
    }

    function ModifyMilestoneList_Pulsar(ProductVersionID, ScheduleID) {
        modalDialog.open({ dialogTitle: 'Add/Remove Items', dialogURL: 'Schedule/Schedule_Pulsar.asp?PVID=' + ProductVersionID + '&ScheduleID=' + ScheduleID + '', dialogHeight: 700, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
    }

    function EditScheduleDescription(ProductVersionID, ScheduleID, IsPulsarProduct) {
        if (IsPulsarProduct == 1)
            modalDialog.open({ dialogTitle: 'Rename Custome Schedule', dialogURL: 'Schedule/ScheduleDescription.asp?PVID=' + ProductVersionID + '&ScheduleID=' + ScheduleID + '&IsPulsarProduct=1', dialogHeight: 260, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
        else
            modalDialog.open({ dialogTitle: 'Rename Schedule', dialogURL: 'Schedule/ScheduleDescription.asp?PVID=' + ProductVersionID + '&ScheduleID=' + ScheduleID + '&IsPulsarProduct=0', dialogHeight: 260, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
    }

    function AddNewSchedule(ProductVersionID, IsPulsarProduct) {
        if (IsPulsarProduct == 1)
            modalDialog.open({ dialogTitle: 'Add Custom Schedule', dialogURL: 'Schedule/CreateScheduleFrame.asp?PVID=' + ProductVersionID + '&IsPulsarProduct=1', dialogHeight: 260, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
        else
            modalDialog.open({ dialogTitle: 'Add Schedule', dialogURL: 'Schedule/CreateScheduleFrame.asp?PVID=' + ProductVersionID + '&IsPulsarProduct=0', dialogHeight: 260, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
    }

    function AddNewScheduleResult(strResult) {
        if (typeof (strResult) != "undefined") {
            window.location.href = "../pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>&ScheduleID=" + strResult;
        }
    }

    function DeleteSchedule(ScheduleID) {
        modalDialog.open({ dialogTitle: 'Delete Schedule', dialogURL: 'Schedule/DeleteSchedule.asp?SID=' + ScheduleID + '', dialogHeight: 150, dialogWidth: 250, dialogResizable: true, dialogDraggable: true });

    }

    function DeactivateSchedule(ScheduleID, ProductVersionID) {
        modalDialog.open({ dialogTitle: 'Deactivate Schedule', dialogURL: 'Schedule/DeactivateSchedule.asp?SID=' + ScheduleID + '&PVID=' + ProductVersionID + '', dialogHeight: 150, dialogWidth: 250, dialogResizable: true, dialogDraggable: true });
    }

    function RemoveScheduleResult(strResult) {
        if (typeof (strResult) != "undefined") {
            window.location.href = "../pmview.asp?ID=<%=PVID%>&Class=<%=sClass%>";
        }
    }

    function ScheduleBatchEdit(ProductVersionID, ScheduleID, Mode) {
        var sMode = '';

        if (Mode === 'Projected') {
            sMode = 'Current';
        } else {
            sMode = 'Actual';
        }
        modalDialog.open({ dialogTitle: 'Batch Edit ' + sMode + '', dialogURL: 'Schedule/ScheduleBatchUpdate.asp?PVID=' + ProductVersionID + '&Mode=' + Mode + '&ScheduleID=' + ScheduleID + '', dialogHeight: 700, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });

    }

    function ScheduleBatchEdit_Pulsar(ProductVersionID, ScheduleID, Mode) {
        var sMode = '';

        if (Mode === 'Projected') {
            sMode = 'Current';
        } else {
            sMode = 'Actual';
        }

        modalDialog.open({ dialogTitle: 'Batch Edit ' + sMode + '', dialogURL: 'Schedule/ScheduleBatchUpdate_Pulsar.asp?PVID=' + ProductVersionID + '&Mode=' + Mode + '&ScheduleID=' + ScheduleID + '', dialogHeight: 700, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
    }

    function CopyMilestoneList(ProductVersionID, ScheduleID) {
        modalDialog.open({ dialogTitle: 'Copy Items', dialogURL: 'Schedule/ScheduleCopy.asp?PVID=' + ProductVersionID + '&ScheduleID=' + ScheduleID + '', dialogHeight: 300, dialogWidth: 650, dialogResizable: true, dialogDraggable: true });
    }

    function ModifyCountryList(ProdID, ProdBrandID, FusionRequirements) {
        modalDialog.open({ dialogTitle: 'Add/Remove Countries', dialogURL: 'Countries/Countries.asp?ProdID=' + ProdID + '&ProdBrandID=' + ProdBrandID + '&IsPulsarProduct=' + FusionRequirements, dialogHeight: 700, dialogWidth: 900, dialogResizable: true, dialogDraggable: true });
    }

    function CopyLocalization(ProdID, ProdBrandID, FusionRequirements) {
        modalDialog.open({ dialogTitle: 'Import', dialogURL: 'Countries/Copy.asp?PVID=' + ProdID + '&BID=' + ProdBrandID + '&IsPulsarProduct=' + FusionRequirements, dialogHeight: 500, dialogWidth: 500, dialogResizable: true, dialogDraggable: true });
    }

    function ModifyRequirementList(intProdID) {
        modalDialog.open({ dialogTitle: 'Add/Remove Requirements', dialogURL: 'Requirements/Requirement.asp?ProdID=' + intProdID, dialogHeight: 650, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function ImportRequirementList(intProdID) {
        modalDialog.open({ dialogTitle: 'Import Requirements', dialogURL: 'Requirements/RequirementImport.asp?ProductID=' + intProdID, dialogHeight: 650, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function RootProperties(RootID) {
        window.showModalDialog("root.asp?ID=" + RootID + "", "Root Properties", "dialogWidth:" + GetWindowSize("width") + "px;dialogHeight:" + GetWindowSize("height") + "px;maximize:yes;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }


    function RootRemove(ProductID, RootID) {
        //save Product and Version ID for results function: ---
        globalVariable.save(ProductID, 'product_id');
        globalVariable.save(RootID, 'root_id');

        modalDialog.open({ dialogTitle: 'Remove Root', dialogURL: 'target/RemoveRoot.asp?ProductID=' + ProductID + '&DeliverableID=' + RootID + '', dialogHeight: 350, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function RootRemoveResults(strID) {
        var ProductID = globalVariable.get('product_id');
        var RootID = globalVariable.get('root_id');

        if (typeof (strID) != "undefined") {
            document.all("DelRow" + ProductID + "_" + RootID).style.display = "none";
        }
    }

    function AdvancedTarget(ProdID, VerID, RootID) {
        var strResult;
        var oDeliverableRow;
        var ExcludeFunComp = false;
        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(RootID, 'root_id');
        globalVariable.save(VerID, 'version_id');


        if (document.getElementById('divFunCompExclude').innerHTML.indexOf('Include') < 0)
            ExcludeFunComp = true;

        modalDialog.open({ dialogTitle: 'Target Version', dialogURL: 'Target/TargetAdvanced.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID + '' + '&ExcludeFunComp=' + ExcludeFunComp, dialogHeight: (GetWindowSize("height")), dialogWidth: (GetWindowSize("width")), dialogResizable: true, dialogDraggable: true });

        if (typeof (SelectedRow) != "undefined") {
            SelectedRow.style.color = "black";
            SelectedRow = null;
        }
    }

    //*****************************************************************
    //Description:  Open Target Advanced Modal Dialog
    //Function:     AdvancedTargetResult();
    //Modified:     Harris, Valerie (9/27/2016) - PBI 26986/Task 27006
    //*****************************************************************
    function AdvancedTargetResult(strResult) {
        var oDeliverableRow;
        var iProductID = globalVariable.get('product_id');
        var iRootID = globalVariable.get('root_id');
        var iVersionID = globalVariable.get('version_id');

        //---If result isn't undefined: ---
        if (typeof (strResult) != "undefined") {
            if (typeof (window.parent.frames["TitleWindow"]) != "undefined") {
                if (typeof (window.parent.frames["TitleWindow"].SavedValue) != "undefined") {
                    window.parent.frames["TitleWindow"].SavedValue.value = window.document.body.scrollTop;
                }
            }

            //---If form processed sucessfully, change background color of the the Root Target's row: ---
            if (strResult === 1 || strResult === '1') {
                if ($(".deliverable" + iProductID + "_" + iRootID + "_" + iVersionID).length > 0) {

                    //get object
                    oDeliverableRow = $(".deliverable" + iProductID + "_" + iRootID + "_" + iVersionID + " td");

                    //loop thru each cell and change background color to signify the record has been updated
                    oDeliverableRow.each(function (index, row) {
                        $(this).attr("bgcolor", "#ccffff");
                    });
                }

                //---Show Refresh button: ---
                $("#btnRefresh").removeClass("hide").addClass("show");
            } else {
                window.location.reload(true);
            }
        }

    }

    function TargetVersion(ProdID, VerID, Type) {

        var strResult;
        if (Type == 0) {
            modalDialog.open({ dialogTitle: 'Target Version', dialogURL: 'Target/TargetQuickSave.asp?ProdID=' + ProdID + '&VersionID=' + VerID + '&Type=' + Type + '&Rejected=1', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
        } else {
            modalDialog.open({ dialogTitle: 'Target Version', dialogURL: 'Target/TargetQuickSave.asp?ProdID=' + ProdID + '&VersionID=' + VerID + '&Type=' + Type + '', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
        }

        setTimeout(function () {
            strResult = modalDialog.getArgument('target_save_status');
        }, 1000);

        setTimeout(function () {
            if (typeof (SelectedRow) != "undefined") {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
            if (typeof (strResult) != "undefined") {
                if (typeof (window.parent.frames["TitleWindow"]) != "undefined") {
                    if (typeof (window.parent.frames["TitleWindow"].SavedValue) != "undefined") {
                        window.parent.frames["TitleWindow"].SavedValue.value = window.document.body.scrollTop;
                    }
                }
                window.location.reload(true);
            }
        }, 1000);


    }

    function UpdateServiceEOLDate(ID, TypeID) {
        var strUpdated;

        if (typeof (SelectedRow) != "undefined") {
            SelectedRow.style.color = "black";
            SelectedRow = null;
        }

        strUpdated = window.showModalDialog("Deliverable/EOLDate.asp?TypeID=" + TypeID + "&ID=" + ID, "", "dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
    }

    function DisplayPilotVersion(ID, RootID, VersionID) {
        var strResult;
        strResult = window.showModalDialog("Deliverable/EOLDate.asp?ID=" + VersionID, "", "dialogWidth:550px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")

        if (typeof (SelectedRow) != "undefined") {
            SelectedRow.style.color = "black";
            SelectedRow = null;
        }
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

    function DisplayVersion(ID, RootID, VersionID) {
        window.showModalDialog("WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID + "", "", "dialogWidth:" + GetWindowSize('width') + "px;dialogHeight:" + GetWindowSize('height') + "px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
    }

    function DisplayVersionResult(refreshPage, strResult) {
        if (typeof (SelectedRow) != "undefined") {
            SelectedRow.style.color = "black";
            SelectedRow = null;
        }
        if (typeof (strResult) != "undefined") {
            modalDialog.cancel(refreshPage)
        }
    }

    function DisplayVersionDetail(ID, RootID, VersionID) {
        var strResult;
        window.open("Query\\DeliverableVersionDetails.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID)

        if (typeof (SelectedRow) != "undefined") {
            SelectedRow.style.color = "black";
            SelectedRow = null;
        }
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

    function Delrows_onmouseout() {
        if (typeof (oPopup) == "undefined")
            return;

        if (!oPopup.isOpen) {
            if (window.event.srcElement.className == "text")
                window.event.srcElement.parentElement.parentElement.style.color = "black";
            else if (window.event.srcElement.className == "cell")
                window.event.srcElement.parentElement.style.color = "black";
        }
    }

    function Delrows_onmouseover() {
        if (typeof (oPopup) == "undefined")
            return;
        if (!oPopup.isOpen) {
            if (window.event.srcElement.className == "cell") {
                window.event.srcElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.style.cursor = "hand";
            }
            else if (window.event.srcElement.className == "text") {
                window.event.srcElement.parentElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
            }

            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    SelectedRow.style.color = "black";

        }
    }

    function EditTestStatus(ProductID, VersionID, RootID, FieldID, ReleaseID, FusionRequirements) {
        if (FusionRequirements == 0) {
            modalDialog.open({ dialogTitle: 'Test Status', dialogURL: 'Deliverable/TestTeam/TestStatus.asp?FieldID=' + FieldID + '&VersionID=' + VersionID + '&ProductID=' + ProductID, dialogHeight: 450, dialogWidth: 700, dialogResizable: true, dialogDraggable: true });
        }
        else {
            modalDialog.open({ dialogTitle: 'Test Status', dialogURL: 'Deliverable/TestTeam/TestStatusPulsar.asp?FieldID=' + FieldID + '&VersionID=' + VersionID + '&ProductID=' + ProductID + '&ReleaseID=' + ReleaseID, dialogHeight: 450, dialogWidth: 700, dialogResizable: true, dialogDraggable: true });
        }
    }


    //
    // Images Tab
    //

    function localizationMenu(ImageID) {
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody = document.getElementById("localizationMenu").innerHTML;
        var NewHeight;
        var NewWidth;

        if (window.event.srcElement.className == "text") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
        else if (window.event.srcElement.className == "cell") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }

        if (window.event.srcElement.className != "check") {

            oPopup.document.body.innerHTML = popupBody.replace(/\[ImageID\]/g, ImageID);
            oPopup.show(lefter, topper, 150, 147, document.body);

            NewHeight = oPopup.document.body.scrollHeight;
            NewWidth = oPopup.document.body.scrollWidth;
            oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
        }
    }

    function localizationMenuFusion(ImageID, ProductDropID) {
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody = document.getElementById("localizationMenuFusion").innerHTML;
        var NewHeight;
        var NewWidth;

        if (window.event.srcElement.className == "text") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
        else if (window.event.srcElement.className == "cell") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }

        if (window.event.srcElement.className != "check") {

            popupBody = popupBody.replace(/\[ImageID\]/g, ImageID);
            popupBody = popupBody.replace(/\[ProductDropID\]/g, ProductDropID);

            if (ProductDropID == 0)
                popupBody = popupBody.replace(/\[DisplayOption1\]/g, "style='display:none'");

            oPopup.document.body.innerHTML = popupBody

            oPopup.show(lefter, topper, 150, 147, document.body);

            NewHeight = oPopup.document.body.scrollHeight;
            NewWidth = oPopup.document.body.scrollWidth;
            oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
        }
    }

    function ChangeDistribution(ProdID, VerID, RootID) {
        var url;
        url = 'Target/ChangeDistribution.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'Modify Distribution', dialogURL: '' + url + '', dialogHeight: 550, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ModDistResult(strResult) {
        var ResultArray;
        var iVersionID = globalVariable.get('version_id');
        var iProductID = globalVariable.get('product_id');

        if (typeof (strResult) != "undefined") {
            ResultArray = strResult.split("|");
            document.all("DistCell" + iVersionID + "_" + iProductID).innerText = ResultArray[0];
            document.all("IMGCell" + iVersionID + "_" + iProductID).innerText = ResultArray[1];
        }

    }

    function DisplayPIProperties(ProdID, RootID, VerID) {
        url = 'Image/PIProperties.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'PreInstall Properties', dialogURL: '' + url + '', dialogHeight: 350, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function DisplayPIPropertiesResults(strResult) {
        var iVersionID = globalVariable.get('version_id');
        var iProductID = globalVariable.get('product_id');

        if (typeof (strResult) != "undefined") {
            document.all("PartCell" + iVersionID + "_" + iProductID).innerText = strResult[0];
            document.all("InIMGCell" + iVersionID + "_" + iProductID).innerText = strResult[1];
        }

    }

    function ChangeImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'Change Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }


    function ChangeIRSImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/Fusion/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'Change Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ChangePulsarImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/Pulsar/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        var width = $(window).width() * 0.80;
        var height = $(window).height() * 0.70;
        modalDialog.open({ dialogTitle: 'Change Images', dialogURL: '' + url + '', dialogHeight: height, dialogWidth: width, dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        //modalDialog.open({ dialogTitle: 'Change Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ViewImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'View Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ViewIRSImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/Fusion/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'View Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ViewPulsarImages(ProdID, VerID, RootID) {
        var url;
        url = 'Target/Pulsar/ChangeImages.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID;
        modalDialog.open({ dialogTitle: 'View Images', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogArguments: 'IMGCell', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }

        //save Product and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ChangeImageResult(strResult) {
        var iVersionID = globalVariable.get('version_id');
        var iProductID = globalVariable.get('product_id');

        if (typeof (strResult) != "undefined") {
            document.all("IMGCell" + iVersionID + "_" + iProductID).innerText = strResult;
        }
    }


    //
    // OTS Tab
    //
    function ShowOTSAdvanced(strID) {
        var i;
        var strIDList = "";
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        if (typeof (chkOTSID.length) == "undefined") {
            if (chkOTSID.checked)
                strIDList = chkOTSID.value
        }
        else {
            for (i = 0; i < chkOTSID.length; i++)
                if (chkOTSID(i).checked)
                    strIDList = strIDList + "," + chkOTSID(i).value
            if (strIDList != "")
                strIDList = strIDList.substr(1)
        }

        if (strIDList.length > 0) {
            strResult = window.open("search/ots/default.asp?lstProduct=" + txtDisplayedProduct.value + "&txtObservationID=" + strIDList, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,status=yes,scrollbars=Yes")
        }
        else {
            strResult = window.open("search/ots/default.asp?lstProduct=" + txtDisplayedProduct.value, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes")
        }
    }

    function ShowOTSDetails(strID) {
        var i;
        var strIDList = "";
        var NewTop;
        var NewLeft;

        var strSort = "";

        if (sortedOn == -1 || sortedOn == 1)
            strSort = "&Sort1Column=o.observationid";
        else if (sortedOn == 2)
            strSort = "&Sort1Column=Priority";
        else if (sortedOn == 3)
            strSort = "&Sort1Column=State";
        else if (sortedOn == 4)
            strSort = "&Sort1Column=owner";
        else if (sortedOn == 5)
            strSort = "&Sort1Column=pm";
        else if (sortedOn == 6)
            strSort = "&Sort1Column=shortdescription";
        //+ ' ' +sortDirection
        if (strSort != "") {
            if (sortDirection == 0)
                strSort = strSort + "&Sort1Direction=asc";
            else
                strSort = strSort + "&Sort1Direction=desc";
        }


        NewLeft = (screen.width - 655) / 2;
        NewTop = (screen.height - 650) / 2;

        if (typeof (strID) != "undefined") {
            strIDList = strID;
        }
        else if (typeof (chkOTSID.length) == "undefined") {
            if (chkOTSID.checked)
                strIDList = chkOTSID.value;
        }
        else {
            for (i = 0; i < chkOTSID.length; i++)
                if (chkOTSID(i).checked)
                    strIDList = strIDList + "," + chkOTSID(i).value;
            if (strIDList != "")
                strIDList = strIDList.substr(1);
        }

        if (strIDList.length > 0) {
            strResult = window.open("search/ots/Report.asp?txtReportSections=1&txtObservationID=" + strIDList + strSort, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No,scrollbars=Yes")
        }
        else {
            alert("You must select at least one observation first");
        }
    }

    function onmouseover_ResetOTS() {
        window.event.srcElement.style.cursor = "hand";
        window.event.srcElement.style.color = "red";
    }

    function onmouseout_ResetOTS() {
        window.event.srcElement.style.color = "blue";
    }

    function onclick_ResetOTS() {
        var i;

        if (typeof (chkOTSID.length) == "undefined") {
            chkOTSID.checked == chkAllOTS.checked;
        }
        else {
            for (i = 0; i < chkOTSID.length; i++)
                chkOTSID(i).checked = chkAllOTS.checked;
        }
    }

    //
    // Tools Project
    //
    function ReorderItems(ID) {
        var strResult;
        strResult = window.showModalDialog("Actions/ScheduleReorder.asp?ID=" + ID, "", "dialogWidth:700px;dialogHeight:380px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

    function ReorderActions(UserID, ReportOption, ProjectID) {
        var strID;

        strID = window.showModalDialog("Actions/WorkingListReorder.asp?ProjectID=" + ProjectID + "&ID=" + UserID + "&ReportOption=" + ReportOption, "", "dialogWidth:900px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        if (typeof (strID) != "undefined") {
            window.location.reload(true);
        }
    }

    //
    // Deliverables Tab : Context Menus
    //

    function DelMenu(ID, RootID, VersionID, Targeted, InImage, CategoryID, TypeID, WorkflowComplete, AccessGroup, SETestLead, ODMTestLead, WWANTestLead, DEVTestLead, ServicePM, Fusion, FusionRequirements, Active, ReleaseID, BSID, DelFilter) {

        if (window.event.srcElement.className == "check")
            return;
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;
        var extraHeight = 0;
        var strImage = "";
        var NewHeight;
        var NewWidth;
        var ShowOnlyTargetedRelease = 1;

        if (DelFilter == "All")
            ShowOnlyTargetedRelease = 0;

        if (window.event.srcElement.className == "text") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
        else if (window.event.srcElement.className == "cell") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    if (SelectedRow != window.event.srcElement.parentElement)
                        SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }


        if (VersionID != 0) {
            var strPath = trim(document.getElementById("Path" + VersionID).value);
            var strImage = trim(document.getElementById("IMGCell" + VersionID + "_" + ID).innerText);
        }

        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";
        if (VersionID == 0) {
            if ((TypeID != 1 && (txtISPM.value != "0" || (txtISPreinstallPM.value == "1" && (CategoryID == 171 || CategoryID == 179 || CategoryID == 170)))) || (TypeID == 1 && (AccessGroup == "1" || AccessGroup == "2" || AccessGroup == "4"))) {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AdvancedTarget(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Target&nbsp;Versions&nbsp;...</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                if (TypeID != 1) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AlertDetails(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Alert&nbsp;Details...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                }

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:RootRemove(" + ID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Remove&nbsp;Root&nbsp;...</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                if (TypeID == 1 && (AccessGroup == "1" || AccessGroup == "2" || AccessGroup == "3" || AccessGroup == "4" || txtSAAdmin.value > 0)) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditSubassembly(" + ID + "," + RootID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Subassembly&nbsp;Number...&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                }

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:RootProperties(" + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Properties&nbsp;</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReloadWindow()'\" >&nbsp;&nbsp;&nbsp;Refresh the Grid </SPAN></FONT></DIV>";


                popupBody = popupBody + "</DIV>";

                oPopup.document.body.innerHTML = popupBody;

                oPopup.show(lefter, topper, 150, 80, document.body);

                //Adjust window size
                NewHeight = oPopup.document.body.scrollHeight;
                NewWidth = oPopup.document.body.scrollWidth;
                if (topper + NewHeight > document.body.clientHeight)
                    topper = document.body.clientHeight - NewHeight;
                if (topper < 0)
                    topper = 0;

                oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
            }
            else
                RootProperties(RootID);
        }
        else if (TypeID == 1) {
            if (AccessGroup == "1" || AccessGroup == "2" || AccessGroup == "3" || AccessGroup == "4") {
                if (WorkflowComplete == 1) {
                    if (AccessGroup == "2" || AccessGroup == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditCommodityStatus(" + ID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Qualification&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (AccessGroup == "3" || AccessGroup == "2" || AccessGroup == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditPilotStatus(" + ID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Pilot&nbsp;Run&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (AccessGroup == "4" || AccessGroup == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditAccessoryStatus(" + ID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Accessory&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (SETestLead == "1" || ODMTestLead == "1" || WWANTestLead == "1" || DEVTestLead == "1" || AccessGroup == "2") {
                        popupBody = popupBody + "<DIV>";
                        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                    }

                    if (SETestLead == "1" || AccessGroup == "2") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",1," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;SE&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (ODMTestLead == "1" || AccessGroup == "2") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",2," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;ODM&nbsp;HW&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (WWANTestLead == "1" || AccessGroup == "2") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",3," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;COMM&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (DEVTestLead == "1" || AccessGroup == "2") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",4," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;DEV&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    if (AccessGroup == "2" || AccessGroup == "1") {
                        popupBody = popupBody + "<DIV>";
                        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";


                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AdvancedTarget(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Advanced&nbsp;Target...</SPAN></FONT></DIV>";
                    }
                    //                  }

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    //   if (ServicePM == 0) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:MultiUpdateTestStatus(" + ID + "," + RootID + ", 0," + ReleaseID + "," + FusionRequirements + ",0," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Batch&nbsp;Update&nbsp;Root&nbsp;Status&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:MultiUpdateTestStatus(0, 0," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + BSID + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Batch&nbsp;Update&nbsp;Product&nbsp;Status&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:MultiUpdateTestStatusLink(" + ID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Batch&nbsp;Update&nbsp;Selected&nbsp;Versions&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:BatchUpdateRestriction(" + RootID + "," + VersionID + "," + ID + "," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Batch&nbsp;Update&nbsp;Restrictions&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                    // }

                }


                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditSubassembly(" + ID + "," + RootID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Subassembly&nbsp;Number...&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

                if (txtServiceCommodityManager.value == "true") {
                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:UpdateServiceEOLDate(" + VersionID + ",2)'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Service&nbsp;EOA&nbsp;Date...&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";
                }

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";


                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReloadWindow()'\" >&nbsp;&nbsp;&nbsp;Refresh the Grid</SPAN></FONT></DIV>";


                popupBody = popupBody + "</DIV>";

                oPopup.document.body.innerHTML = popupBody;

                oPopup.show(lefter, topper, 150, 80, document.body);

                //Adjust window size
                NewHeight = oPopup.document.body.scrollHeight;
                NewWidth = oPopup.document.body.scrollWidth;
                if (topper + NewHeight > document.body.clientHeight)
                    topper = document.body.clientHeight - NewHeight;
                if (topper < 0)
                    topper = 0;

                oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);

            }
            else {
                if (AccessGroup == "5") {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:SendDelEmail(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayPilotVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;End&nbsp;of&nbsp;Availability&nbsp;Date...&nbsp;</SPAN></FONT></DIV>";


                    popupBody = popupBody + "</DIV>";

                    oPopup.document.body.innerHTML = popupBody;

                    oPopup.show(lefter, topper, 150, 80, document.body);

                    //Adjust window size
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
                }
                else if (ODMTestLead == "1" || WWANTestLead == "1" || DEVTestLead == "1" || SETestLead == "1") {
                    if (SETestLead == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",1," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;SE&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }
                    if (ODMTestLead == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",2," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;ODM&nbsp;HW&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }
                    if (WWANTestLead == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",3," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;COMM&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }
                    if (DEVTestLead == "1") {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditTestStatus(" + ID + "," + VersionID + "," + RootID + ",4," + ReleaseID + "," + FusionRequirements + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;DEV&nbsp;Test&nbsp;Status...</SPAN></FONT></DIV>";
                    }

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:SendDelEmail(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                    popupBody = popupBody + "</DIV>";

                    oPopup.document.body.innerHTML = popupBody;

                    oPopup.show(lefter, topper, 150, 80, document.body);

                    //Adjust window size
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
                }
                else if (txtServiceCommodityManager.value == "true") {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:SendDelEmail(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email&nbsp;...</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:UpdateServiceEOLDate(" + VersionID + ",2)'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Service&nbsp;EOA&nbsp;Date...&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";

                    popupBody = popupBody + "</DIV>";

                    oPopup.document.body.innerHTML = popupBody;

                    oPopup.show(lefter, topper, 150, 80, document.body);

                    //Adjust window size
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);

                }
                else if (txtSAAdmin.value > 0) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN><FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditSubassembly(" + ID + "," + RootID + "," + VersionID + "," + ReleaseID + "," + FusionRequirements + "," + ShowOnlyTargetedRelease + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Subassembly&nbsp;Number...&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";
                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                    oPopup.document.body.innerHTML = popupBody;

                    oPopup.show(lefter, topper, 150, 80, document.body);

                    //Adjust window size
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
                }
                else
                    DisplayVersion(ID, RootID, VersionID);
            }
        }
        else if (txtISPM.value == "0" && !(txtISPreinstallPM.value == "1" && (CategoryID == 171 || CategoryID == 179 || CategoryID == 170))) {

            if (strPath != "") {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:GetVersion(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Download</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                if (txtISPreinstall.value == "1") {

                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayPIProperties(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Preinstall&nbsp;Properties&nbsp;</SPAN></FONT></DIV>";

                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                }

                if (strImage != "-") {
                    if (FusionRequirements == 1) {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewPulsarImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    else if (Fusion == 1) {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewIRSImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    else {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                }

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";

                popupBody = popupBody + "</DIV>";

                oPopup.document.body.innerHTML = popupBody;

                oPopup.show(lefter, topper, 150, 60, document.body);

                //Adjust window size
                if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
                }

            }
            else if (txtISPreinstall.value == "1") {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayPIProperties(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Preinstall&nbsp;Properties&nbsp;</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                if (strImage != "-") {
                    if (FusionRequirements == 1) {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewPulsarImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    else if (Fusion == 1) {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewIRSImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    else {
                        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                        popupBody = popupBody + "<FONT face=Arial size=2>";
                        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ViewImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Images...</SPAN></FONT></DIV>";
                    }
                    popupBody = popupBody + "<DIV>";
                    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
                }

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Properties&nbsp;</SPAN></FONT></DIV>";

                popupBody = popupBody + "</DIV>";

                oPopup.document.body.innerHTML = popupBody;

                oPopup.show(lefter, topper, 100, 40, document.body);

                //Adjust window size
                if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
                    NewHeight = oPopup.document.body.scrollHeight;
                    NewWidth = oPopup.document.body.scrollWidth;
                    if (topper + NewHeight > document.body.clientHeight)
                        topper = document.body.clientHeight - NewHeight;
                    if (topper < 0)
                        topper = 0;

                    oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
                }


            }
            else {

                if (AccessGroup == "5")
                    DisplayPilotVersion(ID, RootID, VersionID);
                else
                    DisplayVersion(ID, RootID, VersionID);
            }
        }

        else {
            if (Targeted == 0 && Active == 1) {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:TargetVersion(" + ID + "," + VersionID + ",1)'\" >&nbsp;&nbsp;&nbsp;Target&nbsp;Version&nbsp;</SPAN></FONT></DIV>";
            }
            else if (Targeted == 0 && Active == 0) {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\"></DIV>";
            }
            else {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:TargetVersion(" + ID + "," + VersionID + ",0)'\" >&nbsp;&nbsp;&nbsp;Remove&nbsp;Target&nbsp;</SPAN></FONT></DIV>";
            }

            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            if (FusionRequirements == 1) {
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeTargetNotes_Pulsar(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Target&nbsp;Notes...</SPAN></FONT></DIV>";
            }
            else
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeTargetNotes(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Target&nbsp;Notes...</SPAN></FONT></DIV>";

            if (strImage != "-") {
                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeDistribution(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Distribution...</SPAN></FONT></DIV>";
                if (FusionRequirements == 1) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangePulsarImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Images...</SPAN></FONT></DIV>";
                }
                else if (Fusion == 1) {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeIRSImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Images...</SPAN></FONT></DIV>";
                }
                else {
                    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                    popupBody = popupBody + "<FONT face=Arial size=2>";
                    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeImages(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Update&nbsp;Images...</SPAN></FONT></DIV>";
                }
            }

            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";


            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AdvancedTarget(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Advanced&nbsp;Target...</SPAN></FONT></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AlertDetails(" + ID + "," + VersionID + "," + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Alert&nbsp;Details...</SPAN></FONT></DIV>";

            if (txtISPreinstall.value == "1") {

                popupBody = popupBody + "<DIV>";
                popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";


                popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
                popupBody = popupBody + "<FONT face=Arial size=2>";
                popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayPIProperties(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Preinstall&nbsp;Properties&nbsp;</SPAN></FONT></DIV>";

            }


            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</SPAN></FONT></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + ID + "," + RootID + "," + VersionID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;History</SPAN></FONT></DIV>";

            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + ID + "," + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></FONT></DIV>";

            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<FONT face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReloadWindow()'\" >&nbsp;&nbsp;&nbsp;Refresh the Grid </SPAN></FONT></DIV>";

            popupBody = popupBody + "</DIV>";

            oPopup.document.body.innerHTML = popupBody;

            oPopup.show(lefter, topper, 170, 150, document.body);


            //Adjust window size
            if (oPopup.document.body.scrollHeight > 1 || oPopup.document.body.scrollWidth > 1) {
                NewHeight = oPopup.document.body.scrollHeight;
                NewWidth = oPopup.document.body.scrollWidth;
                if (topper + NewHeight > document.body.clientHeight)
                    topper = document.body.clientHeight - NewHeight;
                if (topper < 0)
                    topper = 0;
                oPopup.show(lefter, topper, NewWidth, NewHeight + 1, document.body);
            }

        }
    }

    function BatchUpdateRestriction(RootID, VersionID, ProdID, ReleaseID, FusionRequirements) {
        modalDialog.open({ dialogTitle: 'Batch Update Restriction', dialogURL: 'deliverable/Restrict/RestrictPulsar.asp?ProductID=' + ProdID + '&VersionID=' + VersionID + '&RootID=' + RootID + '&ReleaseID=' + ReleaseID, dialogHeight: $(window).height() * (60 / 100), dialogWidth: $(window).width() * (60 / 100), dialogResizable: true, dialogDraggable: true });
    }

    function DisplayDeliverableHistory(ProductID, RootID, VersionID, TypeID) {
        window.open("Image/DeliverableChanges.asp?ProductID=" + ProductID + "&RootID=" + RootID + "&VersionID=" + VersionID + "&ActionID=&UserID=&TypeID=" + TypeID);
    }

    function GetVersion(VersionID) {
        var strPath = trim(document.all("Path" + VersionID).value);
        //	window.open ("file://" + strPath);
        window.open("FileBrowse.asp?ID=" + VersionID);
        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }
    }

    //
    // Deliverables Tab
    //
    function ChangeTargetNotes(ProdID, VerID, RootID) {
        modalDialog.open({ dialogTitle: 'Change Target Notes', dialogURL: 'Target/EditExceptions.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID, dialogHeight: 400, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });

        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }
        //save Product ID and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
    }

    function ChangeTargetNotes_Pulsar(ProdID, VerID, RootID) {
        modalDialog.open({ dialogTitle: 'Change Target Notes', dialogURL: 'Target/Pulsar/EditExceptions.asp?ProductID=' + ProdID + '&VersionID=' + VerID + '&RootID=' + RootID, dialogHeight: 400, dialogWidth: 800, dialogResizable: true, dialogDraggable: true });

        if (typeof (SelectedRow) != "undefined") {
            if (SelectedRow != null) {
                SelectedRow.style.color = "black";
                SelectedRow = null;
            }
        }
        //save Product ID and Version ID for results function: ---
        globalVariable.save(ProdID, 'product_id');
        globalVariable.save(VerID, 'version_id');
        globalVariable.save(RootID, 'root_id');
    }

    function ChangeTargetNotesResult(strResult) {
        var iVersionID = globalVariable.get('version_id');
        var iProductID = globalVariable.get('product_id');

        if (typeof (strResult) != "undefined") {
            document.all("NoteCell" + iVersionID + "_" + iProductID).innerText = strResult;
        }
    }

    function SetNoteExists(NotesExists, TargetNotes) {
        document.getElementById("bNoteExists").value = NotesExists;
        document.getElementById("TargetNotes").value = TargetNotes;
    }

    function ChangeTargetNotesResult_Pulsar() {
        var NoteExists = document.getElementById("bNoteExists").value;
        var iVersionID = globalVariable.get('version_id');
        var iProductID = globalVariable.get('product_id');
        var iRootID = globalVariable.get('root_id');
        if (NoteExists == "1") {
            document.getElementById("aNotes" + iVersionID + "_" + iProductID).style.display = "inline";
        }
        else {
            document.getElementById("aNotes" + iVersionID + "_" + iProductID).style.display = "none";
        }
    }

    function ViewReleaseTargetNotes(ProductID, RootID, VersionID) {
        //modalDialog.open({ dialogTitle: 'View Target Notes', dialogURL: 'Target/TargetReleaseEdit.asp?ProductID=' + ProductID + '&RootID=' + RootID + '&VersionID=' + VersionID + "&ViewOnly=1", dialogHeight: 400, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
        modalDialog.open({ dialogTitle: 'View Target Notes', dialogURL: 'Target/TargetReleaseEdit.asp?ProductID=' + ProductID + '&RootID=' + RootID + '&VersionID=' + VersionID, dialogHeight: 400, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function MultiUpdateTestStatus(ProdID, RootID, VersionID, ReleaseID, FusionRequirements, BSID, ShowOnlyTargetedRelease) {
        var strTitle = "Batch Update Root Status";
        if (VersionID > 0) {
            strTitle = "Batch Update Product Status";
        }

        if (FusionRequirements == 1) {
            modalDialog.open({ dialogTitle: strTitle, dialogURL: 'deliverable/commodity/MultiTestStatusPulsar.asp?ProdID=' + ProdID + '&VersionList=' + VersionID + "&RootID=" + RootID + '&ReleaseID=' + ReleaseID + '&FusionRequirements=' + FusionRequirements + '&BSID=' + BSID + '&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
        }
        else {
            modalDialog.open({ dialogTitle: strTitle, dialogURL: 'deliverable/commodity/MultiTestStatus.asp?ProdID=' + ProdID + '&VersionList=' + VersionID + "&RootID=" + RootID, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
        }

    }

    function MultiUpdateTestLeadStatusLink(ProdID) {
        var i;
        var strIDList = "";
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2;
        NewTop = (screen.height - 650) / 2;

        if (typeof (chkVersion) == "undefined") {
            alert("There are no versions on this display that can be updated.");
            return;
        }
        if (typeof (chkVersion.length) == "undefined") {
            if (chkVersion.checked)
                strIDList = chkVersion.value;
        }
        else {
            for (i = 0; i < chkVersion.length; i++)
                if (chkVersion(i).checked)
                    strIDList = strIDList + "," + chkVersion(i).value;
            if (strIDList != "")
                strIDList = strIDList.substr(1);
        }

        if (strIDList.length > 0) {
            modalDialog.open({ dialogTitle: 'Batch Update Test Status', dialogURL: 'deliverable/TestTeam/MultiUpdateTestStatus.asp?ProductID=' + ProdID + '&IDList=' + strIDList, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
        }
        else
            alert("You must select at least one version to edit.");
    }

    function MultiUpdateTestStatusLink(ProdID, ReleaseID, FusionRequirements, ShowOnlyTargetedRelease) {
        var i;
        var strIDList = "";
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2;
        NewTop = (screen.height - 650) / 2;

        if (typeof (chkVersion) == "undefined") {
            alert("There are no versions on this display that can be updated.");
            return;
        }
        if (typeof (chkVersion.length) == "undefined") {
            if (chkVersion.checked)
                strIDList = chkVersion.value;
        }
        else {
            for (i = 0; i < chkVersion.length; i++)
                if (chkVersion(i).checked)
                    strIDList = strIDList + "," + chkVersion(i).value;
            if (strIDList != "")
                strIDList = strIDList.substr(1);
        }

        if (strIDList.length > 0) {
            if (FusionRequirements == 1) {
                modalDialog.open({ dialogTitle: 'Batch Edit Qualification Status', dialogURL: 'deliverable/commodity/MultiTestStatusPulsar.asp?ProdID=' + ProdID + '&VersionList=' + strIDList + '&ReleaseID=' + ReleaseID + '&FusionRequirements=' + FusionRequirements + '&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
            }
            else {
                modalDialog.open({ dialogTitle: 'Batch Edit Qualification Status', dialogURL: 'deliverable/commodity/MultiTestStatus.asp?ProdID=' + ProdID + '&VersionList=' + strIDList, dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
            }
        }
        else
            alert("You must select at least one version to edit.");
    }

    function AlertDetails(ProdID, VersionID, RootID) {
        window.open("AlertDetails.asp?ProdID=" + ProdID + "&RootID=" + RootID + "&VersionID=" + VersionID, "_blank");
    }

    //
    // Deliverables Tab : Commodities
    //
    function EditSubassembly(ID, RootID, VersionID, ReleaseID, FusionRequirements, ShowOnlyTargetedRelease) {
        if (FusionRequirements == 0)
            modalDialog.open({ dialogTitle: 'Edit Subassembly', dialogURL: 'deliverable/commodity/SubAssembly.asp?ProductID=' + ID + '&VersionID=' + VersionID + '&RootID=' + RootID, dialogHeight: $(window).height() * (55 / 100), dialogWidth: $(window).width() * (50 / 100), dialogResizable: true, dialogDraggable: true });
        else
            modalDialog.open({ dialogTitle: 'Edit Subassembly', dialogURL: 'deliverable/commodity/SubAssemblyPulsar.asp?ProductID=' + ID + '&VersionID=' + VersionID + '&RootID=' + RootID + '&ReleaseID=' + ReleaseID + '&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease, dialogHeight: $(window).height() * (55 / 100), dialogWidth: $(window).width() * (50 / 100), dialogResizable: true, dialogDraggable: true });
    }

    function ShowSICommodities(strID) {
        var strResult;
        strResult = window.showModalDialog("Deliverable/Commodity/SICommodities.asp?ID=" + strID, "", "dialogWidth:900px;dialogHeight:800px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
    }

    function EditPilotStatus(ProdID, VersionID, ReleaseID, FusionRequirements, ShowOnlyTargetedRelease) {
        if (FusionRequirements == 0)
            modalDialog.open({ dialogTitle: 'Edit Pilot Status', dialogURL: 'deliverable/commodity/PilotStatus.asp?ProdID=' + ProdID + '&VersionID=' + VersionID, dialogHeight: $(window).height() * (55 / 100), dialogWidth: $(window).width() * (50 / 100), dialogResizable: true, dialogDraggable: true });
        else
            modalDialog.open({ dialogTitle: 'Edit Pilot Status', dialogURL: 'deliverable/commodity/PilotStatusPulsar.asp?ProdID=' + ProdID + '&VersionID=' + VersionID + '&ReleaseID=' + ReleaseID + '&ShowOnlyOneRelease=0&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease, dialogHeight: $(window).height() * (55 / 100), dialogWidth: $(window).width() * (50 / 100), dialogResizable: true, dialogDraggable: true });
    }

    function EditPilotStatusResult(strStatus) {
        if (typeof (strStatus) != "undefined" && strStatus != "") {
            RefreshStatus("PilotStatusCell", strStatus);
        }
        closeModalDialog(false);
    }

    function EditAccessoryStatus(ProdID, VersionID, ReleaseID, FusionRequirements, ShowOnlyTargetedRelease) {
        if (FusionRequirements == 1) {
            modalDialog.open({ dialogTitle: 'Edit Accessory Status', dialogURL: 'deliverable/commodity/AccessoryStatusPulsar.asp?ProdID=' + ProdID + '&VersionID=' + VersionID + '&ReleaseID=' + ReleaseID + '&ShowOnlyOneRelease=0&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease + "&RowID=" + VersionID, dialogHeight: $(window).height() * (80 / 100), dialogWidth: $(window).width() * (70 / 100), dialogResizable: true, dialogDraggable: true });
        }
        else {
            modalDialog.open({ dialogTitle: 'Edit Accessory Status', dialogURL: 'deliverable/commodity/AccessoryStatus.asp?ProdID=' + ProdID + '&VersionID=' + VersionID, dialogHeight: $(window).height() * (80 / 100), dialogWidth: $(window).width() * (70 / 100), dialogResizable: true, dialogDraggable: true });
        }
    }

    function EditAccessoryStatusResult(VersionID, strStatus) {
        if (typeof (strStatus) != "undefined" && strStatus != "") {
            $("#AccessoryStatusCell" + VersionID).html(strStatus);
        }
        closeModalDialog(false);
    }

    //need jquery dialog
    function EditCommodityStatus(ProdID, VersionID, ReleaseID, FusionRequirements, ShowOnlyTargetedRelease) {
        if (FusionRequirements == 1) {
            modalDialog.open({ dialogTitle: 'Edit Qualification Status', dialogURL: 'deliverable/commodity/QualStatusPulsar.asp?ProdID=' + ProdID + '&VersionID=' + VersionID + '&ReleaseID=' + ReleaseID + '&ShowOnlyOneRelease=0&ShowOnlyTargetedRelease=' + ShowOnlyTargetedRelease, dialogHeight: $(window).height() * (80 / 100), dialogWidth: $(window).width() * (70 / 100), dialogResizable: true, dialogDraggable: true });
        }
        else {
            modalDialog.open({ dialogTitle: 'Edit Qualification Status', dialogURL: 'deliverable/commodity/QualStatus.asp?ProdID=' + ProdID + '&VersionID=' + VersionID, dialogHeight: $(window).height() * (80 / 100), dialogWidth: $(window).width() * (70 / 100), dialogResizable: true, dialogDraggable: true });
        }
    }

    function EditCommodityStatusResult(VersionID, ProductDeliverableID, ProductDeliverableReleaseID, strStatus) {
        var blnTargetDisplay = false;
        var blnTargeted;
        if (ProductDeliverableReleaseID == 0) {
            if (typeof (strStatus) != "undefined" && strStatus != "") {
                if (document.all("HWTargetCell" + VersionID) == null)
                    blnTargetDisplay = true;

                if (strStatus.length == 1 || strStatus == "RSTD:Investigating" || strStatus == "UNRS:Investigating")
                    blnTargeted = false;
                else
                    blnTargeted = true;

                if (blnTargetDisplay) {
                    // window.location.reload();
                }
                else {
                    if (blnTargeted)
                        document.all("HWTargetCell" + VersionID).innerText = "Yes";
                    else
                        document.all("HWTargetCell" + VersionID).innerText = " ";
                }

                if (strStatus.length == 1) {
                    document.all("HWStatusCell" + VersionID).innerText = "Not Used";
                }
                else if (strStatus.substring(0, 5) == "Con2:" || strStatus.substring(0, 5) == "Con3:" || strStatus.substring(0, 5) == "RSTD:" || strStatus.substring(0, 5) == "UNRS:") {
                    document.all("HWStatusCell" + VersionID).innerText = strStatus.substr(5);
                }
                else {
                    document.all("HWStatusCell" + VersionID).innerText = strStatus;
                }

                if (strStatus.substring(0, 5) == "RSTD:") {
                    document.all("RestrictedCell" + VersionID).innerText = "Yes";
                }
                else if (strStatus.substring(0, 5) == "UNRS:") {
                    document.all("RestrictedCell" + VersionID).innerHTML = "&nbsp;";
                }
            }

            closeModalDialog(false);
        }
        else {
            if (strStatus != "") {
                var arr = strStatus.split('|');

                if (arr[1] == "True")
                    $("#HWTargetCell" + arr[0]).html("Yes");
                else
                    $("#HWTargetCell" + arr[0]).html("&nbsp;");

                $("#RestrictedCell" + arr[0]).html(arr[2]);
                $("#HWStatusCell" + arr[0]).html(arr[3]);
            }

            closeModalDialog(false);
        }
    }

    //
    // All Tabs
    //
    function HeaderMouseOver() {
        window.event.srcElement.style.cursor = "hand";
        window.event.srcElement.style.color = "red";
    }

    function HeaderMouseOut() {
        window.event.srcElement.style.color = "black";
    }

    function mySettingSetCallback(returnstring) {
    }

    //
    // Image Tab Functions
    //
    function ShowImageCompare(ProdID, PINTest) {
        strResult = window.open("Image/CompareImagesChooseProduct.asp?ProdID=" + ProdID + "&PINTest=" + PINTest, "_blank", "width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function ShowImageCompareFusion(ProdID, PINTest) {
        var WinTop = Math.floor((screen.height - 500) / 4);
        strResult = window.open("Image/Fusion/CompareFusionImage.asp?ProductID=" + ProdID + "&PINTest=" + PINTest, "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function ShowImageCompare_Pulsar(ProdID, PINTest) {
        var WinTop = Math.floor((screen.height - 500) / 4);
        strResult = window.open("Image/Fusion/CompareImage_Pulsar.asp?ProductID=" + ProdID + "&PINTest=" + PINTest, "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function SyncIRSImages(ProdID) {
        var WinTop = Math.floor((screen.height - 500) / 4);
        strResult = window.open("/Pulsar/Product/SelectImageUpdates?pvId=" + ProdID + "&isIRSPD=true", "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function SyncIRSImages_Pulsar(ProdID) {
        var WinTop = Math.floor((screen.height - 500) / 4);
        strResult = window.open("/Pulsar/Product/SelectImageUpdates?pvId=" + ProdID + "&isIRSPD=true", "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function ExtendEOL(prodID, isPulsarProduct) {
        var WinTop = Math.floor((screen.height - 500) / 4);
        strResult = window.open("/Pulsar/Product/MultiMLEOLExtension?pvId=" + prodID + "&isPulsarProduct=" + isPulsarProduct, "_blank", "top=" + WinTop + ",width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }

    function TogglePin(PinID, PinState, strValue, CurrentUserID) {

        jsrsExecute("DefaultProductFilterRSUpdate.asp", mySettingSetCallback, "UpdateSetting", Array(strValue, String(CurrentUserID), "1"));

        if (PinID == 1) {
            if (PinState == 0) {
                DelPIN0.style.display = "none";
                DelPIN1.style.display = "";
            }
            else {
                DelPIN0.style.display = "";
                DelPIN1.style.display = "none";
            }
        }
    }

    //
    // General Tab
    //
    function OpenStatusOptions() {
        modalDialog.open({ dialogTitle: 'Change Log', dialogURL: 'ProductStatusOptions.asp?ID=' + txtID.value + '', dialogHeight: 600, dialogWidth: 350, dialogResizable: true, dialogDraggable: true });
    }

    //
    // General Tab : WHQL Functions
    //
    function RunWhqlWizard(PVID) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2;
        NewTop = (screen.height - 650) / 2;

        strResult = window.open("whql/whqlIdFrame.asp?PVID=" + PVID, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,resizable=yes,menubar=no,toolbar=no,status=no");
        window.location.reload(true);
    }

    function RunWhqlEditWizard(WHQLID) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2;
        NewTop = (screen.height - 650) / 2;

        ShowPropertiesDialog("whql/whqlIdEdit.aspx?WHQLID=" + WHQLID, "WHQL Status", 655, 650);
    }

    function ShowWhqlStatus(PVID) {
        var strResult;
        var Width = 800;
        var Height = 600;
        var NewTop = (screen.height - 600) / 2;
        var NewLeft = (screen.width - 800) / 2;

        ShowPropertiesDialog("whql/whqlStatus.aspx?PVID=" + PVID, "WHQL Status", Width, Height);
    }

    function LeverageWhqlStatus(PVID) {
        var strResult;
        var Width = 450;
        var Height = 575;
        var NewTop = (screen.height - 600) / 2;
        var NewLeft = (screen.width - 800) / 2;

        strResult = window.open("whql/leverageStatus.aspx?PVID=" + PVID, "_blank",
            "Left=" + NewLeft + ",Top=" + NewTop + ",Width=" + Width + ",Height=" + Height +
            ",resizable=yes,menubar=no,toolbar=no,status=no,scrollbars=no");
    }

    function RTMDocClick(strID) {
        var strID;
        strResult = window.open("Product/MilestoneSignoffReport.asp?ID=" + strID, "_blank", "width=900, height=500,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");
    }
    function UpdateRTMDraft(strID, RTMID) {
        modalDialog.open({ dialogTitle: 'Update Product RTM', dialogURL: 'Product/MilestoneSignoff.asp?ID=' + strID + '&ProductRTMID=' + RTMID + '', dialogHeight: 600, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }
    function CreateRTM(strID) {
        modalDialog.open({ dialogTitle: 'Product RTM', dialogURL: 'Product/MilestoneSignoff.asp?ID=' + strID + '&ProductRTMID=0', dialogHeight: 600, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function CreateMRTM(strID) {
        modalDialog.open({ dialogTitle: 'Multiple Products RTM', dialogURL: 'Product/MultipleRTMProgram.asp?ID=' + strID + '', dialogHeight: 600, dialogWidth: 750, dialogResizable: true, dialogDraggable: true });
    }

    function ShowReadinessReportOptions(ProdID, ReportType, TeamID) {
        if (ReportType == 3 && TeamID != 0)
            strResult = window.open("ReadinessReportOptions.asp?ProdID=" + ProdID + "&ReportType=" + ReportType + "&TeamID=" + TeamID, "_blank", "width=500, height=500,location=no, menubar=no, status=no,toolbar=no, scrollbars=yes, resizable=yes");
        else if (ReportType == 3)
            strResult = window.open("ReadinessReportOptions.asp?ProdID=" + ProdID + "&ReportType=" + ReportType, "_blank", "width=500, height=500,location=no, menubar=no, status=no,toolbar=no, scrollbars=yes, resizable=yes");
        else
            strResult = window.open("ReadinessReportOptions.asp?ProdID=" + ProdID, "_blank", "width=500, height=500,location=no, menubar=no, status=no,toolbar=no, scrollbars=yes, resizable=yes");
    }

    function ShowBatchUpdateComponentToProductRelease(url) {
        var popupheight = $(window).height() - 60;
        var popupWidth = $(window).width() - 60;
        /*modalDialog.open({ dialogTitle: 'Batch Update to Add Supported Releases to Components', dialogURL: url, dialogHeight: popupheight, dialogWidth: popupWidth, dialogResizable: true, dialogDraggable: true });*/
        //Batch Update to Add Supported Releases to Components
        window.showModalDialog(url, "", " dialogWidth:" + popupWidth + "px; dialogHeight:" + popupheight + "px; center:Yes; help:No; maximize:no; resizable:no; status:No");
    }

    function ImportTargetingSettings(url) {
        var popupheight = $(window).height() - 60;
        var popupWidth = $(window).width() - 60;
        /*modalDialog.open({ dialogTitle: 'Import Targeting Settings', dialogURL: url, dialogHeight: popupheight, dialogWidth: popupWidth, dialogResizable: true, dialogDraggable: true });*/
        //Import Targeting Settings
        window.showModalDialog(url, "", " dialogWidth:" + popupWidth + "px; dialogHeight:" + popupheight + "px; center:Yes; help:No; maximize:no; resizable:no; status:No");

    }

    function ImportOSImageSettings(url) {
        var popupheight = $(window).height() - 60;
        var popupWidth = $(window).width() - 60;
        /*modalDialog.open({ dialogTitle: 'Import Supported Image Settings', dialogURL: url, dialogHeight: popupheight, dialogWidth: popupWidth, dialogResizable: true, dialogDraggable: true });*/
        //Import Supported Image Settings
        window.showModalDialog(url, "", " dialogWidth:" + popupWidth + "px; dialogHeight:" + popupheight + "px; center:Yes; help:No; maximize:no; resizable:no; status:No");
    }

    function PublishMarketingRequirements(PVID) {
        var strID
        strID = window.parent.showModalDialog("<%= AppRoot %>/PublishMarketingReqFrame.asp?PVID=" + PVID, "", "dialogWidth:650px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;")
        window.location.reload();
    }

    function LaunchProjectExplorer(ID) {
        var WinTop = Math.floor((screen.height - 470) / 2);
        var url = window.location.host;

        window.open("http://<%=Application("IRS_WebServerName")%>/irs/irsplus/default.aspx?link=projectmgmt/ProjectMgmt.aspx?PDID=" + ID, "_blank", "top=" + WinTop + ",height=570,width=1000, location=0,menubar=0,resizable=0,scrollbars=0,status=0,titlebar=0,toolbar=0");

    }

    function AddEditMarketingName(BID, Name, NameType, GeneratedName, PBID, Series) {
        modalDialog.open({ dialogTitle: 'Edit Name', dialogURL: '<%= AppRoot %>/AddEditMarketingNameFrame.asp?BID=' + BID + '&Name=' + Name + '&NameType=' + NameType + '&GeneratedName=' + GeneratedName + '&PBID=' + PBID + '&Series=' + Series, dialogHeight: 400, dialogWidth: 370, dialogResizable: true, dialogDraggable: true });
    }

    function autoResize(id) {
        var newheight;
        var newwidth;

        if (document.getElementById) {
            newheight = document.getElementById(id).contentWindow.document.body.scrollHeight;
            newwidth = document.getElementById(id).contentWindow.document.body.scrollWidth;
        }
        if (newheight != 0)
            document.getElementById(id).height = String(newheight) + "px";//"1064px";//;
        // document.getElementById(id).width = (newwidth) + "px"; //"2000px" //
    }

    //*****************************************************************
    //Description:  Use in parent page; closes modal dialog opend with modalDialog code
    //*****************************************************************
    function closeModalDialog(bReload) {
        if (bReload) {
            modalDialog.cancel(bReload);
        }
        else {
            $("#btnRefresh").removeClass("hide").addClass("show");
            modalDialog.cancel(false);
        }
    };

    function RefreshStatus(cell, newStatus) {
        var arr = newStatus.split('|');
        $("#" + cell + arr[0]).html(arr[1]);
        modalDialog.cancel(false);
    }

    //*****************************************************************
    //Description:  Display Edit Base Unit Group form in a modal dialog
    //Function:     UpdatePlatform();
    //Modified:     Harris, Valerie (9/28/2016) - PBI 26986/Task 27006
    //*****************************************************************
    function UpdatePlatform(ID, ProductVersionID, followMKTName) {
        modalDialog.open({ dialogTitle: 'Update Base Unit Group', dialogURL: 'MobileSE/Today/platform.asp?ID=' + ID + '&ProductVersionID=' + ProductVersionID + '&FollowMKTName=' + followMKTName, dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
    }

    //*****************************************************************
    //Description:  Display Add Base Unit Group form in a modal dialog
    //Function:     AddPlatform();
    //Modified:     Harris, Valerie (9/28/2016) - PBI 26986/Task 27006
    //*****************************************************************
    function AddPlatform(ProductVersionID, followMKTName) {
        modalDialog.open({ dialogTitle: 'Add Base Unit Group', dialogURL: 'MobileSE/Today/platform.asp?ProductVersionID=' + ProductVersionID + '&FollowMKTName=' + followMKTName, dialogHeight: 650, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });
    }

    function SaveSettingForFunExclude(OldSetting) {
        var ProductVersionID = $("#txtID").val(), IsExclude = true;

        if (OldSetting == "Exclude")
            IsExclude = false;

        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Common/SaveExcludeIncompleteWorkflowComponent",
            data: "{ ProductVersionID: " + ProductVersionID + ", IsExclude: " + IsExclude + "}",
            dataType: "json",
            async: false,
            beforeSend: function () {
                $("#divFunCompExclude").spin('small', '#0096D6');
            },
            complete: function () {
                $("#divFunCompExclude").spin(false);
            },
            error: function (msg, status, error) {
                $("#divFunCompExclude").spin(false);
                var errMsg = $.parseJSON(msg.responseText);
                alert("Error when saving setting. Please try again later." + errMsg);
            }

        });

        if (IsExclude)
            document.getElementById('divFunCompExclude').innerHTML = "<a href=javascript:SaveSettingForFunExclude('Exclude')>Exclude Functional Test Components</a>";
        else
            document.getElementById('divFunCompExclude').innerHTML = "<a href=javascript:SaveSettingForFunExclude('Include')>Include Functional Test Components</a>";

    }

    function OpenKeyboardMatrix(ProductVersionID) {
        window.showModalDialog('/Pulsar/Product/ProductPRLList?ProductVersionID=' + ProductVersionID, window, "dialogWidth:1000px;dialogHeight:500px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        //modalDialog.open({ dialogTitle: 'Product PRL List', dialogURL: '/Pulsar/Product/ProductPRLList?ProductVersionID=' + ProductVersionID, dialogHeight: 500, dialogWidth: 1000, dialogResizable: true, dialogDraggable: true });       
    }

    function OpenChangeHistory(ProductVersionID) {
        window.open("/Pulsar/Product/ImageChangeHistory?ProductVersionID=" + ProductVersionID, "_blank", "resizable=1,scrollbars=1,menubar=1,toolbar=1,dependent=1", false);
    }

    window.onresize = resizeIframe;
    function resizeIframe() {
        $('#PulsarproductRequirement').width($(window).width() - 20);
        $('#PulsarproductRequirement').height($(window).height() - 200);
    }

    function UpdateRTPStatus(Id) {
        strResult = window.showModalDialog('/Pulsar/Product/GetUpdateRTPStatus?Id=' + Id, "", "dialogWidth:700px;dialogHeight:645px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strResult) != "undefined") {
            window.location.reload(true);
        }
    }

//-->
</script>
</head>
<body onLoad="return window_onload()">
<div style="display:none"><a href="UpdateUserAccess.asp"></a></div>
<%
dim ColumnCount
dim strDcrStatus
dim strBiosChange
dim strSwChange
dim strCookie
on error resume next
strCookie = ""
strCookie = Request.Cookies("DCRFilterType")
on error goto 0

SELECT CASE strCookie
    CASE "all"
        strBiosChange = "NULL"
        strSwChange = "NULL"
    CASE "dcr"
        strBiosChange = "0"
        strSwChange = "0"
    CASE "bcr"
        strBiosChange = "1"
        strSwChange = "0"
    CASE "scr"
        strBiosChange = "0"
        strSwChange = "1"
    CASE ELSE
        strBiosChange = "NULL"
        strSwChange = "NULL"
END SELECT

strBiosChange = "NULL"
strSwChange = "NULL"

on error resume next
strCookie = ""
strCookie = Request.Cookies("DCRFilterStatus")
on error goto 0
SELECT CASE strCookie
    CASE "all"
      strDcrStatus = 0
    CASE "open"
      strDcrStatus = 1
    CASE "closed"
      strDcrStatus = 2
    CASE ELSE
    	strDcrStatus = 1
END SELECT

%>
<input type="hidden" id="hidDCRFilterType" name="hidDCRFilterType" value="<%= strBiosChange %>" />
<input type="hidden" id="hidDCRFilterScr" name="hidDCRFilterScr" value="<%= strSwChange %>" />
<input type="hidden" id="hidDCRFilterStatus" name="hidDCRFilterStatus" value="<%= strDcrStatus %>" />
<span id="loadingProgress"></span>
<span id="productNameTitle" style="font:bold medium Verdana;">
<%
    dim rtpStatus
    dim typeId
	dim strID
	dim strDescription
	dim strNotes
	dim strManager
	dim strManagerEmail
	dim strSMName
	dim strSMEmail
	dim strSEPMName
	dim strSEPMEmail
	dim strCategroy
	dim strVendor
	dim strPart
	dim strStatusText
	dim strStatusID
	dim strTitleColor
	dim strAvailableForTest
	dim strDisplayedList
	dim blnAdministrator
	dim blnMarketingAdmin
	dim ItemsDisplayed
	dim strSWRow
	dim LastRequirement
	dim strFamilyID
	dim strProductName
	dim blnPreinstallPM
	dim blnCommodityPM
	dim blnTestLead
	dim blnSuperUser
    dim blnPilotEngineer
	dim strPreinstallTeam
	dim strReleaseTeam
	dim strRegulatoryModel
	dim strMinRoHSLevel
    dim strProductReleases
	dim blnPddLocked
	dim strCommodityPM
	dim blnSysAdmin
	dim blnWhqlTeam
	dim blnSETestLead
	dim blnODMTestLead
	dim blnWWANTestLead
	dim blnDEVTestLead
	dim strServiceLifeDate
	dim strEndOfProductionDate
	dim blnCommPM
	dim blnPlatformDevelopmentPM
    dim blnODMHWPM
    dim blnHWPC
	dim blnProcessorPM
	dim blnVideoMemoryPM
	dim blnGraphicsControllerPM
	dim blnHardwarePM
	dim blnServicePM
	dim ServicePMAccess
	dim blnToolsPM
	dim blnActionOwner
	dim strToolAccessList
	dim PreiodType ''1=month, 2=week, 3=quarter
    dim strRCTOSites
    dim strProductGroups
    dim strFusion
    dim strImageTool
    dim blnSupplyChain
    dim strFusionRequirements
    dim strFactoryName
    dim strProductLineName
    dim strBusinessSegmentName
    dim blnODMSEPM
    dim strNonPostPORPRLList
    dim blnOdmPreinstallPM
    dim intReleaseCount
    dim strLatestVer
    dim blnHWPMRole
    dim blnSWPMRole

    'Harris, Valerie -  02/29/2016 - PBI 17178/ Task 17281 - Declare variables  
    Dim bIsPulsarProduct        'Create boolean variable that sets Product's Pulsar/Legacy type
    Dim sComMarketingName
    Dim sConMarketingName
    Dim sSMBMarketingName
    Dim sPOManagerName
    Dim sConfigManagerName
    Dim sProcurementPMName
    Dim sODMSEPMName
    Dim sPlanningPMName
    Dim intReleaseID
    dim ProductImageEdit
    dim strBusinessSegmentId
    intReleaseID = 0
    strFusion=0
    strFusionRequirements = 0
	PeriodType = 1
	ItemsDisplayed = 0
	ServicePMAccess = 0
    strNonPostPORPRLList = ""
    intReleaseCount = 0

    strImageTool = ""
	strRCTOSites = ""
    strProductGroups = ""
	blnSysAdmin = false
	blnToolsPM = false
	blnActionOwner = false
	blnMarketingAdmin	= false
	blnAdministrator = false
	blnPreinstallPM = false
	blnHardwarePM = false
	blnCommodityPM = false
	blnTestLead = false
	blnSuperUser = false
	blnCommPM = false
	blnProcessorPM = false
	blnVideoMemoryPM = false
	blnGraphicsControllerPM = false
	blnServicePM =false
	blnPlatformDevelopmentPM = false
    blnODMHWPM = false
    blnHWPC = false
	blnAccessoryPM = false
	blnPilotEngineer = false
	blnWhqlTeam = false
	blnSETestLead = 0
	blnODMTestLead = 0
	blnWWANTestLead = 0
	blnDEVTestLead = 0
	strToolAccessList = ""
	blnSupplyChain = false
    blnODMSEPM = false
    blnOdmPreinstallPM = false
    blnHWPMRole = false
    blnSWPMRole = false
    
	strServiceLifeDate = ""
	strEndOfProductionDate = ""
    strFactoryName = ""

	regEx.Pattern = "[^0-9a-fA-F#]"

	strTitleColor = "#0000cd"
	on error resume next
	strTitleColor = regEx.Replace(Request.Cookies("TitleColor"), "")
	if strTitleColor = "" then
		strTitleColor = "#0000cd"
	end if
	on error goto 0

	strID = trim(PVID)

	if strID <> "" and isnumeric(strID) then
		dim rs
		dim rs2
        dim rs3
        dim rs4
        dim rsOTS
		dim cn
		dim cm
		dim p

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.CommandTimeout =120
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")
        set rs3 = server.CreateObject("ADODB.recordset")
        set rs4 = server.CreateObject("ADODB.recordset")
        set rsOTS = server.CreateObject("ADODB.recordset")

		rs.Open "usp_GetPddLockStatus " & PVID, cn
		If Not rs.EOF Then
			blnPddLocked = rs("PddLocked")
		Else
			blnPddLocked = False
		End If
		rs.Close


		dim CurrentUser
		dim CurrentUserNTDomainName
		dim CurrentUserName
		dim CurrentUserID
		dim CurrentUserPartnerType
		dim CurrentUserSysAdmin
		dim CurrentWorkgroupID
		dim CurrentUserWorkgroup
		dim CurrentUserEmail
		dim CurrentUserPhone
		dim CurrentUserDefaultTab
		dim DisplayedProductName
		dim strSeries  ' 11/24/2016 Herb,  strSeries is not displayed/printed on the page, can someone be sure this variable is in use or useless?
        dim strSeries1
		dim strBrands
		dim strSEPMID
		dim strPMID
		dim strPORDate
		dim strOnlineReports
		dim strSystemboardComments
		dim strMachinePNPComments
		dim strSysBoardID
		dim strPnPID
		dim strProdType
		dim strDevCenter
		dim strFavs
		dim strFavCount
		dim strPartnername
		dim strPartnerID
		dim strMarketing
		dim strPlatformDevelopment
		dim strSupplyChain
		dim strService
        dim strODMHWPM
        dim strQuality
        dim strBiosLead
		dim strDevCenterName
		dim strImagePO
		dim strReferencePlatform
        dim strLeadProduct
		dim strROMVersion
		dim strROMWebVersion
		dim strOSPreinstall
		dim strDistributionList
		dim strOSWeb
		dim strOS
		dim strHardwarePMs
		dim blnProcurementEngineer
		dim strHardwareAccessGroup
		dim blnMITTestLead
		dim strDomainSite
		dim blnEngineeringCoordinator
		dim blnServiceCommodityManager
		dim intCMProductCount
        dim AgencyVersion : AgencyVersion = 1
        dim blnProductServiceManager : blnProductServiceManager = false
        dim blnCanEditProduct : blnCanEditProduct = false
        dim blnSEPMProducts : blnSEPMProducts = 0
        '------------ Product Properties --------------
        dim CreatedBy : CreatedBy = ""
        dim Created : Created = ""
        dim UpdatedBy : UpdatedBy = ""
        dim Updated : Updated = ""
        dim IsFunCompExclude : IsFunCompExclude = ""
        dim blnAgencyDataMaintainer : blnAgencyDataMaintainer = false
        dim FollowMktName : FollowMktName = 0

		blnProcurementEngineer = false
		blnServiceCommodityManager = false
		blnMITTestLead = false
		blnEngineeringCoordinator = false
		CurrentUserDefaultTab = ""
		strHardwarePMs = ""
		DisplayedProductName = ""
		serSeries = "&nbsp;"
		strBrands = "&nbsp;"
		strImagePO = "&nbsp;"
		strROMVersion = "&nbsp;"
		strROMWebVersion = "&nbsp;"
		strPreinstallTeam = ""
		strReleaseTeam = ""
		strRegulatoryModel = ""
		strMinRoHSLevel = ""
        strProductReleases = ""
		strOSPreinstall = ""
		strDistributionList =""
		strOSWeb = ""
		strHardwareAccessGroup = 0
		intCMProductCount = 0

		'Get User
		dim CurrentDomain
		dim CurrentUserPartner
		dim SAAdmin

		CurrentUser = lcase(Session("LoggedInUser"))
		CurrentUserNTDomainName = CurrentUser

		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetUserInfo"


		Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		p.Value = Currentuser
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		p.Value = CurrentDomain
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute

		set cm=nothing

        CurrentUserPartnerType = 0
		if not (rs.EOF and rs.BOF) then
			CurrentUserName = rs("Name") & ""
			CurrentUserID = rs("ID") & ""
			CurrentUserSysAdmin = rs("SystemAdmin")
			CurrentWorkgroupID = rs("WorkgroupID") & ""
			CurrentUserPartner = trim(rs("PartnerID") & "")
			CurrentUserEmail = rs("Email") & ""
			CurrentUserPhone = rs("Phone") & ""
			CurrentUserWorkgroupID = rs("WorkgroupID") & ""
			blnPreinstallPM = rs("PreinstallPM")
			blnCommodityPM = rs("CommodityPM")
			blnServicePM =  rs("ServicePM")
			blnAccessoryPM = rs("AccessoryPM")
			blnPilotEngineer = rs("SCFactoryEngineer")
			blnProcurementEngineer = rs("ProcurementEngineer")
			blnMITTestLead = rs("MITTestLead")
			blnEngineeringCoordinator= rs("engcoordinator")
			CurrentUserDefaultTab = rs("DefaultProductTab") & ""
			strFavs = trim(rs("Favorites") & "")
			strFavCount = trim(rs("FavCount") & "")
			blnWhqlTeam = rs("WhqlTestTeam")
			blnServiceCommodityManager = rs("ServiceCommodityManager")
			CurrentUserPartnerType = rs("PartnerTypeID") & ""
			SAAdmin = rs("SAAdmin") & ""
			intCMProductCount = rs("CMProductCount") '''SCM Owner (SCMOwnerID) System role has the same permission as Configuration Manager(TDCCMID) and POPM (PMID), included the logic of Role_Object_Permission table.
            ProductImageEdit = rs("ProductImageEdit")
            'permission needed for Edit Product link on Pulsar products
            if rs("CanEditProduct") = "1" then
                blnCanEditProduct = true
            end if

            '10/21/16 - Harris, Valerie - PBI 24551 - permission needed for SEPM Products: ---
            if rs("SEPMProducts") = "1" then
                blnSEPMProducts = 1
            end if
            
            if lcase(trim(rs("domain"))) = "asiapacific" then
				strDomainSite = 2
			else
				strDomainSite = 1
			end if

            if rs("HWProducts") > 0 then 
                blnPlatformDevelopmentPM = true
            end if
            
            if rs("AgencyDataMaintainer") > 0 then
    		    blnAgencyDataMaintainer = true
    		end if
		end if
		rs.Close

		blnTestLead=false
		rs.open "spGetTestLeadsAll 1,1,1",cn,adOpenStatic
		do while not rs.EOF
			if trim(currentuserid) = trim(rs("ID")) then
				if rs("role") = "ODM Test Lead"  then
					blnODMTestLead = 1
				elseif rs("role") = "WWAN Test Lead"  then
					blnWWANTestLead = 1
        		elseif rs("role") = "DEV Test Lead"  then
					blnDEVTestLead = 1
				elseif rs("role") = "SE Test Lead" then
					blnSETestLead = 1
					if rs("PartnerID")  = 1 then
						blnODMTestLead = 1
					end if
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
		if blnODMTestLead =1 or blnWWANTestLead =1 or blnDEVTestLead =1 then 'blnSETestLead or
			blnTestLead = true
		end if

		rs.open "spGetHardwareTeamAccessList " & CurrentUserID & "," & clng(strID),cn,adOpenStatic
		do while not rs.EOF
			if rs("HWTeam") = "ProgramCoordinator" or blnEngineeringCoordinator > 0 then
				blnPlatformDevelopmentPM = true
			elseif rs("HWTeam") = "PlatformDevelopment" and rs("Products") > 0 then
				blnPlatformDevelopmentPM = true
			elseif rs("HWTeam") = "SupplyChain" and rs("Products") > 0 then
				blnPlatformDevelopmentPM = true
			elseif rs("HWTeam") = "Processor" and rs("Products") > 0 then
				blnProcessorPM = true
			elseif rs("HWTeam") = "Comm" and rs("Products") > 0 then
				blnCommPM = true
			elseif rs("HWTeam") = "Commodity" and rs("Products") > 0 then
				blnCommodityPM = true
			elseif rs("HWTeam") = "GraphicsController" and rs("Products") > 0 then
				blnGraphicsControllerPM = true
			elseif rs("HWTeam") = "VideoMemory" and rs("Products") > 0 then
				blnVideoMemoryPM = true
			elseif rs("HWTeam") = "SuperUser" and rs("Products") > 0 then
				blnSuperUser = true
			elseif rs("HWTeam") = "SupplyChain" and rs("Products") > 0 then
				blnSupplyChain = true
            elseif rs("HWTeam") = "ODMSEPM" and rs("Products") > 0 then
                blnODMSEPM = true
            elseif rs("HWTeam") = "ODMHWPM" and rs("Products") > 0 then
                blnODMHWPM = true
            elseif rs("HWTeam") = "HWPC" and rs("Products") > 0 then
                blnHWPC = true
            elseif rs("HWTeam") = "HWPMRole" and rs("Products") > 0 then
                blnHWPMRole = true
            elseif rs("HWTeam") = "SWPMRole" and rs("Products") > 0 then
                blnSWPMRole = true
			end if
			rs.MoveNext
		loop
		rs.Close

        if (blnSuperUser = false) then
            blnSuperUser = blnODMSEPM
        end if

		if blnServicePM then 'or blnPlatformDevelopmentPM then
			blnPlatformDevelopmentPM = true
			blnProcessorPM = true
			blnCommPM = true
			blnVideoMemoryPM = true
			blnGraphicsControllerPM = true
			blnCommodityPM = true
			ServicePMAccess = 1
	    else
	        ServicePMAccess =0
		end if

        'blnODMHWPM is odm hardware pm and blnHWPC is assistant for PM will do the same 
		if blnCommodityPM or  blnCommPM or blnProcessorPM or blnVideoMemoryPM or blnGraphicsControllerPM or blnODMHWPM or blnHWPC then
			blnHardwarePM = true
		else
			blnHardwarePM = false
		end if

		CurrentUserWorkgroup = ""
		if CurrentUserWorkgroupid <> ""and isnumeric(CurrentUserWorkgroupID) then
			rs.Open "spGetWorkgroup " & CurrentUserWorkgroupID,cn,adOpenKeyset
			if not(rs.EOF and rs.BOF) then
				CurrentUserWorkgroup = rs("Name") & ""
			end if
			rs.Close
		end if

		if blnSysAdmin then
			strHardwareAccessGroup = "1"
		elseif blnHardwarePM or blnSuperUser or blnPlatformDevelopmentPM  or blnHWPMRole then
			strHardwareAccessGroup = "2"
		elseif blnPilotEngineer then
			strHardwareAccessGroup = "3"
		elseif blnAccessoryPM then
			strHardwareAccessGroup = "4"
		elseif blnProcurementEngineer then
			strHardwareAccessGroup = "5"
		else
			strHardwareAccessGroup = "0"
		end if
		on error resume next
		strCookie = ""
		Dim intPVID
        intPVID = Request.Cookies("LastProductDisplayed")
        'Response.Write "before set: " + strCookie + "<br />"
		on error goto 0


        'response.Write "CookieID: " + strCookie + " - PVID: " + PVID + "<br />" 
        'if we open product for the first time from the left hand side, 
        'we are looking to see if user has CurrentUserDefaultTab
        'if not we will get from cookie
        'if we still don't have it we will let it go to "General" tab
        if clng(PVID) = 344 or clng(PVID) = 347 or clng(PVID) = 1107 then
            strDisplayedList = "DCR"
        elseif intPVID = PVID and Instr(Request.QueryString, Request.Cookies("PMTab")) > 0 then
            strDisplayedList = Request.Cookies("PMTab")

		elseif intPVID <> PVID and sList = "" then
		    if trim(CurrentUserDefaultTab) <> "" then
				strDisplayedList = CurrentUserDefaultTab
            else
				strDisplayedList = "General"
		    end if
        
		else       
		    on error resume next
			strCookie = ""

            if sList = "" then
                if trim(CurrentUserDefaultTab) <> "" then
				    strCookie = CurrentUserDefaultTab
                else
				    strCookie = Request.Cookies("PMTab")
		        end if
            else 
                strCookie = sList
            end if
			
			on error goto 0
			
            if trim(strCookie) <> "" then
			    regEx.Pattern = "[^0-9a-zA-Z_ ]"
				strDisplayedList = trim(regEx.Replace(strCookie, ""))
			else
				strDisplayedList = "General"                
			end if
        end if

        sList = strDisplayedList

	on error resume next
    
 	Response.Cookies("LastProductDisplayed") = PVID
 	Response.Cookies("PMTab") = strDisplayedList
    
 
	on error goto 0

		strHardwarePMs = ""
		rs.Open "spListHardwarePMsAll",cn,adOpenStatic
		do while not rs.EOF
			strHardwarePMs = strHardwarePMs & "," & trim(rs("ID"))
			rs.MoveNext
		loop
		rs.Close

        Dim strWhqlIDs
        Dim strWhqlEditIDs
        strWhqlIDs = ""
        strWhqlEditIDs = ""
        rs.Open "usp_ListWHQLSubmissions " & clng(strID),cn,adOpenStatic
        do until rs.EOF
            strWhqlIDs = strWhqlIDs & ", " & trim(rs("SubmissionID"))
            strWhqlEditIDs = strWhqlEditIDs & ", " & "<a href=""#"" onclick=""RunWhqlEditWizard(" & rs("ID") & ")"">" & trim(rs("SubmissionID")) & "</a>"
            rs.MoveNext
        loop
        rs.Close

        If Len(strWhqlIDs) > 2 Then
            strWhqlIDs = mid(strWhqlIDs, 2, Len(strWhqlIDs)-1)
            strWhqlEditIDs = mid(strWhqlEditIDs, 2, Len(strWhqlEditIDs)-1)
        End If

        Dim strWhqlStatus        
        strWhqlStatus = "Unknown"

        If Len(strWhqlIDs) = 0 Then
            strWhqlStatus = "Incomplete"
        End If

		dim ShowItem      
		if CurrentUserPartner = "1" then
			ShowItem = ""
		else
			ShowItem = "none"
		end if

		strMachinePNPComments = ""
		strSystemboardComments = ""

		strPORDate = ""
		strOnlineReports = 0
		strProdType = 1
		strDevCenter = ""
		strSQL = "spGetProductVersion " & clng(strID)
		rs.Open strSQL,cn,adOpenForwardOnly
		if (rs.EOF and rs.BOF) and strID <> "-1" then
			Response.Write "Unable to find the selected program.<br><font size=1>ID=" & PVID & "</font>"
			Response.Write "<BR><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & PVID & ")""><font face=verdana size=1>Remove From Favorites</font></a>"
			Response.Write "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
			Response.Write "<span id=EditLink style=""Display:none""></span><span id=StatusLink style=""Display:none""></span><span id=menubar style=""Display:none""></span><span ID=Wait style=""Display:none""></span>"
			Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""1"">"
		else
			Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""0"">"
			if strID <> "-1" then
			    if trim(currentuserid) = trim(rs("TDCCMID") & "") or trim(currentuserid) = trim(rs("SMID") & "") or trim(currentuserid) = trim(rs("PMID") & "") or trim(currentuserid) = trim(rs("SEPMID") & "") then
			        blnAdministrator = true
			    end if
                if rs("FusionRequirements") and pvid <> 1038 then
                    strImageTool = "IRS"
                else
                    'Harris,Valerie - Bug 22201/ Task 22901 - Default Image Tab to IRS 
                    if request("ImageTool") = "" then
                        strImageTool = "IRS"
                    else
                    strImageTool = request("ImageTool")
                    end if
                end if
                if rs("Fusion") then
                    strFusion = 1
                else
                    strFusion = 0
                end if
                if (rs("FusionRequirements")) then
                    strFusionRequirements = 1
                else
                    strFusionRequirements = 0  
                end if      
                'if trim(currentuserid) = trim(rs("ODMSEPM") & "")  then
                    'blnODMSEPM      
        '----------- Product Properties --------------
                Created = rs("Created") & ""
                CreatedBy = rs("CreatedBy") & ""
                Updated = rs("Updated") & ""
                UpdatedBy = rs("UpdatedBy") & ""

                AgencyVersion = rs("AgencyVersion") & ""
			    strRCTOSites = rs("RCTOSites") & "&nbsp;"
				strDescription = REPLACE(rs("Description") & "&nbsp;","?","""")
				strBrands = rs("Brands") & "&nbsp;"
				strSysBoardID = rs("SystemBoardID") & "&nbsp;"
				strPartnername = rs("Partner") & "&nbsp;"
				strMachinePNPComments = rs("MachinePNPComments") & "&nbsp;"
				strSystemboardComments = rs("SystemboardComments") & "&nbsp;"
				strPartnerID = rs("PartnerID") & ""
				strDistributionList = trim(rs("Distribution") & "")
				strPnPID = rs("MachinePNPID") & "&nbsp;"
				strManager = rs("PMName") & "&nbsp;"
				strManagerEmail = rs("PMEmail") & "&nbsp;"
                strProductLineName = rs("ProductLineName") & "&nbsp;"
                strBusinessSegmentName = rs("BusinessSegmentName") & "&nbsp;"
				strSMName = rs("SMName") & "&nbsp;"
				strSMEmail = rs("SMEmail") & "&nbsp;"
				strSEPMName = rs("SEPMName") & "&nbsp;"
				strSEPMEmail = rs("SEPMEmail") & "&nbsp;"
				strSEPMID = rs("SEPMID") & ""
				strROMVersion = rs("CurrentROM") & ""
				strROMWebVersion = trim(rs("CurrentWebROM") & "")
				strDevCenter = trim(rs("DevCenter") & "")
				strImagePO = replace(rs("ImagePO") & "",vbcrlf,"<BR>") & "&nbsp;"
				strFamilyID = rs("ProductFamilyID") & ""
				strProdType = rs("TypeID") & ""
				strServiceLifeDate = rs("ServiceLifeDate") & ""
				strRegulatoryModel = trim(rs("RegulatoryModel") & "")
				strMinRoHSLevel = trim(rs("MinRoHSLevelName") & "")
                strProductReleases = trim(rs("ProductRelease") & "")
				strToolAccessList = replace(rs("ToolAccessList") & ""," ","")
				strCallDataLastUpdated = trim(rs("CallDataLastUpdated") & "")
                strFactoryName = trim(rs("FactoryName") & "")
				if rs("PreinstallTeam") & "" = "1" then
					strPreinstallTeam = "Houston"
				elseif rs("PreinstallTeam") & "" = "2" then
					strPreinstallTeam = "Taiwan"
				elseif rs("PreinstallTeam") & "" = "3" then
					strPreinstallTeam = "Singapore"
				elseif rs("PreinstallTeam") & "" = "4" then
					strPreinstallTeam = "Brazil"
				elseif rs("PreinstallTeam") & "" = "5" then
					strPreinstallTeam = "CDC"
				elseif rs("PreinstallTeam") & "" = "6" then
					strPreinstallTeam = "Houston - Thin Client"
				elseif rs("PreinstallTeam") & "" = "7" then
					strPreinstallTeam = "Mobility"
				elseif rs("PreinstallTeam") & "" = "0" then
					strPreinstallTeam = "No Image Changes"
				else
					strPreinstallTeam = rs("PreinstallTeam") & ":" & "&nbsp;"
				end if

				if rs("ReleaseTeam") & "" = "1" then
					strReleaseTeam = "Houston"
				elseif rs("ReleaseTeam") & "" = "2" then
					strReleaseTeam = "Taiwan"
				elseif rs("ReleaseTeam") & "" = "3" then
					strReleaseTeam = "Mobility"
				else
					strReleaseTeam = "&nbsp;"
				end if

				if rs("DevCenter") = 2 then
					strDevCenterName = "Taiwan - Consumer"
				elseif rs("DevCenter") = 3 then
					strDevCenterName = "Taiwan - Commercial"
				elseif rs("DevCenter") = 4 then
					strDevCenterName = "Singapore"
				elseif rs("DevCenter") = 5 then
					strDevCenterName = "Brazil"
				elseif rs("DevCenter") = 6 then
					strDevCenterName = "Mobility"
				else
					strDevCenterName = "Houston"
				end if
				strOnlineReports = rs("OnLineReports") & ""
				if trim(rs("ComMarketingID")& "") = trim(CurrentUserID) or trim(rs("SMBMarketingID")& "") = trim(CurrentUserID) or trim(rs("ConsMarketingID")& "") = trim(CurrentUserID) or instr(strHardwarePMs & ",","," & trim(CurrentUserID) & "," ) > 0 then
					blnMarketingAdmin = true
				end if
				strPMID = rs("PMID") & ""
				if rs("SMID") & "" <> "" then
					strPMID = strPMID & "_" & rs("SMID")
				end if
				strPMID = "_" & strPMID & "_"
				strProductName = rs("DotsName") 'rs("Name") & " " & rs("Version")
				Response.Write strProductName
				DisplayedProductName = strProductName
                Response.Cookies("ProductName") = strProductName


				if trim(rs("ProductStatus")&"") = "1" then
					strStatus = "Development"
					strPORDate = rs("PDDReleased") & ""
				else
					strStatus = rs("productStatus")&""
				end if

				strOTSName = rs("DOTSName") & ""
				strProductFilename = rs("Name") + " " + rs("Version")

				strPDDPath = rs("PDDPath") & ""
				strSCMPath = rs("SCMPath") & ""
				strAccessoryPath = rs("AccessoryPath") & ""
				strSTLStatusPath = rs("STLStatusPath") & ""
				strProgramMatrixPath = rs("ProgramMatrixPath") & ""
				strMSPEKSExecutionPath = rs("MSPEKSExecutionPath") & ""
				If rs("ServiceID")&"" = CStr(CurrentUserId) Then
			  	    blnProductServiceManager = True
			    End If
                If rs("IsExcludeIncWkfComp") then
                    IsFunCompExclude ="Exclude"
                ELSE
                    IsFunCompExclude ="Include"
                End If
                If rs("AllowFollowMarketingName") then
                    FollowMktName = 1
                End If   
                Response.Cookies("BusinessSegmentId") = rs("BusinessSegmentId")
                strBusinessSegmentId = rs("BusinessSegmentId")
				rs.Close
                '----------- Product Properties --------------
      

                rs.open "spGetProgramGroupsByProduct " & clng(strID),cn
                do while not rs.eof
                    strProductGroups = strProductGroups & ", " & replace(replace(rs("Fullname")," SDM Products","")," ", "&nbsp;")
                    rs.movenext
                loop
                rs.close
                if strProductGroups <> "" then
                    strProductGroups = mid(strProductGroups,3)
                end if

                'check if all the prls are in post por status
                rs.open "usp_Image_CanImportImageDefinition " & clng(strID), cn
		        do while not rs.EOF
			        strNonPostPORPRLList = rs("PRLList")
			        rs.MoveNext
		        loop
		        rs.Close

    			If strProdType <> "2" And strDisplayedList = "SCM" And strFusionRequirements = 0 Then server.Transfer("scm/pmview.asp")
                If strProdType <> "2" And strDisplayedList = "SCM" And strFusionRequirements = 1 Then server.Transfer("SupplyChain/pmview.asp")
                If (strProdType <> "2" And strDisplayedList = "Calls") or trim(CurrentUserPartnerType)= "2" Then Server.Transfer("service/pmview.asp")
	            If (AgencyVersion = "2" And strDisplayedList = "Agency") Then Server.Transfer("agency/pmview.asp")


				If strROMVersion = "" and (strStatus = "Development" or strStatus = "Definition" )then
					rs.open "spListTargetedBIOSVersions " & clng(strID),cn, adOpenStatic
					if not (rs.EOF and rs.BOF) then
						strROMVersion = "Targeted:&nbsp;" & rs("TargetedVersions") & ""
					end if
					rs.Close
				elseif strStatus <> "Development" and strStatus <> "Definition" then
					if strROMVersion <> "" then
						strROMVersion = "Factory:&nbsp;" & strROMVersion
					else
						strROMVersion = "Factory:&nbsp;Unknown"
					end if
				end if

				if strROMWebVersion <> "" and strROMVersion <> "" then
					strROMVersion = strROMVersion & "&nbsp;&nbsp;Web:&nbsp;" & strROMWebVersion
				elseif strROMWebVersion <> "" and strROMVersion = "" then
					strROMVersion = "Web:&nbsp;" & strROMWebVersion
				end if



				rs.Open "usp_SelectEndOfProduction " & clng(strID), cn, adOpenStatic
				if not rs.EOF then
				    strEndOfProductionDate = rs(0) & ""
				End If
				rs.Close

				If strEndOfProductionDate = "" Then
				    strEndOfProductionDate = "Not Available"
				End If

				rs.Open "spListProductOSAll " & clng(strID),cn,adOpenForwardOnly
				do while not rs.EOF
					if not isnull(rs("preinstall")) then
						if rs("preinstall") then
							strOSPreinstall = strOSPreinstall & ", " & rs("ShortName")
						end if
					end if
					if not isnull(rs("Web")) then
						if rs("Web") then
							strOSWeb = strOSWeb & ", " & rs("ShortName")
						end if
					end if
					rs.MoveNext
				loop
				rs.Close
				if strOSWeb <> "" then
					strOSweb = mid(strOSWeb,2)
				end if
				if strOSPreinstall <> "" then
					strOSPreinstall = mid(strOSPreinstall,2)
				end if

				if strOSWeb <> "" and strOSPreinstall <> "" then
					strOS = "<table><tr><td><font size=1 face=verdana>Preinstall:&nbsp;&nbsp;</font></td><td><font size=1 face=verdana>" & strOSPreinstall & "</font></td></tr><tr><td><font size=1 face=verdana>Web:</font></td><td><font size=1 face=verdana>" & strOSWeb & "</font></td></tr></table>"
				elseif strOSWeb <> "" then
					strOS = "Web: " & strOSWeb
				elseif strOSPreinstall <> "" then
					strOS = "Preinstall: " & strOSPreinstall
				else
					strOS = "&nbsp;"
				end if

				rs.Open "spListSystemTeam " & clng(strID) & ",0",cn,adOpenForwardOnly
				strMarketing = ""
				strPlatformDevelopment="&nbsp;"
				strSupplyChain ="&nbsp;"
				strService = "&nbsp;"
                strODMHWPM = "&nbsp;"
				strFinance = "&nbsp;"
                strQuality = "&nbsp;"
                strBiosLead = "&nbsp;"
                strCommodityPM = "&nbsp;"
                'Harris, Valerie -  02/29/2016 - PBI 17178/ Task 17281 - Define Additional System Team Variables
                sComMarketingName = "&nbsp;"
                sConMarketingName = "&nbsp;"
                sSMBMarketingName = "&nbsp;"
                sPOManagerName = "&nbsp;"
                sConfigManagerName = "&nbsp;"
                sProcurementPMName = "&nbsp;"
                sODMSEPMName = "&nbsp;"
                sPlanningPMName = "&nbsp;"

				do while not rs.EOF
					strEmailName = "<a href=""mailto:" & rs("Email") & """>" & longname(rs("Name")) & "</a>"
					select case lcase(trim(rs("Role") & ""))
					case "program office program manager"
                        sPOManagerName = strEmailName
                    case "configuration manager"
                        sConfigManagerName = strEmailName
					case "commercial marketing","consumer marketing","smb marketing"
                        If strFusionRequirements = 1 Then 'Create list of marketing for pulsar display
                            If instr(strMarketing,strEmailName) = 0 then
							strMarketing = strMarketing & "<BR>" &  strEmailName
						    End If
                        Else 'Get each marketing value for legacy display
                            Select Case lcase(trim(rs("Role") & ""))
                                Case "commercial marketing"						
                                    sComMarketingName = strEmailName
                                Case "consumer marketing"  
                                    sConMarketingName = strEmailName
                                Case "smb marketing"          
                                    sSMBMarketingName = strEmailName
                            End Select
                        End If
					case "platform development pm"
						strPlatformDevelopment = strEmailName
					case "supply chain"
						strSupplyChain = strEmailName
					case "service"
						strService = strEmailName
                    case "odm hw pm"
                        strODMHWPM = strEmailName                        
					case "commodity pm"
						strCommodityPM = strEmailName
					case "finance"
						strFinance = strEmailName
                    case "quality"
                        strQuality = strEmailName
                    case "bios lead"
                        strBiosLead = strEmailName
                    case "procurement pm"
                        sProcurementPMName = strEmailName
                    case "odm system engineering pm"
                        sODMSEPMName = strEmailName
                    case "planning pm"
                        sPlanningPMName = strEmailName
					end select

                    if (trim(rs("Role") & "") = "ODM PIN PM") and ( (rs("ID") & "") = CurrentUserID )  then                 
                        blnOdmPreinstallPM = true
                    end if

					rs.MoveNext
				loop
				if strMarketing <> "" then
					strMarketing = mid(strMarketing,5)
				else
					strMarketing = "&nbsp;"
				end if
				rs.Close

				rs.Open "spGetReferencePlatform " & clng(strID),cn,adOpenForwardOnly
				strReferencePlatform="&nbsp;"
				if not(rs.EOF and rs.BOF) then
					strReferencePlatform = rs("Name") & ""
				end if
				rs.Close

                rs.Open "usp_ProductVersion_Release " & clng(strID) & "," & clng(strBusinessSegmentId), cn, adOpenForwardOnly
				strLeadProduct="&nbsp;"
				 do while not rs.eof
                if rs("LeadProductreleaseDesc") <> "" then
                    strLeadProduct = strLeadProduct & rs("LeadProductreleaseDesc") & "," 
                end if
                rs.movenext
                loop
				rs.Close

				dim strShortName, strShortName1, strShortNameTemp, strShortNameTemp1
				dim strLogoBadge, strLogoBadge1
				dim strPHWebFamily, strPHWebFamily1
				dim strBrandName, strBrandName1
				dim strKMAT, strKMAT1
				dim strLastScmPublish, strLastScmPublish1
				dim strServiceTag, strServiceTag1
				dim strBIOSBranding,strBIOSBranding1
                dim strMasterLabel, strMasterLabel1
                'dim strBTOServiceTagName, strBTOServiceTagName1
                dim strCTOModelNumber, strCTOModelNumber1

				strLogoBadge = ""
                strLogoBadge1 = ""
				strShortName = ""
                strShortName1 = ""
                strShortNameTemp = ""
                strShortNameTemp1 = ""
				strPHWebFamily = ""
                strPHWebFamily1 = ""
				strBrandName = ""
                strBrandName1=""
				strKMAT = ""
                strKMAT1=""
				strLastScmPublish = ""
                strLastScmPublish1=""
				strServiceTag = ""
                strServiceTag1 = ""
				strBIOSBranding = ""
                strBIOSBranding1=""
                strMasterLabel = ""
                strMasterLabel1=""
                'strBTOServiceTagName = ""
                'strBTOServiceTagName1=""
                strCTOModelNumber = ""
                strCTOModelNumber1=""

        
            dim isCMPermission
            isCMPermission = false ' able to edit LogoBadge(Legacy) and Service Tag(Legacy & Pulsar) and other fields in "Marketing Names:" in "General Information"
            if (intCMProductCount > 0 or CurrentUserSysAdmin) then
                isCMPermission = true
            end if
            

				dim Logo, Logo1
				Logo = ""
                Logo1 = ""
                '09/02/2016 -- Malichi -- Added Master Label, BTO Service Tag Name, and CTO Model Number links for Pulsar products
				if (strFusionRequirements = 0) then
                    strSeries = "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">HP Brand Name (Service Tag up)</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">BIOS Branding</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
                    strSeries1 =  strSeries & ""
                elseif (strFusionRequirements = 1 and FollowMKTName = 0) then
                    strSeries = "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge C Cover</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">HP Brand Name (Service Tag up)</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">BIOS Branding</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Model Number (Service Tag down)</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">CTO Model Number</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
                    strSeries1 = strSeries & ""
                end if
              

                rs.Open  "usp_GetBrands4Product " & clng(strID) & ",1," &FollowMKTName ,cn,adOpenForwardOnly
				do while not rs.EOF	
                    if (FollowMKTName <> 1) then             
                        
                    'Marketing Names
                        'long name
                        if (rs("LongName") <> "") Then
                            strSeries1 = strSeries1 & rs("LongName") & "<BR>"
                        else
                            if rs("StreetName") <> "" then
                                strSeries1 =  strSeries1 & rs("StreetName") & " "
                                strSeries1 = strSeries1 & rs("SeriesName")
                                if trim(rs("Suffix") & "") <> "" then
				    				strSeries1 = strSeries1 & " " & trim(rs("Suffix") & "")
				    			end if
				    			strSeries1 = strSeries1 & "<BR>"
                            end if
                        end if

                        'Short Name
                        strShortNameTemp1 = "" 'clear values
                        if (rs("ShortName") <> "") Then
                        	strShortNameTemp1= strShortNameTemp1 & rs("ShortName")	
                        else
                            strShortNameTemp1= strShortNameTemp1 & rs("StreetName2") & " "
                            if rs("ShowSeriesNumberInShortName") then
				    			strShortNameTemp1 = strShortNameTemp1 & rs("SeriesName")
                            else
				    			strShortNameTemp1 = strShortNameTemp1
                            end if
                        end if

                        if strShortNameTemp1 <> "" and isCMPermission then
                            If strFusionRequirements = 1 Then
				    		    strShortName1 = strShortName1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & strShortNameTemp1 & "',6," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">" & strShortNameTemp1 & "</a><BR>"
                            else
                                strShortName1 = strShortName1 & strShortNameTemp1 & "<BR>"
                            end if					
				    	elseif isCMPermission then
                            If strFusionRequirements = 1 Then
				    		    strShortName1 = strShortName1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',6," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">Add</a><BR>"
                            end if						
				    	else
				    		strShortName1 = strShortName1 & strShortNameTemp1 & "<BR>"
                        end if
                    


                    'logo name
                    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs - removed 1311: ---
                    if isCMPermission then
						Logo1 = rs("StreetName3") & " "
                    end if
                    if rs("ShowSeriesNumberInLogoBadge") then
                        if rs("SplitSeriesForLogoAndBrand") then
                            if isCMPermission then
                                Logo1 = Logo1 & val(rs("SeriesName"))
                            end if
                        else
                            if isCMPermission then
								Logo1 = Logo1 & rs("SeriesName")							
                            end if
                        end if

					end if
                    if (rs("LogoBadge") = "") then
                       	if isCMPermission then
							Logo1 = rs("StreetName3") & " "
						else
							strLogoBadge1 =  strLogobadge1 & rs("StreetName3") & " "
						end if						
                        if rs("ShowSeriesNumberInLogoBadge") then
                            if rs("SplitSeriesForLogoAndBrand") then
                                if isCMPermission then
                                    Logo1 = Logo1 & val(rs("SeriesName"))
                                else
								    strLogoBadge1 = strLogobadge1 & val(rs("SeriesName"))
								end if
                            else
                                if isCMPermission then
								    Logo1 = Logo1 & rs("SeriesName")
								else
								    strLogoBadge1 = strLogobadge1 & rs("SeriesName")
								end if
                            end if

						end if                     
                    end if 
                    if trim(rs("LogoBadge") & "") <> "" and Logo1 <> "" then
                        If strFusionRequirements = 0 Then
						strLogoBadge1 = strLogoBadge1 & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & rs("LogoBadge") & "',2,'" & Logo1 & "'," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">" & trim(rs("LogoBadge")) & "</a><BR>"
                        else						
                        strLogoBadge1 = strLogoBadge1 & trim(rs("LogoBadge")) & "<BR>" ' Pulsar Product
                        end if
					elseif Logo1 <> "" then ' and trim(rs("LogoBadge") & "") = ""
                        If strFusionRequirements = 0 Then
						strLogoBadge1 = strLogoBadge1 & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & Logo1 & "',2,'" & Logo1 & "'," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">" & trim(Logo1) & "</a><BR>"
					    else
                        strLogoBadge1 = strLogoBadge1 & trim(Logo1) & "<BR>" ' Pulsar Product
                        end if
					elseif isCMPermission then 'and Logo1 = ""
                        If strFusionRequirements = 0 Then
						strLogoBadge1 = strLogoBadge1 & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",0,2,'" & Logo1 & "'," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">Add</a><BR>"
                        end if						
					else ' Logo1 = "" and not isCMPermission
						strLogoBadge1 = strLogoBadge1 & rs("LogoBadge") & "<BR>"
					end if
                        
                    if rs("ServiceTag") <> "" and isCMPermission then
                        If strFusionRequirements = 1 Then
						    strServiceTag1 = strServiceTag1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("ServiceTag") & "',7," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">" & rs("ServiceTag") & "</a><BR>"
                        else
                            strServiceTag1 = strServiceTag1 & rs("ServiceTag") & "<BR>"
                        end if					
					elseif isCMPermission then
                        If strFusionRequirements = 1 Then
						    strServiceTag1 = strServiceTag1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',7," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">Add</a><BR>"
                        end if						
					else
						strServiceTag1 = strServiceTag1 & rs("ServiceTag") & "<BR>"
                    end if

                    'BIOS Branding
                    if trim(rs("BIOSBranding") & "") <> "" and isCMPermission then
                        strBIOSBranding1 = strBIOSBranding1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BIOSBranding") & "',9," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">" & trim(rs("BIOSBranding")) & "</a><BR>"		
					elseif isCMPermission then
                        strBIOSBranding1 = strBIOSBranding1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',9," & rs("ProductBrandID") & ",'" & rs("SeriesName") & "')"">Add</a><BR>"			
					else
						strBIOSBranding1 = strBIOSBranding1 & rs("BIOSBranding") & "<BR>"
                    end if

                    if rs("MasterLabel") <> "" and isCMPermission then
                        If strFusionRequirements = 1 Then
						    strMasterLabel1 = strMasterLabel1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("MasterLabel") & "',8," & rs("ProductBrandID") & ",'" & rs("SeriesID") & "')"">" & rs("MasterLabel") & "</a><BR>"
                        else
                            strMasterLabel1 = strMasterLabel1 & rs("MasterLabel") & "<BR>"
                        end if					
					elseif isCMPermission then
                        If strFusionRequirements = 1 Then
						    strMasterLabel1 = strMasterLabel1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',8," & rs("ProductBrandID") & ",'" & rs("SeriesID") & "')"">Add</a><BR>"
                        end if						
					else
						strMasterLabel1 = strMasterLabel1 & rs("MasterLabel") & "<BR>"
                    end if

                    'PER EFREN'S REQUEST - DO NOT REMOVE
                    'BTO Service Tag Name
                    'if trim(rs("BTOServiceTagName") & "") <> "" and isCMPermission then
					'	strBTOServiceTagName1 = strBTOServiceTagName1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BTOServiceTagName") & "',4," & rs("ProductBrandID") & ")"">" & rs("BTOServiceTagName") & "</a><BR>"
					'elseif isCMPermission then
					'	strBTOServiceTagName1 = strBTOServiceTagName1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',4," & rs("ProductBrandID") & ")"">Add</a><BR>"
					'else
					'	strBTOServiceTagName1 = strBTOServiceTagName1 & rs("BTOServiceTagName") & "<BR>"
					'end if

                    'CTO Model Number
                    if trim(rs("CTOModelNumber") & "") <> "" and isCMPermission then
						strCTOModelNumber1 = strCTOModelNumber1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("CTOModelNumber") & "',5," & rs("ProductBrandID") & ",'" & rs("SeriesID") & "')"">" & rs("CTOModelNumber") & "</a><BR>"
					elseif isCMPermission then
						strCTOModelNumber1 = strCTOModelNumber1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',5," & rs("ProductBrandID") & ",'" & rs("SeriesID") & "')"">Add</a><BR>"
					else
						strCTOModelNumber1 = strCTOModelNumber1 & rs("CTOModelNumber") & "<BR>"
					end if
                'PHWeb Names
                    'Brand Name
                    if (rs("BrandName") <> "") then
                        strBrandName1 = strBrandName1 & rs("BrandName") & "<BR>"
                    else
                        strBrandName1 = strBrandName1 & rs("StreetName") & " "
                        if rs("ShowSeriesNumberInBrandname") then
                            if rs("SplitSeriesForLogoAndBrand") then
    						    strBrandName1 = strBrandName1 & val(rs("SeriesName"))
	                        else
    						    strBrandName1 = strBrandName1 & rs("SeriesName")
                            end if
    				    else
						    strBrandName1 = strBrandName1
					    end if
                        strBrandName1 = strBrandName1 & "<BR>"
                    end if
                    'Family Name
                     if (rs("FamilyName") <> "") then
                        strPHWebFamily1 = strPHWebFamily1 & rs("FamilyName") & "<BR>"
                    else
                        if rs("productversion") & "" <> "" then
							if lcase(rs("productfamily")&"") = "davos"  and right(rs("productversion") & "",3) = "1.0" then
								strPHWebFamily1 = strPHWebFamily1 & left(rs("product"),len(rs("product"))-1) & "X - " &  rs("StreetName") & " "
							elseif isnumeric(mid(rs("productversion") & "",len(rs("productversion") & ""),1)) then
								strPHWebFamily1 = strPHWebFamily1 & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-1) & "X - " &  rs("StreetName") & " "
							elseif len(rs("productversion")) > 1 then
								strPHWebFamily1 = strPHWebFamily1 & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-2) & "X - " &  rs("StreetName") & " "
							else
                            	strPHWebFamily1 = strPHWebFamily1 & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  rs("productversion") & "X - " &  rs("StreetName") & " "
                                
                            end if
                        else
								strPHWebFamily1 = strPHWebFamily1 & rs("product") & " " & rs("RASSegment") & " - " &  rs("StreetName") & " "
						end if
                        strPHWebFamily1 = strPHWebFamily1 & rs("SeriesName")
                        strPHWebFamily1 = strPHWebFamily1 & "<BR>"
                    end if
                    'KMAT
                    end if 'end if for FollowMKTName <> 1
                    strKMAT1 = strKMAT1 & rs("KMAT") & " "
                    strKMAT1 = strKMAT1 & "<BR>"
                    'Last SCM Publish
                    strLastScmPublish1 = strLastScmPublish1 & rs("LastPublishDt") & " "
                    strLastScmPublish1 = strLastScmPublish1 & "<BR />"
                    
                    rs.MoveNext
				loop
				rs.Close      
                Dim iCount : iCount = 0  
				rs.Open  "spListbrands4Product " & clng(strID) & ",1",cn,adOpenForwardOnly
				do while not rs.EOF
					if trim(rs("SeriesSummary") & "") <> "" then
						SeriesArray = split(rs("SeriesSummary"),",")
						for i = 0 to ubound(SeriesArray)
							if trim(seriesArray(i)) <> "" then
								if rs("StreetName") <> "" then
									strSeries =  strSeries & rs("StreetName") & " "
                                    if strShortNameTemp1 <>"" then
                                       strShortNameTemp = strShortNameTemp1
                                    else
									    trShortNameTemp= trShortNameTemp & rs("StreetName2") & " "
                                    END IF
									if isCMPermission then
									    Logo = rs("StreetName3") & " "
									else
									    strLogoBadge =  strLogobadge & rs("StreetName3") & " "
									end if

								    strBrandName = strBrandName & rs("StreetName") & " "
								    strKMAT = strKMAT & rs("KMAT") & " "
								    strLastScmPublish = strLastScmPublish & rs("LastPublishDt") & " "
								end if

								if rs("productversion") & "" <> "" then
									if lcase(rs("productfamily")&"") = "davos"  and right(rs("productversion") & "",3) = "1.0" then
										strPHWebFamily = strPHWebFamily & left(rs("product"),len(rs("product"))-1) & "X - " &  rs("StreetName") & " "
									elseif isnumeric(mid(rs("productversion") & "",len(rs("productversion") & ""),1)) then
										strPHWebFamily = strPHWebFamily & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-1) & "X - " &  rs("StreetName") & " "
									elseif len(rs("productversion")) > 1 then
								        strPHWebFamily = strPHWebFamily & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-2) & "X - " &  rs("StreetName") & " "
							        else
                            	        strPHWebFamily = strPHWebFamily & left(rs("product"),len(rs("product")) - len (rs("productversion")) -  1) & " " & rs("RASSegment") & " " &  rs("productversion") & "X - " &  rs("StreetName") & " "
                              
        
                                    end if
                                else
										strPHWebFamily = strPHWebFamily & rs("product") & " " & rs("RASSegment") & " - " &  rs("StreetName") & " "
								end if

								strSeries = strSeries & seriesArray(i)
                                if rs("ShowSeriesNumberInShortName") then
								    strShortNameTemp = strShortNameTemp & seriesArray(i)
                                else
								    strShortNameTemp = strShortNameTemp
                                end if

								if rs("ShowSeriesNumberInLogoBadge") then
                                    if rs("SplitSeriesForLogoAndBrand") then
                                        if isCMPermission then
                                            Logo = Logo & val(seriesArray(i))
                                        else
								            strLogoBadge = strLogobadge & val(seriesArray(i))
								        end if
                                    else
                                        if isCMPermission then
								            Logo = Logo & seriesArray(i)
								        else
								            strLogoBadge = strLogobadge & seriesArray(i)
								        end if
                                    end if

								end if
								if rs("ShowSeriesNumberInBrandname") then
                                    if rs("SplitSeriesForLogoAndBrand") then
    								    strBrandName = strBrandName & val(seriesArray(i))
	                                else
    								    strBrandName = strBrandName & seriesArray(i)
                                    end if
    							else
								    strBrandName = strBrandName
								end if
								strPHWebFamily = strPHWebFamily & seriesArray(i)

								
								if trim(rs("Suffix") & "") <> "" then
									strSeries = strSeries & " " & trim(rs("Suffix") & "")
								end if

								strSeries = strSeries & "<BR>"

                                iCount = 0
                                rs4.Open  "usp_SelectShortNameMarketingNames " & trim(rs("ProductBrandID") & ""),cn,adOpenForwardOnly
                                if not(rs4.EOF and rs4.bof) then
                                     do while not rs4.EOF
						                if trim(strShortNameTemp & "") <> "" and trim(rs4("Series") & "") = trim(seriesArray(i)) and isCMPermission and strFusionRequirements = 1 then
                                            strShortName = strShortName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & strShortNameTemp & "',6," & rs("ProductBrandID") & ",'" & trim(seriesArray(i)) & "')"">" & strShortNameTemp & "</a><BR>"
                                            iCount = 1
                                        elseif trim(strShortNameTemp & "") <> "" and trim(rs4("Series") & "") = trim(seriesArray(i)) then
                                            strShortName = strShortName & strShortNameTemp & "<BR>"
                                            iCount = 1                                
                                        end if
                                        rs4.MoveNext
                                    loop 
                                end if
                                if iCount = 0 and isCMPermission and strFusionRequirements = 1 then
                                    strShortName = strShortName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & strShortNameTemp & "',6," & rs("ProductBrandID") & ",'" & trim(seriesArray(i)) & "')"">Add</a><BR>"			       
                                elseif iCount = 0 then
                                    strShortName = strShortName & strShortNameTemp & "<BR>"
                                end if
                                rs4.Close

								strBrandName = strBrandName & "<BR>"
								strPHWebFamily = strPHWebFamily & "<BR>"
								strKMAT = strKMAT & "<BR>"
								strLastScmPublish = strLastScmPublish & "<BR />"
								
                                if trim(rs("LogoBadge") & "") <> "" and Logo <> "" then
                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & rs("LogoBadge") & "',2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">" & trim(rs("LogoBadge")) & "</a><BR>"
                                    else								    
                                    strLogoBadge = strLogoBadge & trim(rs("LogoBadge")) & "<BR>"
                                    end if
								elseif Logo <> "" then
                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & Logo & "',2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">" & trim(Logo) & "</a><BR>"
                                    else
                                    strLogoBadge = strLogoBadge & trim(Logo) & "<BR>"
                                    end if
								
                                elseif isCMPermission then
                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",0,2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if    
								else
								    strLogoBadge = strLogoBadge & "<BR>"
								end if
								
								strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                if rs("ServiceTag") <> "" and isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strServiceTag = strServiceTag & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("ServiceTag") & "',7," & rs("ProductBrandID") & ",'')"">" & rs("ServiceTag") & "</a><BR>"
                                    else
                                        strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                    end if					
					            elseif isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strServiceTag = strServiceTag & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',7," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if						
					            else
						            strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                end if

                                if rs("MasterLabel") <> "" and isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strMasterLabel = strMasterLabel & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("MasterLabel") & "',8," & rs("ProductBrandID") & ",'')"">" & rs("MasterLabel") & "</a><BR>"
                                    else
                                        strMasterLabel = strMasterLabel & rs("MasterLabel") & "<BR>"
                                    end if					
					            elseif isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strMasterLabel = strMasterLabel & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',8," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if						
					            else
						            strMasterLabel = strMasterLabel & rs("MasterLabel") & "<BR>"
                                end if
                                
                                'PER EFREN'S REQUEST - DO NOT REMOVE
                                'BTO Service Tag Name
                                'if trim(rs("BTOServiceTagName") & "") <> "" and isCMPermission then
						        '    strBTOServiceTagName = strBTOServiceTagName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BTOServiceTagName") & "',4," & rs("ProductBrandID") & ")"">" & rs("BTOServiceTagName") & "</a><BR>"
					            'elseif isCMPermission then
						        '    strBTOServiceTagName = strBTOServiceTagName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',4," & rs("ProductBrandID") & ")"">Add</a><BR>"
					            'else
						        '    strBTOServiceTagName = strBTOServiceTagName & rs("BTOServiceTagName") & "<BR>"
					            'end if

                                'CTO Model Number
                                if trim(rs("CTOModelNumber") & "") <> "" and isCMPermission then
						            strCTOModelNumber = strCTOModelNumber & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("CTOModelNumber") & "',5," & rs("ProductBrandID") & ",'')"">" & rs("CTOModelNumber") & "</a><BR>"
					            elseif isCMPermission then
						            strCTOModelNumber = strCTOModelNumber & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',5," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
					            else
						            strCTOModelNumber = strCTOModelNumber & rs("CTOModelNumber") & "<BR>"
					            end if

                                iCount = 0
                                rs3.Open  "usp_SelectBIOSBrandingMarketingNames " & trim(rs("ProductBrandID") & ""),cn,adOpenForwardOnly
                                if not(rs3.EOF and rs3.bof) then
                                     do while not rs3.EOF
						                if trim(rs3("BIOSBranding") & "") <> "" and trim(rs3("Series") & "") = trim(seriesArray(i)) and isCMPermission then
                                            strBIOSBranding = strBIOSBranding & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BIOSBranding") & "',9," & rs("ProductBrandID") & ",'" & trim(seriesArray(i)) & "')"">" & rs3("BIOSBranding") & "</a><BR>"
                                            iCount = 1
                                        'Bug 21318 - Display BIOS branding even for users without permission
                                        elseif trim(rs3("BIOSBranding") & "") <> "" and trim(rs3("Series") & "") = trim(seriesArray(i)) then
                                            strBIOSBranding = strBIOSBranding & rs3("BIOSBranding") & "<BR>"
                                            iCount = 1                                
                                        end if
                                        rs3.MoveNext
                                  loop 
                                end if
                                if iCount = 0 and isCMPermission then
                                    strBIOSBranding = strBIOSBranding & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',9," & rs("ProductBrandID") & ",'" & trim(seriesArray(i)) & "')"">Add</a><BR>"			       
                                elseif iCount = 0 then
                                    strBIOSBranding = strBIOSBranding & rs("BIOSBranding") & "<BR>"
                                end if

                                rs3.Close

							end if
						next
					elseif trim(rs("SeriesSummary") & "") = "" then
								if rs("StreetName") <> "" then
									strSeries =  strSeries & rs("StreetName") & " "
									strShortNameTemp= strShortNameTemp & rs("StreetName2") & " "

									if isCMPermission then
									    Logo = rs("StreetName3") & " "
									else
									    strLogoBadge =  strLogobadge & rs("StreetName3") & " "
									end if

								    strBrandName = strBrandName & rs("StreetName") & " "
								    
								    strKMAT = strKMAT & rs("KMAT") & " "
								end if

								if rs("productversion") & "" <> "" then
									if lcase(rs("productfamily")&"") = "davos"  and right(rs("productversion") & "",3) = "1.0" then
										strPHWebFamily = strPHWebFamily & left(rs("product"),len(rs("product"))-1) & "X - " &  rs("StreetName") & " "
									elseif isnumeric(mid(rs("productversion") & "",len(rs("productversion") & ""),1)) then
										strPHWebFamily = strPHWebFamily & rs("productfamily") & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-1) & "X - " &  rs("StreetName") & " "
									else
										strPHWebFamily = strPHWebFamily & rs("productfamily") & " " & rs("RASSegment") & " " &  left(rs("productversion"), len(rs("productversion"))-2) & "X - " &  rs("StreetName") & " "
									end if
								end if

								if trim(rs("Suffix") & "") <> "" then
									strSeries = strSeries & " " & trim(rs("Suffix") & "")
								end if

								strSeries = strSeries & "<BR>"

                                if strShortNameTemp <> "" and isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strShortName = strShortName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & strShortNameTemp & "',6," & rs("ProductBrandID") & ",'')"">" & strShortNameTemp & "</a><BR>"
                                    else
                                        strShortName = strShortName & strShortNameTemp & "<BR>"
                                    end if					
					            elseif isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strShortName = strShortName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',6," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if						
					            else
						            strShortName = strShortName & strShortNameTemp & "<BR>"
                                end if

								strBrandName = strBrandName & "<BR>"
								strPHWebFamily = strPHWebFamily & "<BR>"
								strKMAT = strKMAT & "<BR>"

								if trim(rs("LogoBadge") & "") <> "" and Logo <> "" then

                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & rs("LogoBadge") & "',2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">" & trim(rs("LogoBadge")) & "</a><BR>"
                                    else								    
                                    strLogoBadge = strLogoBadge & trim(rs("LogoBadge")) & "<BR>"
                                    end if
								elseif Logo <> "" then
                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",'" & Logo & "',2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">" & trim(Logo) & "</a><BR>"
                                    else 
                                    strLogoBadge = strLogoBadge & trim(Logo) & "<BR>"
                                    end if								
                                elseif isCMPermission then
                                    If strFusionRequirements = 0 Then
								    strLogoBadge = strLogoBadge & "<a href=""javascript:AddEditMarketingName(" & rs("ID") & ",0,2,'" & Logo & "'," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
				                end if				    
								else
								    strLogoBadge = strLogoBadge & "<BR>"
								end if

                                strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                if rs("ServiceTag") <> "" and isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strServiceTag = strServiceTag & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("ServiceTag") & "',7," & rs("ProductBrandID") & ",'')"">" & rs("ServiceTag") & "</a><BR>"
                                    else
                                        strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                    end if					
					            elseif isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strServiceTag = strServiceTag & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',7," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if						
					            else
						            strServiceTag = strServiceTag & rs("ServiceTag") & "<BR>"
                                end if
					
                                if rs("MasterLabel") <> "" and isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strMasterLabel = strMasterLabel & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("MasterLabel") & "',8," & rs("ProductBrandID") & ",'')"">" & rs("MasterLabel") & "</a><BR>"
                                    else
                                        strMasterLabel = strMasterLabel & rs("MasterLabel") & "<BR>"
                                    end if					
					            elseif isCMPermission then
                                    If strFusionRequirements = 1 Then
						                strMasterLabel = strMasterLabel & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',8," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
                                    end if						
					            else
						            strMasterLabel = strMasterLabel & rs("MasterLabel") & "<BR>"
                                end if
					      
                                'PER EFREN'S REQUEST - DO NOT REMOVE
                                'BTO Service Tag Name
                                'if trim(rs("BTOServiceTagName") & "") <> "" and isCMPermission then
						        '    strBTOServiceTagName = strBTOServiceTagName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BTOServiceTagName") & "',4," & rs("ProductBrandID") & ")"">" & rs("BTOServiceTagName") & "</a><BR>"
					            'elseif isCMPermission then
						        '    strBTOServiceTagName = strBTOServiceTagName & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',4," & rs("ProductBrandID") & ")"">Add</a><BR>"
					            'else
						        '    strBTOServiceTagName = strBTOServiceTagName & rs("BTOServiceTagName") & "<BR>"
					            'end if

                                'CTO Model Number
                                if trim(rs("CTOModelNumber") & "") <> "" and isCMPermission then
						            strCTOModelNumber = strCTOModelNumber & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("CTOModelNumber") & "',5," & rs("ProductBrandID") & ",'')"">" & rs("CTOModelNumber") & "</a><BR>"
					            elseif isCMPermission then
						            strCTOModelNumber = strCTOModelNumber & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',5," & rs("ProductBrandID") & ",'')"">Add</a><BR>"
					            else
						            strCTOModelNumber = strCTOModelNumber & rs("CTOModelNumber") & "<BR>"
					            end if

								if trim(rs("BIOSBranding") & "") <> "" and (intCMProductCount > 0 or CurrentUserSysAdmin) then
                                    strBIOSBranding = strBIOSBranding & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'" & rs("BIOSBranding") & "',9," & rs("ProductBrandID") & ",'')"">" & rs("BIOSBranding") & "</a><BR>"							
                                elseif (intCMProductCount > 0 or CurrentUserSysAdmin) then
                                    strBIOSBranding = strBIOSBranding & "<a href=""javascript:ShowMarketingNameDialog(" & rs("ID") & ",'',9," & rs("ProductBrandID") & ",'')"">Add</a><BR>"						    
								else
								    strBIOSBranding = strBIOSBranding & rs("BIOSBranding") & "<BR>"
                                end if



							'end if
						'next
					end if
					rs.MoveNext
				loop
				rs.Close

                '09/02/2016 -- Malichi -- Added Master Label, BTO Service Tag Name, and CTO Model Number links for Pulsar products
				if (strFusionRequirements = 0) then
				    strSeries = strSeries &  "</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strShortName & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strLogoBadge & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strServiceTag & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strBIOSBranding & "</TD></TR></table>"
                    strSeries1 = strSeries1 &  "</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strShortName1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strLogoBadge1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strServiceTag1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strBIOSBranding1 & "</TD></TR></table>"
                elseif (strFusionRequirements = 1 and FollowMktName =0) then
                    strSeries = strSeries &  "</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strShortName & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strLogoBadge & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strServiceTag & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strBIOSBranding & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strMasterLabel & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strCTOModelNumber & "</TD></TR></table>"
                    strSeries1 = strSeries1 &  "</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strShortName1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strLogoBadge1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strServiceTag1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strBIOSBranding1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strMasterLabel1 & "</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">" & strCTOModelNumber1 & "</TD></TR></table>"
				end if

                if clng(strID) = 1192 then
                  strSeries =  "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
				  strSeries = strSeries &  "HP EliteBook Revolve 810 G2</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">HP EliteBook Revolve</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">palm rest: EliteBook | ant cover: Revolve 810</TD></TR></table>"
                    strSeries1 =  "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
				  strSeries1 = strSeries1 &  "HP EliteBook Revolve 810 G2</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">HP EliteBook Revolve</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">palm rest: EliteBook | ant cover: Revolve 810</TD></TR></table>"
                end if

                if clng(strID) = 1500 then
                  strSeries =  "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
				  strSeries = strSeries &  "HP EliteBook Revolve 810 G3</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">HP EliteBook Revolve</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">palm rest: EliteBook | ant cover: Revolve 810 G3</TD></TR></table>"
                    strSeries1 =  "<TABLE cellspacing=0 cellpadding=0><TR><TD width=210 nowrap style=""font-size:xx-small;font-weight:bold"">Long Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Short Name</TD><TD width=10>&nbsp;</td><TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo Badge</TD></TR><TR><TD  nowrap valign=top style=""font-size:xx-small"">"
				  strSeries1 = strSeries1 &  "HP EliteBook Revolve 810 G3</TD><TD nowrap width=10>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">HP EliteBook Revolve</TD><TD nowrap width=30>&nbsp;</td><TD nowrap valign=top style=""font-size:xx-small"">palm rest: EliteBook | ant cover: Revolve 810 G3</TD></TR></table>"
                end if
                
				'Pull other PMs and SMs and SEPMs
				rs.Open "spListPMsActive 0",cn,adOpenForwardOnly
				do while not rs.EOF
					strPMID = strPMID & "_" & rs("ID")
					rs.MoveNext
				loop
				rs.Close

				'Determine who has full edit rights on the product properties screen
				rs.Open "spListPMsActive 3",cn,adOpenForwardOnly
				blnEditProductProperties = false
				do while not rs.EOF
					if trim(currentuserid) =  trim(rs("ID")) then
           				blnEditProductProperties = true
                        exit do
					end if
					rs.MoveNext
				loop
				rs.Close



				'Verify Access is OK
				if trim(CurrentUserPartner) <> "1" then
                    dim boolAllowPartners
                    boolAllowPartners = false

        			if trim(strPartnerID) = trim(CurrentUserPartner) then
                        boolAllowPartners = true
                    else
        
				        rs.Open "SELECT ProductPartnerId FROM PartnerODMProductWhitelist WHERE UserPartnerId = " + CurrentUserPartner + ";",cn,adOpenForwardOnly
				        do while not rs.EOF
					        if trim(strPartnerID) =  trim(rs("ProductPartnerId")) then
           				        boolAllowPartners = true
                                exit do
					        end if
					        rs.MoveNext
				        loop
				        rs.Close
        
					end if


					if boolAllowPartners = false then
						set rs = nothing
						set cn=nothing
						Response.Redirect "NoAccess.asp?Level=0"
					end if                    

				end if


			else
				if trim(CurrentUserPartner) <> "1" then
						set rs = nothing
						set cn=nothing
						Response.Redirect "NoAccess.asp?Level=0"
				end if

				strDescription = "This is a virtual project that contains all change requests assigned to sustaining projects as well as any request with the Sustaining System Team listed as the System Team Rep."
				strManager = "Sampson, Ava" & "&nbsp;"
				strSMName = "N/A" & "&nbsp;"
				strSEPMName = "N/A" & "&nbsp;"
				strStatus = "Development"
				strOTSNAme = ""
				strBrands = "Various"
				Response.Write "Sustaining System Team"
				rs.Close
			end if
		if strProdType =2 then
			if not (strDisplayedList = "Action" or strDisplayedList = "Status"  or strDisplayedList = "Documents") then
				strDisplayedList = "Issue"
			end if
			if not (ucase(sList) = "TOOL_WORKING" or ucase(sList) = "TOOL_ROADMAP"  or ucase(sList) = "TOOL_TASKS" or ucase(sList) = "TOOL_ISSUES") then
				sList = "Tool_Working"
			end if
		end if
		if strPMID <> "" then
			strPMID= strPMId & "_"
		end if
		if blnAdministrator or CurrentUserSysAdmin or strSEPMID = CurrentUSerID or instr(trim(strPMID),"_" & trim(CurrentUSerID) & "_") > 0  then
			blnAdministrator = true
		else
			blnAdministrator = false
		end if
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if CurrentWorkgroupID = 15 or CurrentWorkgroupID = 22 or blnAdministrator then 'CurrentUSerID = 739 or CurrentUSerID = 652 or CurrentUSerID = 396 then
			blnPreinstall = true
		end if

        blnDCRProjectPM = false
        if not blnEditProductProperties then
            rs.open "spListDCRProjectPMs",cn,adOpenKeyset
            do while not rs.eof
			    if trim(CurrentUserID) = trim(rs("ID")) then
                    blnDCRProjectPM = true
                    exit do
                end if
                rs.movenext
            loop
            rs.close
		end if

		rs.Open "spListToolsPMs",cn,adOpenKeyset
		do while not rs.EOF
			if trim(CurrentUserID) = trim(rs("ID")) then
				blnToolsPM = true
				exit do
			end if
			rs.MoveNext
		loop
		rs.Close

		if instr("," & strToolAccessList & ",","," & trim(CurrentUserID) & ",")> 0 then
			blnActionOwner = true
		else
			rs.Open "spListToolsProjectOwners " & clng(request("ID")),cn,adOpenKeyset
			do while not rs.EOF
				if trim(CurrentUserID) = trim(rs("ID")) then
					blnActionOwner = true
					exit do
				end if
				rs.MoveNext
			loop
			rs.Close
		end if

    if clng(request("ID")) <> 344 and clng(request("ID")) <> 347 and clng(request("ID")) <> 1107 then
        if strFusionRequirements = 1 THEN 
	        response.write " Information (Pulsar)"

            'Harris, Valerie -  02/29/2016 - PBI 17178/ Task 17281 - Create boolean variable that sets Product's Pulsar/Legacy type
            bIsPulsarProduct = True
        else
            response.write " Information (Legacy)"

            'Harris, Valerie -  02/29/2016 - PBI 17178/ Task 17281 - Create boolean variable that sets Product's Pulsar/Legacy type
            bIsPulsarProduct = False
        end if

	end if

%>
</span><br /><br />
<table id="FavLinksTable"><tr>
<%

Dim sEditLink
Dim sCloneLink
'yong 7/20/2015, cloning is for pulsar product only; set sCloneLink = "" for legacy product
'malichi 07/19/2016, Product Backlog Item 16765: Marketing role needs permissions to Edit Product (Pulsar product)
'malichi 09/27/2016, Bug 27243: Production: Edit Product is missing from projects inside DCR Project
if (clng(request("ID")) = 344 or clng(request("ID")) = 347 or clng(request("ID")) = 1107) then 'ID Spec Changes, Core BIOS Changes, SW Spec Changes
    if trim(PVID) <> "-1" and blnCanEditProduct then
	    sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,1)"">Edit Product</a> |"
	    sCloneLink = "<a href=""javascript:ShowProperties(" & PVID & ",1,1)"">Clone Product</a> |"
    end if
elseif strFusionRequirements = 1 then 'Pulsar Product
    if trim(PVID) <> "-1" and blnCanEditProduct then
	    sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,1)"">Edit Product</a> |"
	    sCloneLink = "<a href=""javascript:ShowProperties(" & PVID & ",1,1)"">Clone Product</a> |"
    end if
elseif (blnadministrator) then  'Legacy product and blnadministrator is populated by system team roles, not permissions, users and roles for Pulsar products
    if trim(PVID) <> "-1" then
        sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,0)"">Edit Product</a> |"
        sCloneLink = ""
    end if
else
    if blnToolsPM and trim(strProdType) = "2" then
        if trim(PVID) <> "-1" then
            sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,0)"">Edit Product</a> |"
            sCloneLink = ""
        end if
    elseif blnDCRProjectPM and trim(strProdType) = "4" then
        if trim(PVID) <> "-1" then
            sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,0)"">Edit Product</a> |"
            sCloneLink = ""
        end if
    elseif blnEditProductProperties and trim(strProdType) <> "2" then
        if trim(PVID) <> "-1" then
            sEditLink = "<a href=""javascript:ShowProperties(" & PVID & ",0,0)"">Edit Product</a> |"
            sCloneLink = ""
        end if
    end if
end if
%>
<td style="white-space:nowrap;font-size:xx-small;" id="EditLink"><%=sEditLink %></td>
<td style="white-space:nowrap;font-size:xx-small;" id="CloneLink"><%=sCloneLink %></td>
<td style="white-space:nowrap;display:none;font-size:xx-small;" id="RFLink"><a href="javascript:RemoveFavorites(<%=PVID%>)">Remove From Favorites</a></td>
<td style="white-space:nowrap;display:none;font-size:xx-small;" id="AFLink"><a href="javascript:AddFavorites(<%=PVID%>)">Add To Favorites</a></td>
<%
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    if blnSupplyChain = true then   
%>
    <td style="white-space:nowrap;font-size:xx-small;" id="SIAssignments"> | <a href="javascript:SIAssignments(<%=PVID%>)">SI Assignments</a></td>
<%end if%>
<td style="white-space:nowrap;display:none;font-size:xx-small;" id="StatusLink"> | <a href="Productstatus.asp?Product=<%=DisplayedProductName%>&amp;ID=<%=PVID%>">Real-Time Status Report</a></td>
<%if strDisplayedList <> CurrentUserDefaultTab and trim(strProdType) <> "2" and clng(PVID) <> 344 and clng(PVID) <> 347 and clng(PVID) <> 1107 then%>
	<td style="white-space:nowrap;font-size:xx-small;" id="DefaultTabLink">| <a href="javascript:SetDefaultDisplay('<%=strDisplayedList%>',<%=CurrentUserID%>)">Set Default List</a></td>
<%end if%>
</tr></table>

<%if strProdType = "2" then%>
<br />
<table cellspacing="1" cellpadding="1" width="100%" border="1" bordercolor="tan" bgcolor="ivory">
	<tr>
	    <td valign="top" nowrap width="100" bgColor="cornsilk"><strong><font size="1">Manager:</font></strong></td>
		<td align="left"><font size="1"><a href="mailto:<%=strManagerEmail%>"><%=strManager%></a></font></td></tr>
	<%if trim(strDistributionList) <> "" and trim(lcase(strManagerEmail)) <> trim(lcase(strDistributionList & "&nbsp;")) then%>
	<tr>
	    <td valign="top" nowrap width="100" bgColor="cornsilk"><strong><font size="1">Distribution&nbsp;List:</font></strong></td>
		<td align="left"><font size="1"><a href="mailto:<%=strDistributionList%>"><%=strDistributionList%></a></font></td></tr>
	<%end if%>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">Roadmap:</font></strong></td>
		<td colspan="4"><font size="1"><a href="actions/Roadmap.asp?ID=<%=PVID%>" target="_blank"><%=strProductName%> Roadmap</a></font></td></tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">Status:</font></strong></td>
		<td colspan="4"><font size="1"><a href="Reports/ProjectStatus.asp?ID=<%=PVID%>&amp;Sections=5,4" target="_blank">Current Status Report</a></font></td></tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">Description:</font></strong></td>
		<td colspan="4"><font size="1"><%=strDescription%></font></td></tr>

	</table>
	<%end if%>
<br>
<%if strProdType <> 2 then %>
    <%if clng(request("ID")) = 344 or clng(request("ID")) = 347 or clng(request("ID")) = 1107 then%>
    <span style="display:none">
    <%else %>
    <span>
    <%end if%>

	<table style="display:none" Id="menubar" Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0" cellpadding="2">
	<tr bgcolor="<%=strTitleColor%>">
		<td id="CellDCR" style="Display:none" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('DCR',1)">Change&nbsp;Requests</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellDCRb" style="Display:" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Change&nbsp;Requests&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellAction" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Action',1)">Action&nbsp;Items</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellActionb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Action&nbsp;Items&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellIssue" style="display:none" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Issue',1)">Issues</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellIssueb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Issues&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellStatus" style="Display:none" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Status',1)">Status&nbsp;Notes</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellStatusb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Status&nbsp;Notes&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellOTS" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('OTS',1)">Observations</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellOTSb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Observations&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellAgency" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Agency',1)">Certifications</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellAgencyb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Certifications&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellPMR" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('PMR',1)">SMR</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellPMRb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;SMR&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellCalls" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Calls',1)">Service</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellCallsb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Service&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellGeneral" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('General',1)">General</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellGeneralb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;General&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellOpportunity" style="Display:none" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Opportunity',1)">Improvements</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellOpportunityb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Improvements&nbsp;&nbsp;&nbsp;</font></td>
		</tr>
		<tr bgcolor="<%=strTitleColor%>">
		<td id="CellRequirements" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Requirements',1)">Requirements</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellRequirementsb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Requirements&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellCountry" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Country',1)">Localization</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellCountryb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Localization&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellLocal" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Local',1)">Images</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellLocalb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Images&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellDeliverables" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Deliverables',1)">Deliverables</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellDeliverablesb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Deliverables&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellSchedule" style="Display:" width="10"><font size="1" color="black">&nbsp;<a href="javascript:SelectTab('Schedule',1)">Schedule</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellScheduleb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Schedule&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellSCM" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('SCM',1)">Supply&nbsp;Chain</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellSCMb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">Supply&nbsp;Chain&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellDocuments" width="10"><font size="1" color="white">&nbsp;<a href="javascript:SelectTab('Documents',1)">Documents</a>&nbsp;&nbsp;&nbsp;</font></td>
		<td id="CellDocumentsb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black">&nbsp;Documents&nbsp;&nbsp;&nbsp;</font></td>
    </tr>
	</table>
    </span>
<%else%>
	<table style="display:none" Id="menubar" Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0" cellpadding="2">
<%

			if request("ListFilter") = "All" or ( (not (blnActionOwner or blnToolsPM)) and trim(strProdType) = "2" and request("ListFilter") = "") then 'or request("ListFilter") = ""
				strListFilter = "All"
				strListFilterOption = "My"
				strOtherTabDefault = "All"
				strOwnerFilterList = ""
			else
				strListFilter = "My"
				strListFilterOption = "All"
				strOtherTabDefault = "My"
				strOwnerFilterList = "lstOwners=" & currentuserid & "&"
			end if
	%>


	<tr bgcolor="<%=strTitleColor%>">
	
	<%if sList="" or sList="Tool_Working" then%>
		<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Working&nbsp;List&nbsp;&nbsp;</font>
	<%else%>
		<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="PMView.asp?List=Tool_Working&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strOtherTabDefault%>">Working&nbsp;List</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if sList="Tool_Roadmap" then%>
		<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Roadmap&nbsp;&nbsp;</font>
	<%else%>
		<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="PMView.asp?List=Tool_Roadmap&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strOtherTabDefault%>">Roadmap</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if sList="Tool_Tasks" then%>
		<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Tasks&nbsp;&nbsp;</font>
	<%else%>
		<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="PMView.asp?List=Tool_Tasks&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strOtherTabDefault%>">Tasks</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if sList="Tool_Issues" then%>
		<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Issues&nbsp;&nbsp;</font>
	<%else%>
		<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="PMView.asp?List=Tool_Issues&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strOtherTabDefault%>">Issues</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	</tr></table>
	<%

		dim DisplayEdit
		dim DisplayNewRequest
		dim ToolContactList
		dim strDetailedReportOwnerFilter
		if blnActionOwner or blnToolsPM then
			DisplayEdit = ""
			DisplayNewRequest = "none"
		else
			DisplayEdit = "none"
			DisplayNewRequest = ""
		end if
		ToolContactList = trim(strSMEmail)
		if strDistributionList <> "" then
			ToolContactList= ToolContactList & ";" & strDistributionList
		end if
		if strFilterList = "My" then
			strDetailedReportOwnerFilter = "&amp;lstOwner=" & currentUserID
		else
			strDetailedReportOwnerFilter = ""
		end if
	    if strListFilterOption = "My" then
			strReorderOption = "2"
	    else
			strReorderOption = "3"
	    end if
	%>

	<%if sList="Tool_Alerts" then%>
		<font size="1"><br>&nbsp;</font><font size="1" face="verdana"><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolAction(<%=PVID%>,0,2);">Add</a> | </span><font size="1" face="verdana"><a href="javascript:window.print();">Print</a> | <a href="javascript:Export(15);">Export</a> | <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=2">Search</a> | <font size="1" face="verdana"><a href="PMView.asp?List=Tool_Alerts&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strListFilterOption%>">Display <%=strListFilterOption%> Alerts</a><br><br><font size="2"><b><%=strListFilter%> Alerts</b></font><br>
	<%elseif  sList="Tool_Tasks" then%>
		<font size="1"><br>&nbsp;</font><font size="1" face="verdana"><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolAction(<%=PVID%>,0,2);">Add</a> | </span><span style="display:<%=DisplayNewRequest%>"><a href="mailto:<%=ToolContactList%>?Subject=<%=DisplayedProductName%> - New Request">New&nbsp;Request</a> | </span><font size="1" face="verdana"><a href="javascript:window.print();">Print</a> | <a href="javascript:Export(15);">Export</a> | <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=2">Search</a> | <a target="_blank" href="Query/ActionReport.asp?txtFunction=2&amp;lstProducts=<%=strID%>&amp;lstStatus=0,1,7&amp;lstType=2&amp;<%=strOwnerFilterList%>txtTitle=<%=strProductName%> Detailed Task List">Detailed Report</a> | <a href="javascript:ReorderActions(<%=CurrentUserID%>,<%=strReorderOption%>,<%=PVID %>);">Reorder</a> | <font size="1" face="verdana"><a href="PMView.asp?List=Tool_Tasks&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strListFilterOption%>">Display <%=strListFilterOption%> Items</a><br><br><font size="2"><b><%=strListFilter%> Open Tasks</b></font><br>
	<%elseif  sList="Tool_Roadmap" then%>
		<font size="1"><br>&nbsp;</font><font size="1" face="verdana"><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolSchedule(<%=PVID%>);">Add</a> | </span><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolAction(<%=PVID%>,0,2);"><a href="javascript:ReorderItems(<%=PVID%>);">Reorder</a> | </span><a target="_blank" href="actions/Roadmap.asp?ID=<%=PVID%>">View Report</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print</a> | <a href="javascript:Export(15);">Export</a><br><br><font size="2"><b>Roadmap Items</b></font><br>
	<%elseif  sList="Tool_Working" or sList=""  then%>
		<font size="1"><br>&nbsp;</font><font size="1" face="verdana"><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolAction(<%=PVID%>,1,2);">Add</a> | </span><span style="display:<%=DisplayNewRequest%>"><a href="mailto:<%=ToolContactList%>?Subject=<%=DisplayedProductName%> - New Request">New&nbsp;Request</a> | </span><font size="1" face="verdana"><a href="javascript:window.print();">Print</a> | <a href="javascript:Export(15);">Export</a> | <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=2">Search</a> | <a target="_blank" href="Query/ActionReport.asp?txtFunction=2&amp;lstProducts=<%=strID%><%=strDetailedReportOwnerFilter%>&amp;WorkingList=1&amp;lstStatus=0,1,7&amp;lstType=2&amp;<%=strOwnerFilterList%>txtTitle=<%=strProductName%> Detailed Task List">Detailed Report</a> | <a href="javascript:ReorderActions(<%=CurrentUserID%>,1);">Reorder</a> | <font size="1" face="verdana"><a href="PMView.asp?List=Tool_Working&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strListFilterOption%>">Display <%=strListFilterOption%> Items</a><br><br><font size="2"><b><%=strListFilter%> Working List</b></font><br>
	<%elseif  sList="Tool_Issues" then%>
		<font size="1"><br>&nbsp;</font><font size="1" face="verdana"><span style="display:<%=DisplayEdit%>"><a href="javascript:AddToolAction(<%=PVID%>,0,1);">Add</a> | </span><font size="1" face="verdana"><a href="javascript:window.print();">Print</a> | <a href="javascript:Export(15);">Export</a> | <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=1">Search</a> | <a target="_blank" href="Query/ActionReport.asp?txtFunction=2&amp;lstProducts=<%=strID%>&amp;lstStatus=0,1,7&amp;lstType=1&amp;<%=strOwnerFilterList%>txtTitle=<%=strProductName%> Detailed Task List">Detailed Report</a> | <font size="1" face="verdana"><a href="PMView.asp?List=Tool_Issues&amp;ID=<%=clng(strID)%>&amp;Class=Arrow0&amp;ListFilter=<%=strListFilterOption%>">Display <%=strListFilterOption%> Items</a><br><br><font size="2"><b><%=strListFilter%> Open Issues</b></font><br>
	<%end if%>
<%end if%>

<%
	strCookie = ""
	on error resume next
	strCookie = Request.Cookies("PMStatus")
	on error goto 0

	if strCookie = "All" then
		strStatusText = "Open"
		strDelStatusText = "Targeted"
		strStatusID = 0
		strStatusDisplay = "All"
		strDelStatusDisplay = "All"
		strAgencyText = "Selected"
		strAgencyDisplay = "All"
	else
		strStatusText = "All"
		strDelStatusText = "All"
		strStatusID = 1
		strStatusDisplay = "Open"
		strDelStatusDisplay = "Targeted"
		strAgencyText = "All"
		strAgencyDisplay = "Selected"

	end if



	dim rowcount
%>

<%
'######################################
'	Schedule Tabs
'######################################

'
' If we're drawing the schedule screen, create a set of tabs based on the releases tied to the current product version
'
	If strDisplayedList = "Schedule" Then
		Dim m_ScheduleID, m_ScheduleName, bFirstWrite, m_ScheduleReleaseID, m_ProductVersionReleaseID
		Dim dw
		Set dw = New DataWrapper
		Set cm = dw.CreateCommandSP(cn, "usp_SelectSchedule")
		dw.CreateParameter cm, "@p_ProductVersionID", adInteger, adParamInput, 8, PVID
		dw.CreateParameter cm, "@p_Active_YN", adChar, adParamInput, 1, "Y"
		Set rs = dw.ExecuteCommandReturnRS(cm)

		bFirstWrite = True

		If Not rs.EOF Then

%>
<br>
<table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
	<tr>
		<td valign="top">
			<table width="100%"><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width="100%">
			<table><tr><td width="10%"><b>Release(s):</b></td><td width="90%">
<%
            Dim m_RequestScheduleID 
            if Request.QueryString("ScheduleID") then 
                m_RequestedScheduleID = Request.QueryString("ScheduleID")
                Call SaveDBCookie("ProductSchedule" & PVID, m_RequestedScheduleID, CurrentUserID)                
            else                 
                m_RequestedScheduleID = GetDBCookie("ProductSchedule" & PVID, CurrentUserID) 
                if m_RequestedScheduleID = "" then m_RequestedScheduleID = 0 end if
            end if
    
			Do Until rs.EOF
				If Not bFirstWrite Then
					Response.Write "&nbsp;|&nbsp;"
				End If

				If (m_RequestedScheduleID = 0 And m_ScheduleID = "") Or (rs("pddlocked") = 1 And m_RequestedScheduleID = "") Or (CLng(m_RequestedScheduleID) = CLng(rs("schedule_id"))) Then
					m_ScheduleID = rs("schedule_id")
					m_ScheduleName = rs("schedule_name") & ""
					m_ScheduleDescription = ""
					m_ScheduleReleaseID = rs("releaseid") & ""
                    m_ProductVersionReleaseID = rs("ProductVersionReleaseID") & ""
					Response.Write server.HTMLEncode(m_ScheduleName)                 
				Else 
					Response.Write "<a href=""javascript:scheduleLink_onClick(" & rs("schedule_id") & ")"">" & server.HTMLEncode(rs("schedule_name")) & "</a>"
				End If
                
				bFirstWrite = False
				rs.MoveNext
			Loop

			Response.Write "</td></tr></table></td></tr></table>"

			If Len(Trim(m_ScheduleDescription)) > 0 Then
				m_ScheduleDescription = " - " & m_ScheduleDescription
			End If

		End If
		rs.Close
		set cm = nothing
		set dw = nothing
	End If

'######################################
'	Localization Tabs
'######################################

'
' If we're drawing the schedule screen, create a set of tabs based on the releases tied to the current product version
'
	If strDisplayedList = "Country" Then
        Dim m_BrandID, m_BrandName
		Set dw = New DataWrapper
        if strFusionRequirements = 1 then 'pulsar product            
            Set cmd = dw.CreateCommAndSP(cn, "usp_GetProductCombinedSCMs")
            dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, PVID
            Set rs = dw.ExecuteCommAndReturnRS(cmd)           
        else
            Set cm = dw.CreateCommandSP(cn, "spListBrands4Product")
		    dw.CreateParameter cm, "@ProdID", adInteger, adParamInput, 8, PVID
		    dw.CreateParameter cm, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
		    Set rs = dw.ExecuteCommandReturnRS(cm)     
        end if       

		bFirstWrite = True

		If Not rs.EOF Then
			%>
            <br>
            <table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
	        <tr>
		        <td valign="top">
                    <table><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table></td>
		        <td width="100%">
			        <table><tr><td style="width:20%"><b>Brand:</b></td><td style="width:80%">
                        <%          
                            Do Until rs.EOF
				                'Response.Write "<td><a href=""javascript:void(0)"">" & server.HTMLEncode(rs("schedule_name")) & "</a></td>"
				                If Not bFirstWrite Then
					                Response.Write "&nbsp;|&nbsp;"
				                End If

				                If (Request("ProductBrandID") = "" And m_BrandID = "") Or (CLng(rs("ProductBrandID")) = CLng(Request("ProductBrandID"))) Then
					                m_BrandID = rs("ProductBrandID")
					                m_BrandName = rs("Name")
					                Response.Write server.HTMLEncode(m_BrandName)
				                Else
					                Response.Write "<a href=""javascript:brandLink_onClick(" & rs("ProductBrandID") & ",'" & strDisplayedList & "')"">" & server.HTMLEncode(rs("Name")) & "</a>"
				                End If

				                bFirstWrite = False
				                rs.MoveNext
			                Loop
			                Response.Write "</td></tr>"
                            'Task 31400 : Modify table to display Releases column and add releases filter
                            if strFusionRequirements = 1 then 
                                Dim rsReleases, dwReleases, cmdReleases
		                        Set rsReleases = server.CreateObject("ADODB.recordset")
                                Set dwReleases = New DataWrapper                        
                                Set cmdReleases = dwReleases.CreateCommAndSP(cn, "usp_Product_GetProductReleases")
                                dwReleases.CreateParameter cmdReleases, "@p_intProductVersionID", adInteger, adParamInput, 8, PVID
                                Set rsReleases = dwReleases.ExecuteCommAndReturnRS(cmdReleases)      
                        
                                bFirstWrite = false              
                                %>
                                <tr>
                                    <td style="width:20%"><b>Release(s):</b></td>
                                    <td style="width:80%">
                                        <%
                                            if (Request("ProductRelease") = "") then
                                                Response.Write "All"
                                            else
                                                Response.Write "<a href=""javascript:releaseLink_onClick('','" & strDisplayedList & "')"">All</a>"
                                            end if
                                            Do until rsReleases.EOF
                                                If Not bFirstWrite Then
					                                Response.Write "&nbsp;|&nbsp;"
				                                End If
                                                
                                                if (rsReleases("ReleaseName") = Request("ProductRelease")) then
                                                    Response.Write server.HTMLEncode(Request("ProductRelease"))
                                                    intReleaseID = rsReleases("ReleaseID")
                                                else
                                                    Response.Write "<a href=""javascript:releaseLink_onClick('" & rsReleases("ReleaseName") & "','" & strDisplayedList & "')"">" & server.HTMLEncode(rsReleases("ReleaseName")) & "</a>"
                                                end if
                                                bFirstWrite = False
                                                rsReleases.MoveNext
                                            Loop
                                        %>
                                    </td>
                                </tr>
                                <% 
                                rsReleases.Close
		                        set cmdReleases = nothing
		                        set dwReleases = nothing
                           end if 
                           'end task 31400         
                          
                            Response.Write "</table></td></tr></table>"

		                    End If 
		                    rs.Close
		                    set cm = nothing
		                    set dw = nothing
	                        End If

                        %>                    

<div style="display:none;" id="DCRFilters">
 <%if clng(PVID) <> 344 and clng(PVID) <> 347 and clng(PVID) <> 1107 then %>
  <br />
  <%end if%>
<table class="DisplayBar">
<tr>
  <td style="vertical-align:top"><table><tr><td class="DisplayTitle">Display:&nbsp;&nbsp;&nbsp;</td></tr></table></td>
  <td style="width:100%">
    <table style="width:100%;">
       <tr style="display:none"><td class="DisplayFilterType">Change Type:</td>
       <td style="width:100%">
      <%
		strCookie = ""
		on error resume next
		strCookie = Request.Cookies("DCRFilterType")
		on error goto 0

        If strCookie = "all" Or strCookie = "" Then
            Response.Write "All"
         Else
            Response.Write "<a href=""javascript:setDcrFilterType('all');"">All</a>"
         End If
         Response.Write "&nbsp;|&nbsp;"
         If strCookie = "dcr" Then
            Response.Write "DCR"
         Else
            Response.Write "<a href=""javascript:setDcrFilterType('dcr');"">DCR</a>"
         End If
         Response.Write "&nbsp;|&nbsp;"
         If strCookie = "bcr" Then
            Response.Write "BCR (BIOS)"
         Else
            Response.Write "<a href=""javascript:setDcrFilterType('bcr');"">BCR (BIOS)</a>"
         End If
         Response.Write "&nbsp;|&nbsp;"
         If strCookie = "scr" Then
            Response.Write "SCR"
         Else
            Response.Write "<a href=""javascript:setDcrFilterType('scr');"">SCR</a>"
         End If
      %>
      </td></tr>
      <tr><td class="DisplayFilterType">Status:</td><td width=100%>
            <%
			strCookie = ""
			on error resume next
			strCookie = Request.Cookies("DCRFilterStatus")
			on error goto 0

        If strCookie = "all" Then
            Response.Write "All"
            strStatusDisplay = "All"
         Else
            Response.Write "<a href=""javascript:setDcrFilterStatus('all');"">All</a>"
         End If
         Response.Write "&nbsp;|&nbsp;"
         If strCookie = "open"  Or strCookie = "" Then
            Response.Write "Open"
            strStatusDisplay = "Open"
         Else
            Response.Write "<a href=""javascript:setDcrFilterStatus('open');"">Open</a>"
         End If
         Response.Write "&nbsp;|&nbsp;"
         If strCookie = "closed" Then
            Response.Write "Closed"
            strStatusDisplay = "Closed"
         Else
            Response.Write "<a href=""javascript:setDcrFilterStatus('closed');"">Closed</a>"
         End If
      %>
       </td>
     </tr>
     <tr>
        <td class="DisplayFilterType"></td>
        <td width=100%>
         <!-- <a href="#" onClick="DCRWorkflowScorecard(<%=CurrentUserID%>);">DCR Workflow Scorecard</a>-->
        </td>
     </tr>
   </table>
</td>
</tr>
</table>
</div>
<span style="Display:none" ID="AddChangeLink"><font size="1"><br></font><font size="1" face="verdana"> <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=3">Search</a> | <a href="javascript:AddChange(<%=PVID%>);">Add New</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(3);">Export List to Excel</a> | <span style="Display:<%=ShowItem%>"><a href="javascript:ShowOptions(3);">Export Details to Excel</a> |</span><font size="1" face="verdana"><a href="javascript:DcrPddExport('<%= PVID%>');">PDD Export</a><br><br><font size="2"><b><%=strStatusDisplay%> Change Requests</b></font></span>
<span style="Display:none" ID="AddActionLink"><font size="1"><br></font><font size="1" face="verdana"> <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=2">Search</a> |  <a href="javascript:AddTask(<%=PVID%>);">Add New</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(2);">Export List to Excel</a> | <span style="Display:<%=ShowItem%>"><a href="javascript:ShowOptions(2);">Export Details to Excel</a> | </span><font size="1" face="verdana"><a href="javascript:DisplayStatus('<%=strStatusText%>');">Display <%=strStatusText%> Items</a><br><br><font size="2"><b><%=strStatusDisplay%> Action Items</b></font></span>
<span style="Display:none" ID="AddOpportunityLink"><font size="1"><br></font><font size="1" face="verdana"> <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=5">Search</a> |  <a href="javascript:AddOpportunity(<%=PVID%>);">Add New</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(13);">Export List to Excel</a> | <span style="Display:<%=ShowItem%>"><a href="javascript:ShowOptions(5);">Export Details to Excel</a> | </span><font size="1" face="verdana"><a href="javascript:DisplayStatus('<%=strStatusText%>');">Display <%=strStatusText%> Items</a><br><br><font size="2"><b><%=strStatusDisplay%> Improvement Opportunities</b></font></span>
<span style="Display:none" ID="AddIssueLink"><font size="1"><br></font><font size="1" face="verdana"> <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=1">Search</a> |  <a href="javascript:AddIssue(<%=PVID%>);">Add New</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(1);">Export List to Excel</a> | <span style="Display:<%=ShowItem%>"><a href="javascript:ShowOptions(1);">Export Details to Excel</a> | </span><font size="1" face="verdana"><a href="javascript:DisplayStatus('<%=strStatusText%>');">Display <%=strStatusText%> Items</a><br><br><font size="2"><b><%=strStatusDisplay%> Issues</b></font></span>
<span style="Display:none" ID="AddStatusLink"><font size="1"><br></font><font size="1" face="verdana"> <a target="_blank" href="Query/actions.asp?ID=<%=PVID%>&amp;Type=4">Search</a> |  <a href="javascript:AddStatus(<%=PVID%>);">Add New</a> | <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(7);">Export List to Excel</a><span style="Display:<%=ShowItem%>"> | <a href="javascript:ShowOptions(4);">Export Details to Excel</a></span><br><br><font size="2"><b>Status Notes</b></font></span>

<%
    'Check if user has Regions Edit Permission
    Dim HasPemission
    HasPemission = 0

    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")
    cm.CommandType = 4
	cm.CommandText = "usp_USR_ValidatePermission"
	
	Set p = cm.CreateParameter("@p_intUserId", 200, &H0001, 15)
	p.Value = CurrentUserID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@p_PName", 200, &H0001, 100)
	p.Value = "Regions.Edit"
	cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
    if not(rs.EOF and rs.BOF) then
		HasPemission = rs("HasPermission")
	end if
    rs.Close
  
    Dim Agency_EditPermission : Agency_EditPermission = 0
    Dim Regulatory_EditPermission : Regulatory_EditPermission = 0
		
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")
    cm.CommandType = 4
	cm.CommandText = "usp_USR_ValidatePermission"
	
	Set p = cm.CreateParameter("@p_intUserId", 200, &H0001, 15)
	p.Value = CurrentUserID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@p_PName", 200, &H0001, 100)
	p.Value = "Agency.Edit"
	cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
    if not(rs.EOF and rs.BOF) then
		Agency_EditPermission = rs("HasPermission")
	end if
    rs.Close

    ' get Regulatory Edit permission validate
    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")
    cm.CommandType = 4
	cm.CommandText = "usp_USR_ValidatePermission"
	
	Set p = cm.CreateParameter("@p_intUserId", 200, &H0001, 15)
	p.Value = CurrentUserID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@p_PName", 200, &H0001, 100)
	p.Value = "Regulatory_Edit"
	cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
    if not(rs.EOF and rs.BOF) then
		Regulatory_EditPermission = rs("HasPermission")
	end if
    rs.Close

    If (Agency_EditPermission + Regulatory_EditPermission) > 0 Then
        blnAgencyDataMaintainer = true
    Else
        blnAgencyDataMaintainer = false
    End If
    
%>

<span style="Display:none" ID="AddGeneralLink"><font size="1"><br></font><font size="1" face="verdana"><font size="2"><b>General Information:</b></font></span>
<span style="Display:none" ID="AddCountryLink"><font size="1"><br></font><font size="1" face="verdana">
<%if HasPemission  then%>
	<a href="javascript:ModifyCountryList(<%=PVID%>, <%=m_BrandID%>, <%=strFusionRequirements%>);">Add/Remove Countries</a> |
<%end if%>
<font size="1" face="verdana"><% If Not blnPddLocked Then %><a href="javascript:CopyLocalization(<%=PVID%>, <%=m_BrandID%>, <%=strFusionRequirements%>);"><% End If %>Import<% If Not blnPddLocked Then %></a><% End If %> | <a href="javascript:window.print();">Print</a> | <a href="javascript:Export(12);">Export to Excel</a> | <a href="Countries/rptProductCountries.asp?ID=<%=PVID%>" Target="New">Country List</a> | <a href="Countries/rptCountriesWhereUsed.asp?ID=<%=PVID%>" Target="New">Localizations List</a> | <a href="Countries/rptProductLocalization.asp?ID=<%=PVID%>" Target="New">View Matrix</a> | <a href="Countries/rptDiffSelection.asp">Comparison Reports</a><br><br>
<font size="2"><b><%= m_BrandName%> Product Localization Information:</b></font></span>

<div style="Display:none; font-size:10px; font-family:Verdana; margin:0px; padding:0px;" ID="AddScheduleLink"><br>

    <% if strFusionRequirements = 1 then %> 
        <span class="admin-scheduletab"><a href="javascript:ModifyMilestoneList_Pulsar(<%=PVID%>,<%=m_ScheduleID%>);">Add/Remove Items</a> | <% if m_ProductVersionReleaseID = 0 then %><a href="javascript:EditScheduleDescription(<%=PVID%>,<%=m_ScheduleID%>, 1);">Rename Custom Schedule</a> |<% end if %> <a href="javascript:AddNewSchedule(<%=PVID%>, 1);">Add Custom Schedule</a> | <a href="javascript:CopyMilestoneList(<%=PVID%>,<%=m_ScheduleID%>);">Copy Items</a> | <a href="javascript:ScheduleBatchEdit_Pulsar(<%=PVID%>,<%=m_ScheduleID%>, 'Projected');">Batch Edit Current</a> |  <a href="javascript:ScheduleBatchEdit_Pulsar(<%=PVID%>,<%=m_ScheduleID%>, 'Actual');">Batch Edit Actual</a> |</span>
    <%else %>
	    <span class="admin-scheduletab"><a href="javascript:ModifyMilestoneList(<%=PVID%>,<%=m_ScheduleID%>);">Add/Remove Items</a> | <a href="javascript:EditScheduleDescription(<%=PVID%>,<%=m_ScheduleID%>, 0);">Rename Schedule</a> | <a href="javascript:AddNewSchedule(<%=PVID%>, 0);">Add Schedule</a> | <a href="javascript:CopyMilestoneList(<%=PVID%>,<%=m_ScheduleID%>);">Copy Items</a> | <a href="javascript:ScheduleBatchEdit(<%=PVID%>,<%=m_ScheduleID%>, 'Projected');">Batch Edit Current</a> |  <a href="javascript:ScheduleBatchEdit(<%=PVID%>,<%=m_ScheduleID%>, 'Actual');">Batch Edit Actual</a> |</span>
    <%end if%>

<span><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(9);">Export List to Excel</a>
&nbsp;|&nbsp;<a href="reports/pdd_export_schedule.asp?PVID=<%=PVID%>">PDD Export</a></span>

<% if strFusionRequirements = 1 then %>
    <%if m_ProductVersionReleaseID = 0 Then %>
     <span class="sysadmin-scheduletab" style="margin:0px; padding:0px;">| <a href="javascript:DeleteSchedule(<%=m_ScheduleID%>);">Delete Custom Schedule</a></span>
    <%end if%>
<% else %>
    <%if clng(m_ScheduleReleaseID) > 1 Then %>
     <span class="admin-scheduletab" style="margin:0px; padding:0px;">| <a href="javascript:DeactivateSchedule(<%=m_ScheduleID%>, <%=PVID%>);">Deactivate Custom Schedule</a></span>
     <span class="sysadmin-scheduletab" style="margin:0px; padding:0px;">| <a href="javascript:DeleteSchedule(<%=m_ScheduleID%>);">Delete Custom Schedule</a></span>
    <%end if%>
<% end if %>

<br/><br /></div>
<%
if strOnlineReports <> "1" then
%>
	<span style="font-size:12px;margin:0px; padding:0px;"><strong><%=m_ScheduleName%> Schedule</strong><%=m_ScheduleDescription%></span>
<%
else
%>
	<span style="font-size:12px;margin:0px; padding:0px;"><strong><%=m_ScheduleName%></strong><%=m_ScheduleDescription%></span>
<%end if%>
<div style="Display:none" ID="AddRequirementsLink"><font size="1"></font><font size="1" face="verdana">
    <% if strFusionRequirements = 0 then 
    'moved the pulsar product to mvc so this table here only needed for legacy product
    %>   
    <table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
			    <tr>
				    <td valign="top">
					    <table width=100%><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
					    <td width="100%">
                        <table width=100%>
					    <tr>
                            <TD>Root&nbsp;Components</td>
                        </tr>
                        </table>                   
                    </td>
                </tr>
    </table>
<%end if %>
<%'if request("Display") <> "PRL" then 
    'for pulsar product, the requirement tab will combine the previous prl and new grid into one page so convert the logic request("Display" to use strFusionRequirements
   if strFusionRequirements = 0 then
    %>

        <%if blnAdministrator  or  blnMarketingAdmin then%>
	         <a href="javascript:ModifyRequirementList(<%=PVID%>);">Add/Remove Requirements</a> |
	         <a href="javascript:ImportRequirementList(<%=PVID%>);">Import Requirements</a> |

        <%end if%>
        <font size="1" face="verdana"><a href="javascript:window.print();">Print List</a> | <a href="javascript:Export(8);">Export List to Excel</a> | <a href="javascript:ExportWord(8);">Export List to Word</a><br><br>
        <font size="1"><font color="red">To view the approved requirements for this product, click here to go to its
        <%if trim(strPddPath) <> "" then%>
		    <a href="<%=strPddPath%>" target="new">Product Definition Document</a>.
        <%else%>
		        Product Definition Document.
        <%end if%>
<%  end if %>

<!--Listed below are all the master deliverables available to fulfill the product requirements.  Click the Deliverables tab above to see the subset of this list that has been selected by the system team.-->
 </font></div><!--<b>Master Deliverables List</b>-->
<span style="Display:none" ID="AddOTSLink"><font size="1"><br></font><font size="1" face="verdana"> <a href="javascript:ShowOTSAdvanced(<%=PVID%>);">Search</a> | <a href="javascript:ShowOTSDetails();">Details</a> | <a href="javascript:window.print();">Print Observation List</a> | <a href="javascript:Export(4);">Export to Excel</a><br><br><font size="2"><b>Open Observations</b></font></span>
<span style="Display:none" ID="AddAgencyLink"><font size="1"><br></font><font size="1" face="verdana"> <a href="javascript:window.print();">Print List</a> | <a href="/ipulsar/SCM/SCM_Certifications_ExportExcel.aspx?PVID=<%=PVID%>&PMStatus=<%=strAgencyDisplay%>&reportCategory=<%=reportCategory %>">Export to Excel</a> | <a href="javascript:AgencyPddExport(<%=PVID%>);">PDD Export</a><font size="2"> | </font><font size="1" face="verdana"><a href="javascript:DisplayStatus('<%=strStatusText%>');">Display <%=strAgencyText%> Items</a> <% If blnAgencyDataMaintainer = true Then %> <font size="2"> | </font> <a href="javascript:ShowReleaseStatus(0,'',0,'','<%=PVID%>','');">Batch Update For Release Status</a> <% End If %> <br><br><font size="2"><b><%=strAgencyDisplay%> Agency Certifications</b></font></span>

<span style="Display:none" ID="AddDeliverablesLink"><font size="1"><br></font>

<%

	'Get Saved Default Deliverable Settings
	dim strDefaultUserSetting
	rs.Open "spGetDefaultProductFilter " & CurrentUSerID & ",1",cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strDefaultUserSetting = ""
	else
		strDefaultUserSetting = rs("Setting") & ""
	end if
	rs.Close


dim strProductRelease
dim strDelFilter
dim strDelType
dim strDelTeam
dim strDelDefaultFilter
dim strDelDefaultType
dim strDelDefaultTeam
dim strPinSetting
dim blnDefaultFiltersDisplayed
strDelDefaultFilter = GetKeyValue(strDefaultUserSetting,"DelFilter")
strDelDefaultType = GetKeyValue(strDefaultUserSetting,"DelType")
strDelDefaultTeam = GetKeyValue(strDefaultUserSetting,"DelTeam")


	blnDefaultFiltersDisplayed = true

strProductRelease = ""
if trim(request("DelType")) = "" and trim(strDelDefaultType) <> "" then
	strDelType = strDelDefaultType
else
    regEx.Pattern = "[^A-Z]"
	strDelType = regEx.Replace(Request("DelType"), "")
    dim strSQLCommand
	if request("DelType") <> strDelDefaultType then
		blnDefaultFiltersDisplayed = false  
        strSQLCommand = "Update Employee_UserSettings SET Setting = '" & SetKeyValue(strDefaultUserSetting, "DelType", strDelDefaultType, request("DelType")) & "' where EmployeeID=" & CurrentUSerID & " and UserSettingsID=1" 
    else 
        strSQLCommand = "Insert Into Employee_UserSettings (EmployeeID, UserSettingsID, Setting) Values(" & CurrentUSerID & ", 1, '" & SetKeyValue(strDefaultUserSetting, "DelType", strDelDefaultType, request("DelType")) & "')"   
	end if
    cn.Execute strSQLCommand, adExecuteNoRecords 
end if

if trim(request("DelTeam")) = "" and trim(strDelDefaultTeam) <> "" and trim(request("DelType")) = ""then
	strDelTeam = strDelDefaultTeam
else
    regEx.Pattern = "[^0-9a-zA-Z]"
	strDelTeam = regEx.Replace(Request("DelTeam"), "")
	'strDelTeam = clng(Request("DelTeam"))

	if request("DelTeam") <> strDelDefaultTeam then
		blnDefaultFiltersDisplayed = false
	end if
end if

dim strDelFilterLink
dim strListHeaderName
strDelFilterLink = ""
if Request("DelFilter") = "Image" and strDelType <> "SW" and strDelType <> "DOC" and strDelType <> "" then
	strDelFilter = "Targeted"
elseif Request("DelFilter") = "" and strDelDefaultFilter = "Image" and strDelType <> "SW" and strDelType <> "DOC" and strDelType <> "" then
	strDelFilter = "Targeted"
elseif trim(Request("DelFilter")) = ""  and trim(strDelDefaultFilter) <> ""  then
	strDelFilter = strDelDefaultFilter
else
	strDelFilter = Request("DelFilter")
	if request("DelFilter") <> strDelDefaultFilter then
		if not(request("DelFilter") = "" and strDelDefaultFilter = "Targeted") then
			blnDefaultFiltersDisplayed = false
		end if
	end if
end if

if strDelFilter = "Targeted" then
	strDelFilterLink = "&chkTarget=on"
elseif strDelFilter = "Image" then
	strDelFilterLink = "&chkInImage=on"
else
	strDelFilterLink = ""
end if

	strPinSetting = ""
	if trim(strDelType) <> "" then
		strPinSetting = strPinSetting & "&DelType=" & strDelType
	end if
	if trim(strDelFilter) <> "" then
		strPinSetting = strPinSetting & "&DelFilter=" & strDelFilter
	end if
	if trim(strDelTeam) <> "" and strDelType = "HW" then
		strPinSetting = strPinSetting & "&DelTeam=" & strDelTeam
	end if

	if strPinSetting <> "" then
		strPinSetting = mid(strPinSetting,2)
	end if

	dim strTypeID
	dim strTypeFilter
	if strDelType = "" or strDelType = "SW" then
		strTypeID = "2"
		strTypeFilter = "&TypeID=2"
	elseif strDelType = "HW" then
		strTypeID = "1"
		strTypeFilter = "&TypeID=1"
	elseif strDelType = "FW" then
		strTypeID = "3"
		strTypeFilter = "&TypeID=3"
	elseif strDelType = "DOC" then
		strTypeID = "4"
		strTypeFilter = "&TypeID=4"
	end if

	dim strCategoryList
	strCategoryList = ""

	dim strTeamList
	dim strTeamName
	strTeamList = ""
	strTeamName = ""

    if strTypeID = 2 then
        strSQl = "spListDeliverableCoreTeams4ProductSW " & clng(PVID)
        rs.open strSQL,cn,adOpenKeyset
        do while not rs.EOF
		    if trim(strDelTeam) <> trim(rs("ID")) then
			    strTeamList = strTeamList & "&nbsp;| <a href=""pmview.asp?" & AddURLParameter(AddURLParameter(Request.QueryString,"DelTeam",rs("ID")),"DelType",strDelType) & """>" & replace(rs("name")," ","&nbsp;") & "</a>"
		    else
			    strTeamList = strTeamList & "&nbsp;| " & replace(rs("name")," ","&nbsp;")
			    strTeamName =rs("Name") &  " "
		    end if
            rs.MoveNext
        loop
        rs.Close
    else
	    rs.Open "spListDeliverableCategoryTeams " & strTypeID ,cn,adOpenStatic
	    do while not rs.EOF
		    if trim(strDelTeam) <> trim(rs("ID")) then
			    strTeamList = strTeamList & "&nbsp;| <a href=""pmview.asp?" & AddURLParameter(AddURLParameter(Request.QueryString,"DelTeam",rs("ID")),"DelType",strDelType) & """>" & replace(rs("name")," ","&nbsp;") & "</a>"
		    else
			    strTeamList = strTeamList & "&nbsp;| " & replace(rs("name")," ","&nbsp;")
			    strTeamName =rs("Name") &  " "
		    end if
		    rs.MoveNext
	    loop
	    rs.Close
    end if
	if strTeamList <> "" then
		strTeamList = mid(strTeamList,9)
	end if
	if trim(strDelTeam) = "" then
		strTeamList = strTeamList & "&nbsp;| All"
	else
		strTeamList = strTeamList & "&nbsp;| <a href=""pmview.asp?" & AddURLParameter(AddURLParameter(Request.QueryString,"DelTeam",""),"DelType",strDelType) & """>All</a>"
	end if
%>
<table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
	<tr>
		<td valign="top"><table><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width=100%>
			<table style="width:100%" border=0>
				<tr>
					<td><font size="1" face="verdana"><b>Type:</b></font></td><td nowrap width="100%"><font size="1" face="verdana">
					<%if strDelType = "SW"  or strDelType = "" then%>
						Software&nbsp;|&nbsp;
						<%
						strListHeaderName = "Software"
						strPddExportLink = "pdd_software"
						%>
					<%else%>
						<a href="pmview.asp?<%=AddURLParameter(AddURLParameter(Request.QueryString,"DelType","SW"),"DelTeam","")%>">Software</a>&nbsp;|&nbsp;
					<%end if%>
					<%if strDelType = "HW" then%>
						Hardware&nbsp;|&nbsp;
						<%
						strListHeaderName = "Hardware"
						strPddExportLink = "pdd_hardware"
						%>
					<%else%>
						<a href="pmview.asp?<%=AddURLParameter(AddURLParameter(Request.QueryString,"DelType","HW"),"DelTeam","")%>">Hardware</a>&nbsp;|&nbsp;
					<%end if%>
					<%if strDelType = "FW" then%>
						Firmware&nbsp;|&nbsp;
						<%strListHeaderName = "Firmware"%>
					<%else%>
						<a href="pmview.asp?<%=AddURLParameter(AddURLParameter(Request.QueryString,"DelType","FW"),"DelTeam","")%>">Firmware</a>&nbsp;|&nbsp;
					<%end if%>
					<%if strDelType = "DOC" then%>
						Documentation
						<%strListHeaderName = "Documentation"%>
					<%else%>
						<a href="pmview.asp?<%=AddURLParameter(AddURLParameter(Request.QueryString,"DelType","DOC"),"DelTeam","")%>">Documentation</a>
					<%end if%>
					</font></td>
				</tr>
				<tr>
					<%if (strDelType = "HW" or strDelType = "SW" or strDelType = "") and strTeamList <> "" then%>
						<td valign=top><font size="1" face="verdana"><b>Team:&nbsp;&nbsp;</b></font></td>
						<td><font size="1" face="verdana"><%=strTeamList%></font></td>
					<%end if%>
				</tr>
				<tr>
					<%if strDelType = "HW" and strCategoryList <> "" then%>
						<td valign="top"><font size="1" face="verdana"><b>Category:</b>&nbsp;&nbsp;</font></td>
						<td><font size="1" face="verdana"><%=strCategoryList%></font></td>
					<%end if%>
				</tr>
				<tr>
					<td><font size="1" face="verdana"><b>Filter:</b></font></td><td><font size="1" face="verdana">
						<%if strDelFilter = "" or strDelFilter = "Targeted"  then%>
							Targeted&nbsp;|&nbsp;
							<%strListHeaderName = "Targeted " & strTeamName & strListHeaderName  & " Deliverables"%>
						<%else%>
							<a href="pmview.asp?<%=AddURLParameter(Request.QueryString,"DelFilter","Targeted")%>">Targeted</a>&nbsp;|&nbsp;
						<%end if%>
						<%if strDelFilter = "Image" then%>
							In&nbsp;Image&nbsp;|&nbsp;
							<%strListHeaderName = strListHeaderName & " Deliverables In Image"%>
						<%elseif strDelType = "SW" or strDelType = "" or strDelType = "DOC" then%>
							<a href="pmview.asp?<%=AddURLParameter(Request.QueryString,"DelFilter","Image")%>">In&nbsp;Image</a>&nbsp;|&nbsp;
						<%end if%>
						<%if strDelFilter = "PINImage" then%>
							In&nbsp;PIN&nbsp;Image&nbsp;|&nbsp;
							<%strListHeaderName = strListHeaderName & " Deliverables In PIN Test Image"%>
						<%elseif (CurrentUserWorkgroupID = 15 or CurrentUserWorkgroupID=22) and (strDelType = "SW" or strDelType = "" or strDelType = "DOC") then%>
							<a href="pmview.asp?<%=AddURLParameter(Request.QueryString,"DelFilter","PINImage")%>">In&nbsp;PIN&nbsp;Image</a>&nbsp;|&nbsp;
						<%end if%>
						<%if strDelFilter = "Roots" then%>
							Roots&nbsp;|&nbsp;
							<%strListHeaderName = "Supported Root " & strTeamName & strListHeaderName & " Deliverables"%>
						<%else%>
							<a href="pmview.asp?<%=AddURLParameter(Request.QueryString,"DelFilter","Roots")%>">Roots</a>&nbsp;|&nbsp;
						<%end if%>
						<%if strDelFilter = "All" then%>
							All
							<%strListHeaderName = "All Supported " & strTeamName & strListHeaderName & " Deliverables"%>
						<%else%>
							<a href="pmview.asp?<%=AddURLParameter(Request.QueryString,"DelFilter","All")%>">All</a>
						<%end if%>

					</font></td>
				</tr>
                <%if strFusionRequirements = 1 then 
		        Set rsReleases = server.CreateObject("ADODB.recordset")
                Set dwReleases = New DataWrapper                        
                Set cmdReleases = dwReleases.CreateCommAndSP(cn, "usp_Product_GetProductReleases")
                dwReleases.CreateParameter cmdReleases, "@p_intProductVersionID", adInteger, adParamInput, 8, PVID
                Set rsReleases = dwReleases.ExecuteCommAndReturnRS(cmdReleases)      
                        
                bFirstWrite = false%>
                    <tr>
                    <td nowrap><b>Release(s):</b></td>
                    <td>
                        <%
                            if (Request("ProductRelease") = "") then
                                Response.Write " All"
                            else
                                Response.Write " <a href=""pmview.asp?" & AddURLParameter(Request.QueryString,"ProductRelease","""") & """>All</a>"
                            end if
                            Do until rsReleases.EOF
								If Not bFirstWrite Then
									Response.Write "&nbsp;|&nbsp;"
								End If
			
                                if (rsReleases("ReleaseName") = Request("ProductRelease")) then
                                    Response.Write server.HTMLEncode(Request("ProductRelease"))
                                    intReleaseID = rsReleases("ReleaseID")
                                else
                                    Response.Write "<a href=""pmview.asp?" & AddURLParameter(Request.QueryString,"ProductRelease",server.HTMLEncode(rsReleases("ReleaseName"))) & """>" & server.HTMLEncode(rsReleases("ReleaseName")) & "</a>"
                                end if 
                                intReleaseCount = rsReleases("ReleaseCount")     
                                bFirstWrite = False
                                NumReleases = NumReleases +1
                                rsReleases.MoveNext
                            Loop
                        %>
                    </td>
                </tr>
                <% 
                rsReleases.Close
		        set cmdReleases = nothing
		        set dwReleases = nothing
            end if%>
			</table>
		</td><td style="display:none" valign="top"><a href="javascript:SetDefaultDisplay('<%=Request.querystring%>',<%=CurrentUserID%>);">Set default display</a></td>
			<td nowrap width="100%" align="right" valign="top">

			<%
				strFilterSave = strPinSetting 
				if blnDefaultFiltersDisplayed and strDefaultUserSetting = "" and strFilterSave = "" then
			%>
				<a ID="DelPIN0" style="display:none" href="javascript:TogglePin(1,0,'<%=strFilterSave%>',<%=CurrentUserID%>);"><img SRC="images/PIN_out.gif" border="0" WIDTH="21" HEIGHT="20"></a>
				<a ID="DelPIN1" style="display:none" href="javascript:TogglePin(1,1,'',<%=CurrentUserID%>);"><img SRC="images/PIN_in.gif" border="0" WIDTH="22" HEIGHT="20"></a>
			<%
				elseif blnDefaultFiltersDisplayed then
			%>

				<a ID="DelPIN0" style="display:none" href="javascript:TogglePin(1,0,'<%=strFilterSave%>',<%=CurrentUserID%>);"><img SRC="images/PIN_out.gif" border="0" WIDTH="21" HEIGHT="20"></a>
				<a ID="DelPIN1" style="display:" href="javascript:TogglePin(1,1,'',<%=CurrentUserID%>);"><img SRC="images/PIN_in.gif" border="0" WIDTH="22" HEIGHT="20"></a>
			<%
				else
			%>

				<a ID="DelPIN0" href="javascript:TogglePin(1,0,'<%=strFilterSave%>',<%=CurrentUserID%>);"><img SRC="images/PIN_out.gif" border="0" WIDTH="21" HEIGHT="20"></a>
				<a ID="DelPIN1" style="display:none" href="javascript:TogglePin(1,1,'',<%=CurrentUserID%>);"><img SRC="images/PIN_in.gif" border="0" WIDTH="22" HEIGHT="20"></a>
			<%
				end if
			%>
			</td>
		</tr>
		</table>
		</td>
	</tr>
</table><br>
<%'end if%>
<% if trim(strDelTeam) = "2" then%>
	<font size="1" face="verdana"> <a target="_blank" href="query/deliverables.asp?ID=<%=PVID%>&amp;CommodityMatrix=1">
<%else%>
	<font size="1" face="verdana"> <a target="_blank" href="query/deliverables.asp?ID=<%=PVID%>">
<%end if%>
Search</a> |
<a href="javascript:window.print();">Print</a> |

<%if strDelType = "HW"  then%>
	<a href="javascript:Export(5);">Export</a> |
<%else%>
	<a target="_blank" href="query/DelReport.asp?txtFunction=5&lstProducts=<%=PVID%>&Type=<%=trim(strTypeID)%>&cboFormat=1<%=strDelFilterLink%>">Export</a> |
<%end if%>

<%if strDelType = "HW"  then%>
	<%if (blnHardwarePM or CurrentUserSysAdmin or blnPilotEngineer or blnAccessoryPM) and strDelFilter = "Roots" then%> 
        <%if strDelFilter = "All" then%>
        <a href="javascript:MultiUpdateTestStatusLink(<%=PVID%>,<%=intReleaseID%>,<%=strFusionRequirements%>,0);">Update Qual Status</a> | 
        <%else%>
        <a href="javascript:MultiUpdateTestStatusLink(<%=PVID%>,<%=intReleaseID%>,<%=strFusionRequirements%>,1);">Update Qual Status</a> |
        <%end if%>
    <% end if %>
	<%if (blnCommodityPM or blnTestLead or blnSETestLead) and strDelFilter <> "Roots" then%>
		<a href="javascript:MultiUpdateTestLeadStatusLink(<%=PVID%>);">Test Lead Status</a> |
	<%end if%>
	<% if trim(strDelTeam) <> "4" then%>
		<a target="_blank" href="Deliverable/HardwareMatrix.asp?lstProducts=<%=PVID%>&lstTeamID=<%=trim(strDelTeam)%>">Qual Matrix</a> |
        <a target="_blank" href="Deliverable/HardwareMatrix.asp?ReportFormat=2&lstProducts=<%=PVID%>&lstTeamID=<%=trim(strDelTeam)%>">Subassembly Matrix</a> |
		<a target="_blank" href="Deliverable/HardwareMatrix.asp?ReportFormat=5&lstProducts=<%=PVID%>&lstTeamID=<%=trim(strDelTeam)%>">Service Matrix</a> |
		<a target="_blank" href="Deliverable/HardwareMatrix.asp?ReportFormat=6&lstProducts=<%=PVID%>&lstTeamID=<%=trim(strDelTeam)%>">Samples Matrix</a> |
	<% else%>
		<a target="_blank" href="Deliverable/HardwareMatrix.asp?ReportFormat=4&lstProducts=<%=PVID%>&lstTeamID=4">Deliverable Matrix</a> |
	<%end if%>
<%else%>
	<a target="_blank" href="Image/DeliverableMatrix.asp?ProdID=<%=PVID%>&amp;PINTest=0">Deliverable Matrix</a> |
<%end if%>

	<a target="_blank" href="Image/DeliverableMatrix.asp?ProdID=<%=PVID%>&amp;PINTest=1">PIN Test Matrix</a> |

<a target="_blank" href="Image/DeliverableChanges.asp?ProductID=<%=PVID%><%=strTypeFilter%>">View History</a>
<%if strDelType <> "HW" then%>
	 | <a href="javascript: ShowReadinessReportOptions(<%=PVID%>,1,0);">Readiness&nbsp;Report</a>
<%elseif trim(strDelTeam) = "" then%>
	 | <a href="javascript: ShowReadinessReportOptions(<%=PVID%>,3,0);">Readiness&nbsp;Report</a>
<%else%>
	 | <a href="javascript: ShowReadinessReportOptions(<%=PVID%>,3,<%=clng(strDelTeam)%>);">Readiness&nbsp;Report</a>
<%end if%>

<%if strDelType = "" or strDelType = "SW" or strDelType = "HW" then%>
	 | <a href="deliverable/<%=strPddExportLink%>.asp?PVID=<%=PVID%>">PDD Export</a>
<%end if%>

<%if (blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM or (strHardwareAccessGroup =1 or strHardwareAccessGroup =2 or strHardwareAccessGroup =4)) then %>
     | <div id="divFunCompExclude" style="display: inline">
			<a  href="javascript: SaveSettingForFunExclude('<%=IsFunCompExclude%>');"> <%=IsFunCompExclude%> Functional Test Components </a>
	   </div>
<%else%>
     | <p> <%=IsFunCompExclude%> Functional Test Components </p>
<%end if%>
<%if strFusionRequirements = 1 and NumReleases>1 then
        if blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM or (strHardwareAccessGroup =1 or strHardwareAccessGroup =2 or strHardwareAccessGroup =4) or  blnHardwarePM or CurrentUserSysAdmin or blnPilotEngineer or blnAccessoryPM  or blnCommodityPM or blnSETestLead or blnODMTestLead or blnWWANTestLead or blnDEVTestLead then  'need to check user right %>
    	| <a href="javascript: ShowBatchUpdateComponentToProductRelease('/Pulsar/Product/BatchUpdateComponentToProductRelease?ProductVersionID=<%=PVID%>&ReleaseID=0&TargetReleaseID=0');">Batch Update Component Supported Releases with Targeting</a>

<%      end if    
  end if%>
<%if strFusionRequirements = 1 and (blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM or (strHardwareAccessGroup =1 or strHardwareAccessGroup =2 or strHardwareAccessGroup =4) or  blnHardwarePM or CurrentUserSysAdmin or blnPilotEngineer or blnAccessoryPM  or blnCommodityPM or blnSETestLead or blnODMTestLead or blnWWANTestLead or blnDEVTestLead) then %>
      
    	| <a href="javascript: ImportTargetingSettings('/Pulsar/Product/ImportTargetingSettings?ProductVersionID=<%=PVID%>&TargetReleaseID=0');">Import Targeting Settings</a>  
        | <a href="javascript: ImportOSImageSettings('/Pulsar/Product/ImportOSImageSettings?ProductVersionID=<%=PVID%>&TargetReleaseID=0');">Import Supported Image Settings</a>

<% end if%>
<br><br><font size="2"><b><%=strListHeaderName%></b></font></span>&nbsp&nbsp&nbsp<button type="button" id="btnRefresh" onClick="ReloadWindow();" style="background-color: #cccccc; color:#000000; font-size: 9px; font-weight:bolder;" class="hide">Refresh the Grid</button>
<!-- End of DeliverableLink-->

<span style="Display:none" ID="AddLocalLink"><font size="1"><br></font><font size="1" face="verdana">


<table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
	<tr>
		<td valign="top">
			<table><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width="100%">
			<table><tr><td><b>Image&nbsp;Status:</b></td><td width="100%">
				<%if request("ImageActiveType") = "Active"  or request("ImageActiveType") = "" then%>
					&nbsp;&nbsp;Active&nbsp;|&nbsp;
					<%strListHeaderName = "Active"%>
				<%else%>
					&nbsp;&nbsp;<a href="pmview.asp?ID=<%=PVID%>&amp;ImageActiveType=Active&amp;ImageType=<%=trim(request("ImageType"))%>&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=<%=strImageTool%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>">Active</a>&nbsp;|&nbsp;
				<%end if%>
				<%if request("ImageActiveType") = "NotReleased" then%>
					Not&nbsp;Released&nbsp;|&nbsp;
					<%strListHeaderName = "NotReleased"%>
				<%else%>
					<a href="pmview.asp?ID=<%=PVID%>&amp;ImageActiveType=NotReleased&amp;ImageType=<%=trim(request("ImageType"))%>&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=<%=strImageTool%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>">Not&nbsp;Released</a>&nbsp;|&nbsp;
				<%end if%>
				<%if request("ImageActiveType") = "InFactory" then%>
					In&nbsp;Factory&nbsp;|&nbsp;
					<%strListHeaderName = "InFactory"%>
				<%else%>
					<a href="pmview.asp?ID=<%=PVID%>&amp;ImageActiveType=InFactory&amp;ImageType=<%=trim(request("ImageType"))%>&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=<%=strImageTool%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>">In&nbsp;Factory</a>&nbsp;|&nbsp;
				<%end if%>
				<%if request("ImageActiveType") = "Inactive" then%>
					Inactive&nbsp;|&nbsp;
					<%strListHeaderName = "Inactive"%>
				<%else%>
					<a href="pmview.asp?ID=<%=PVID%>&amp;ImageActiveType=Inactive&amp;ImageType=<%=trim(request("ImageType"))%>&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=<%=strImageTool%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>">Inactive</a>&nbsp;|&nbsp;
				<%end if%>
				<%if request("ImageActiveType") = "All" then%>
					All&nbsp;&nbsp;
					<%strListHeaderName = "All"%>
				<%else%>
					<a href="pmview.asp?ID=<%=PVID%>&amp;ImageActiveType=All&amp;ImageType=<%=trim(request("ImageType"))%>&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=<%=strImageTool%>&ProductReleaseID=<%=request("ProductReleaseID")%>&ProductOSReleaseID=<%=request("ProductOSReleaseID")%>">All</a>&nbsp;&nbsp;
				<%end if%>
				</td></tr>           
                    <tr>
                        <td><b>Image&nbsp;Tool:</b></td>
                        <td>
                        <%
                            if strImageTool = "IRS" then
                                if (strFusionRequirements = 0) then
                                    response.write "&nbsp;&nbsp;<a href=""pmview.asp?ID=" & PVID & "&amp;ImageActiveType=" & request("ImageActiveType") & "&amp;ImageType=" & replace(request("ImageType")," ","") & "&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=Excalibur"">Excalibur</a>&nbsp;&nbsp;|"
                                end if
                                response.write "&nbsp;&nbsp;IRS&nbsp;&nbsp;"
                           else
                                if (strFusionRequirements = 0) then
                                    response.write "&nbsp;&nbsp;Excalibur&nbsp;&nbsp;|"
                                end if
                                response.write "&nbsp;&nbsp;<a href=""pmview.asp?ID=" & PVID & "&amp;ImageActiveType=" & request("ImageActiveType") & "&amp;ImageType=" & replace(request("ImageType")," ","") & "&amp;Class=Arrow2&amp;List=Local&amp;ImageTool=IRS"">IRS</a>&nbsp;&nbsp;"
                            end if

                        %>
                        </td>
                    </tr>
                    <% 
					if strFusionRequirements = 1 then 
						createImageReleaseFiltrer("Local")
						Call createImageOSReleaseFiltrer("Local", 1)
					else
						Call createImageOSReleaseFiltrer("Local", 0)
					end if
					
					%>
					
			</table>
		</td></tr>
		</table>
		</td>
	</tr>
</table><br>


<%if strImageTool = "IRS" then%>
    <%if blnAdministrator or CurrentWorkgroupID = 15 or  blnMArketingAdmin  or ProductImageEdit = "1" then
        if (strFusionRequirements = 0) then %>
	        <a href="javascript:AddImageFusion();">Add New</a> | <a href="javascript:ImportImageFusion();" title="Import">Import</a> |
        <%else %>
            <a href="javascript:AddImagePulsar();" title="Import">Add New</a> | <a title="Import OS Definition" href="javascript:ImportImagePulsar('<%=strNonPostPORPRLList%>');">Import OS Definition</a> |
       <%end if %>
    <%end if%>
     <font size="1" face="verdana"><a href="javascript:window.print();">Print&nbsp;List</a>&nbsp;| <a href="javascript:Export(6);">Excel&nbsp;Export</a>&nbsp;
    <%If CurrentWorkgroupID = 15 or CurrentWorkgroupID = 22 Then %>
       | <a href="javascript:EditDriveDefinitions();">Edit Drives</a>&nbsp;
    <%End If %>
    <% if (strFusionRequirements = 0) then %>
       | <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=PVID%>">Rollout&nbsp;Plan</a>
       | <a href="javascript: SyncIRSImages(<%=PVID%>)">Sync&nbsp;Images&nbsp;With&nbsp;IRS</a>&nbsp;
       | <a target="_blank" href="Image/IRSComponentList.asp?ID=<%=PVID%>">IRS&nbsp;Components</a>      
       | <a target="_blank" href="Image/Fusion/Localization.asp?ProdID=<%=PVID%>&amp;PINTest=0">Image&nbsp;Matrix</a>&nbsp;
       | <a href="javascript: ExtendEOL(<%=PVID%>, false)">Extend the EOL for Images</a>&nbsp;
     <%else %>
        | <a target="_blank" href="Image/fusion/Buildplan_Pulsar.asp?ID=<%=PVID%>&ProductReleaseID=<%=request("ProductReleaseID")%>">Rollout&nbsp;Plan</a>
        | <a target="_blank" href="../Pulsar/Report/RolloutPlanReport?productVersionId=<%=PVID%>&irsImg=1&releaseLinkId=">Pulsar&nbsp;Rollout&nbsp;Plan</a>
        | <a href="javascript: SyncIRSImages_Pulsar(<%=PVID%>)">Sync&nbsp;Images&nbsp;With&nbsp;IRS</a>&nbsp;
        | <a target="_blank" href="Image/Fusion/IRSComponentList_Pulsar.asp?ID=<%=PVID%>">IRS&nbsp;Components</a> 
        | <a target="_blank" href="Image/Fusion/Localization_Pulsar.asp?ProdID=<%=PVID%>&amp;PINTest=0&ProductReleaseID=<%=request("ProductReleaseID")%>">Image&nbsp;Matrix</a>&nbsp;
        | <a href="javascript: OpenKeyboardMatrix(<%=PVID%>)">Keyboard&nbsp;Layout&nbsp;Matrix</a>&nbsp;
        | <a href="javascript: OpenChangeHistory(<%=PVID%>)">Change&nbsp;History</a>&nbsp;
        | <a href="javascript: ExtendEOL(<%=PVID%>, true)">Extend the EOL for Images</a>&nbsp;
     <%end if%>
       </font>
    <br><br><font size="2"><b>
        &nbsp;IRS Image Definitions
    </b></font></span>
<%else%>
    <%if blnAdministrator or CurrentWorkgroupID = 15 or  blnMArketingAdmin or ProductImageEdit = "1" then%>
	    <a href="javascript:AddImage();">Add New</a> | <a href="javascript:ImportImage();">Import</a> |
    <%end if%>

    <% if CurrentWorkgroupID = 15 or CurrentWorkgroupID = 22 or CurrentUserSysAdmin or strSEPMID = CurrentUSerID or instr(trim(strPMID),"_" & trim(CurrentUSerID) & "_") > 0 or ProductImageEdit = "1"  then %>
     <font size="1" face="verdana">
         <a href="javascript:window.print();">Print&nbsp;List</a>&nbsp;| 
         <a href="javascript:Export(6);">Excel&nbsp;Export</a>&nbsp;| 
         <a target="_blank" href="Image/GenerateText.asp?ProdID=<%=PVID%>">Text&nbsp;Files</a>&nbsp;| 
         <a target="_blank" href="Image/DRDVDTextFile.asp?ProdID=<%=PVID%>">DRDVD&nbsp;Files</a>&nbsp;| 
         <a href="javascript: ShowImageCompare(<%=PVID%>,1)">Validate&nbsp;PIN&nbsp;Images</a>&nbsp;|
         <a target="_blank" href="Image/Localization.asp?ProdID=<%=PVID%>&amp;PINTest=0">Image&nbsp;Matrix</a>&nbsp;| 
         <a target="_blank" href="Image/Localization.asp?ProdID=<%=PVID%>&amp;PINTest=1">PIN&nbsp;Matrix</a>&nbsp;| 
         <a target="_blank" href="Image/rptTestMatrix.asp?ID=<%=PVID%>">SKU&nbsp;Matrix</a>&nbsp;| 
         <a target="_blank" href="Image/Buildplan.asp?ID=<%=PVID%>">Rollout&nbsp;Plan</a>
    <% else %>
     <font size="1" face="verdana">
         <a href="javascript:window.print();">Print List</a> | 
         <a href="javascript:Export(6);">Excel Export</a> | 
         <a target="_blank" href="Image/GenerateText.asp?ProdID=<%=PVID%>">Text Files</a> | 
         <a target="_blank" href="Image/Localization.asp?ProdID=<%=PVID%>&amp;PINTest=0">Image Matrix</a> | 
         <a target="_blank" href="Image/rptTestMatrix.asp?ID=<%=PVID%>">SKU Matrix</a> | 
         <a target="_blank" href="Image/Buildplan.asp?ID=<%=PVID%>">Rollout Plan</a>
    <%end if%>

      <%if strPORDate <> "" and false then%>
	    | <a target="_blank" href="Image/Changes.asp?ID=<%=PVID%>">View History</a>
      <%end if%>
      <%if blnAdministrator or ProductImageEdit = "1" then%>
       | <a href="javascript:RevImage();">Rev Images</a>
     <%end if%>
       | <a href="javascript:OemReadyVerify();">MDA Verification</a>
    <%If CurrentWorkgroupID = 15 or CurrentWorkgroupID = 22 Then %>
       | <a href="javascript:EditDriveDefinitions();">Edit Drives</a>
    <%End If %>
        | <a target="_blank" href="Image/IRSComponentList.asp?ID=<%=PVID%>">IRS&nbsp;Components</a>

    <br><br><font size="2"><b>
        &nbsp;Excalibur Image Definitions
    </b></font></span>
<%end if%>


<span ID="Wait">

<%
Response.Write  "<BR><font face=verdana size=2>Loading Information.  Please wait...</font>"
Response.Write "</span>"
'######################################
'	OTS Tabs
'######################################
if strDisplayedList <> "OTS" then
	Response.Write "<Table style=""Display:none"" ID=TableOTS><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
	if strOTSName = "" then
		Response.Write "<BR><Table style=""Display:none"" ID=TableOTS><TR><TD><b><font color=red size=2>No OTS Link Defined</font><b></td></tr></table>"
	else
		on error resume next
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetOTSPlatformDefectsSorted"

		Set p = cm.CreateParameter("@Product", 200, &H0001,20)
		p.Value = left(strOTSNAme,20)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Status", 3, &H0001)
		p.Value = 0
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute
		Set cm=nothing


		dim strVersion
		dim OTSAvailable
		if err.number = 0  then
			OTSAvailable = true
			else
			OTSAvailable = false
		end if

		If not OTSAvailable then

	%>
 <br><table style="Display:none" ID="TableOTS"><tr><td><b><font color="red" size="2">The OTS system is not available at this time.</font><b></td></tr></table>
  <%
  elseif not(rs.EOF and rs.bof) then
  %>
<br><table style="Display:none" ID="TableOTS" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">

  <thead>
    <tr>
	<td nowrap width="20" bgColor="cornsilk" vAlign="middle"><input type="checkbox" id="chkAllOTS" name="chkAllOTS" Language="javascript" onClick="onclick_ResetOTS();" style="WIDTH:16;HEIGHT:16"></td>
	<td nowrap width="50" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 1,0,2);">Number</a> </font> </td>
    <td nowrap width="30" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 2,0,2);">PR</a></font></td>
    <td nowrap width="150" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 3,0,2);">Status</a></font></td>
    <td nowrap width="90" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 4,0,2);">Owner</a></font></td>
    <td nowrap width="90" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 5,0,2);">PM</a></font></td>
    <td nowrap width="90" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 6,0,2);">Milestone</a></font></td>
    <td nowrap width="90" bgColor="cornsilk" vAlign="middle"><font size="1"><a href="javascript: SortTable( 'TableOTS', 7,0,2);">Summary</a></font></td>
    </tr>
  </thead>
  <%
	rowcount = 0
	do while not rs.EOF
		rowcount = rowcount + 1	  %>
  <tr id="otschangerows" LANGUAGE="javascript" onMouseOver="return changerows_onmouseover()" onMouseOut="return changerows_onmouseout()" onClick="return OTSrows_onclick('<%=rs("ObservationID")%>')" oncontextmenu="javascript:OTSrows_onclick('<%=rs("ObservationID")%>');return false;">
	<td nowrap width="20" bgColor="cornsilk" vAlign="middle"><input type="checkbox" style="WIDTH:16;HEIGHT:16" id="chkOTSID" name="chkOTSID" value="<%=rs("ObservationID")%>"></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ObservationID") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Priority") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("State") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("owner") & "&nbsp;"%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("programmanager") & "&nbsp;"%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("gatingmilestone") & "&nbsp;"%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("shortdescription") & ""%></font></td>


  </tr>

  <%	rs.MoveNext
	loop%>
</table>
<%
	Response.Write "<BR><font size=1 face=verdana>Observations Displayed: " & RowCount & "</font>"
	else%>
 <br><table style="Display:none" ID="TableOTS"><tr><td><font size="2">No open observations found for this program.</font></td></tr></table>
<%end if
		rs.Close

	end if 'Skip OTS - No link defined

end if 'Skip OTS - not selected list

'----------------------------------
'######################################
'	DCR Tabs
'######################################
ColumnCount = 0
if strDisplayedList <> "DCR" then
	Response.Write "<Table style=""Display:none"" ID=TableDCR><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
    strSQL = "spListActionItems " & PVID & ",3," & clng(strDcrStatus) & "," & strBiosChange & "," & strSwChange
    rs.Open strSQL,cn,adOpenForwardOnly
  If not(rs.EOF and rs.BOF) then%>
<table ID="TableDCR" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <thead>
	<tr>
	<td onClick="SortTable('TableDCR', 0,1,2);" nowrap width="50" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Number</strong> </font> </td>
	<%if trim(PVID) = "-1" then
		ColumnCount = ColumnCount + 1
	%>
		<td onClick="SortTable( 'TableDCR', <%=ColumnCount%> ,0,2);" nowrap width="120" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Product</strong></font></td>
    <%end if%>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,0,2);" nowrap width="100" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Submitter</strong></font></td>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+2%> ,4,2);" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>ZSRP Ready</strong></font></td>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+3%> ,4,2);" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>AV Required</strong></font></td>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+4%> ,4,2);" nowrap width="150" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Qualification Required</strong></font></td>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+5%> ,0,2);" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Status</strong></font></td>
     <%   ColumnCount = ColumnCount + 5

	if trim(strDcrStatus) <> "1" then 'strStatusID <> "1" then	%>
		<td onClick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,2,2);" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Approved</strong></font></td>
		<td onClick="SortTable( 'TableDCR', <%=ColumnCount+2%> ,2,2);" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();"><strong>Available</strong></font></td>
    <%
		ColumnCount = ColumnCount +2
    end if%>
    <td onClick="SortTable( 'TableDCR', <%=ColumnCount+1%> ,0,2);" nowrap width="100%" bgColor="cornsilk"><strong><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Summary </font></strong></td>
    <td nowrap width="100%" bgColor="cornsilk"><strong><font size="1" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Release </font></strong></td>
</tr>
  </thead>
  <tbody>
  <%do while not rs.EOF  %>
  <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="changerows" LANGUAGE="javascript" onMouseOver="return changerows_onmouseover()" onMouseOut="return changerows_onmouseout()" onClick="return changerows_onclick()" oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
  <%
   if Not IsNull(rs("ZsrpRequired")) Then
        If rs("ZsrpRequired") Then
            If IsNull(rs("ZsrpReadyActualDt")) Then
                If IsNull(rs("ZsrpReadyTargetDt")) Then
                    strZsrpReady = "<span class=""text"">&nbsp;</span>"
                Else                   
                    strZsrpReady = rs("ZsrpReadyTargetDt") & ""
                        If DateDiff("d", NOW(), strZsrpReady) < 0 Then
                            strZsrpReady = "<span style=""color:red;"" class=""text"">" & strZsrpReady & "</span>"
                        Else
                            strZsrpReady = "<span class=""text"">" & strZsrpReady & "</span>"
                        End If
                End If
            Else
                strZsrpReady = "<span class=""text"">Ready</span>"
            End If
        Else
            strZsrpReady = "<span class=""text"">N/A</span>"
        End If
    Else
        strZsrpReady = "<span class=""text"">N/A</span>"
    End If

	if isnull(rs("TargetDate")) then
		strTarget = ""
	else
			strTarget = formatdatetime(rs("TargetDate"),2)
	end if
	if isnull(rs("ActualDate")) then
		strActual = "&nbsp;"
	else
			strActual = formatdatetime(rs("ActualDate"),2)
	end if
	if isnull(rs("AvailableForTest")) then
		strAvailableForTest = "&nbsp;"
	else
		strAvailableForTest = formatdatetime(rs("AvailableForTest"),2)
	end if

    'AV Required and Qualification Required
    Dim strAVRequired
    Dim strQualificationRequired
      strAVRequired = ""
      strQualificationRequired = ""

      If Not IsNull(rs("AVRequired")) Then
         If rs("AVRequired") Then
            strAVRequired = "<span class=""text"">Yes</span>"
         Else
            strAVRequired = "<span class=""text"">No</span>"
         End If
      Else
          strAVRequired = "<span class=""text"">No</span>"
      End If

      If Not IsNull(rs("QualificationRequired")) Then
         If rs("QualificationRequired") Then
            strQualificationRequired = "<span class=""text"">Yes</span>"
         Else
           strQualificationRequired = "<span class=""text"">No</span>"
         End If
      Else
          strQualificationRequired = "<span class=""text"">No</span>"
      End If
     'END AV Required and Qualification Required

	Select case rs("status")
	case 1
		strStatus = "Open"
	case 2
		strStatus = "Closed"
	case 3
		strStatus = "Need Info"
	case 4
		strStatus = "Approved"
	case 5
		strStatus = "Disapproved"
	case 6
		strStatus = "Investigating"
	case else
		strStatus = "N/A"
	end select

	ItemsDisplayed = ItemsDisplayed + 1
  %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ID") & ""%></font></td>
	<%if trim(PVID) = "-1" then%>
		<td valign="top" class="cell"><font size="1" class="text"><%=rs("Product") & ""%></font></td>
	<%end if%>
	<td nowrap valign="top" class="cell"><font size="1" class="text"><%=rs("Submitter") & ""%></font></td>
	<td style="white-space:nowrap; vertical-align:top; font-size:xx-small" class="cell"><%=strZsrpReady %></td>
    <td style="white-space:nowrap; vertical-align:top; font-size:xx-small" class="cell"><%=strAVRequired%></td>
    <td style="white-space:nowrap; vertical-align:top; font-size:xx-small" class="cell"><%=strQualificationRequired %></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strStatus%></font></td>
    <%if trim(strDcrStatus) <> "1" then 'if strStatusID <> "1" then%>
		<td valign="top" class="cell"><font size="1" class="text"><%=strActual%></font></td>
		<td valign="top" class="cell"><font size="1" class="text"><%=strAvailableForTest%></font></td>
	<%end if%>
	<td valign="top" class="cell"><font size="1" class="text"><%=server.HTMLEncode(rs("summary")) & "&nbsp;"%></font></td>
    <td class="cell"><font size="1" class="text"><%=rs("ProductVersionRelease") %></font></td>
  </tr>

  <%	rs.MoveNext
	loop
	%>
	</tbody>
</table>

<%else%>
<table ID="TableDCR" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No open change requests found.</font>
</td></tr></table>

<%end if

'Action Items
rs.Close
end if
'######################################
'	Issue Tabs
'######################################

if strDisplayedList <> "Issue" then
	Response.Write "<Table style=""Display:none"" ID=TableIssue><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
 strSQL = "spListActionItems " & PVID & ",1," & clng(strStatusID)
rs.Open strSQL,cn,adOpenForwardOnly

  If not(rs.EOF and rs.BOF) then%>
<table ID="TableIssue" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <thead>
  <tr>
	<td onClick="SortTable( 'TableIssue', 0,1,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="50" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Number</strong> </font> </td>
    <td onClick="SortTable( 'TableIssue', 1,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="120" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Owner</strong></font></td>
    <td onClick="SortTable( 'TableIssue', 2,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Status</strong></font></td>
    <td onClick="SortTable( 'TableIssue', 3,2,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Target Date</strong></font></td>
    <td onClick="SortTable( 'TableIssue', 4,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100%" bgColor="cornsilk"><strong><font size="1">Summary </font></strong></td>
  </tr>
  </thead>
  <tbody>
  <%

ItemsDisplayed = 0
  do while not rs.EOF  %>
  <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="issuerows" LANGUAGE="javascript" onMouseOver="return issuerows_onmouseover()" onMouseOut="return issuerows_onmouseout()" onClick="return issuerows_onclick()" oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
  <%
	if isnull(rs("TargetDate")) then
		strTarget = ""
	else
			strTarget = formatdatetime(rs("TargetDate"),2)
	end if

	Select case rs("status")
	case 1
		strStatus = "Open"
	case 2
		strStatus = "Closed"
	case 3
		strStatus = "Need Info"
	case 4
		strStatus = "Approved"
	case 5
		strStatus = "Disapproved"
	case 6
		strStatus = "Investigating"
	case else
		strStatus = "N/A"
	end select

	ItemsDisplayed = ItemsDisplayed + 1

  %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ID") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Owner") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strStatus%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strTarget%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("summary") & "&nbsp;"%></font></td>
  </tr>

  <%	rs.MoveNext
	loop
	%></tbody>
</table>


<%else%>
<table ID="TableIssue" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No open issues found for this program.</font>
</td></tr></table>

<%end if


rs.Close
end if


'######################################
'	Opportunity Tabs
'######################################

if strDisplayedList <> "Opportunity" then
	Response.Write "<Table style=""Display:none"" ID=TableOpportunity><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
 strSQL = "spListActionItems " & PVID & ",5," & clng(strStatusID)
rs.Open strSQL,cn,adOpenForwardOnly

  If not(rs.EOF and rs.BOF) then%>
<table ID="TableOpportunity" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <thead>
  <tr>
	<td onClick="SortTable( 'TableOpportunity', 0,1,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="50" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Number</strong> </font> </td>
    <td onClick="SortTable( 'TableOpportunity', 1,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100%" bgColor="cornsilk"><strong><font size="1">Summary </font></strong></td>
    <td onClick="SortTable( 'TableOpportunity', 2,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Metric Impacted</strong></font></td>
    <td onClick="SortTable( 'TableOpportunity', 3,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="70" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Impact</strong></font></td>
    <td onClick="SortTable( 'TableOpportunity', 4,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="70" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Net Affect</strong></font></td>
    <td onClick="SortTable( 'TableOpportunity', 5,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="70" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Status</strong></font></td>
    <td onClick="SortTable( 'TableOpportunity', 6,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="120" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Owner</strong></font></td>
  </tr>
  </thead>
  <tbody>
  <%
	ItemsDisplayed = 0
  do while not rs.EOF  %>
  <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="issuerows" LANGUAGE="javascript" onMouseOver="return issuerows_onmouseover()" onMouseOut="return issuerows_onmouseout()" onClick="return issuerows_onclick()" oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
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

	Select case rs("status")
	case 1
		strStatus = "Open"
	case 2
		strStatus = "Closed"
	case 3
		strStatus = "Need Info"
	case 4
		strStatus = "Approved"
	case 5
		strStatus = "Disapproved"
	case 6
		strStatus = "Investigating"
	case else
		strStatus = "N/A"
	end select

	select case rs("Priority")
	case 1
		strImpact = "1-High"
	case 2
		strImpact = "2-Medium"
	case 3
		strImpact = "3-Low"
	case else
		strImpact = "&nbsp;"
	end select

	if rs("AffectsCustomers") = 1 then
		strAffect = "Positive"
	elseif rs("AffectsCustomers") = 2 then
		strAffect = "Negative"
	else
		strAffect = "&nbsp;"
	end if
	ItemsDisplayed = ItemsDisplayed + 1

  %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ID") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("summary") & "&nbsp;"%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("AvailableNotes") & "&nbsp;"%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strImpact %></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strAffect%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strStatus%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Owner") & ""%></font></td>
  </tr>

  <%	rs.MoveNext
	loop
	%></tbody>
</table>
    

<%else%>
<table ID="TableOpportunity" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No open improvement opportunities found for this program.</font>
</td></tr></table>

<%end if


rs.Close
end if

'######################################
'	Status Tabs
'######################################

if strDisplayedList <> "Status" then
	Response.Write "<Table style=""Display:none"" ID=TableStatus><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
 strSQL = "spListActionItems " & PVID & ",4,0"
rs.Open strSQL,cn,adOpenForwardOnly

  If not(rs.EOF and rs.BOF) then  %>
<table ID="TableStatus" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">

  <thead>
	<td onClick="SortTable( 'TableStatus', 0,1,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="50" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Number</strong> </font> </td>
    <td onClick="SortTable( 'TableStatus', 1,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="120" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Owner</strong></font></td>
    <td onClick="SortTable( 'TableStatus', 2,2,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Date</strong></font></td>
    <td onClick="SortTable( 'TableStatus', 3,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100%" bgColor="cornsilk"><strong><font size="1">Summary </font></strong></td>
  </thead>
  <%
	ItemsDisplayed = 0
  do while not rs.EOF  %>
  <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="issuerows" LANGUAGE="javascript" onMouseOver="return issuerows_onmouseover()" onMouseOut="return issuerows_onmouseout()" onClick="return statusrows_onclick()" oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
  <%
	if isnull(rs("Created")) then
		strTarget = "&nbsp;"
	else
		if trim(rs("LastModified") & "") = "" then
			strTarget = formatdatetime(rs("Created"),2)
		else
			strTarget = formatdatetime(rs("LastModified"),2)
		end if


	end if

	ItemsDisplayed = ItemsDisplayed + 1

  %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ID") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Owner") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strTarget%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("summary") & "&nbsp;"%></font></td>
  </tr>

  <%	rs.MoveNext
	loop
	%>
</table>

<%else%>
<table ID="TableStatus" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No open status notes found for this program.</font>
</td></tr></table>

<%end if


rs.Close
end if

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=Ivory valign=top><TD>"
  MidRow = "</TD><TD valign=top  class=""cell"">"
  MidRowCenter = "</TD><TD align=center valign=top  class=""cell"">"
  PostRow = "</FONT></TD></TR>"

'######################################
'	Deliverables Tabs
'######################################

if strDisplayedList <> "Deliverables" then
	Response.Write "<Table style=""Display:none"" ID=TableDeliverables><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
	strPartNumberCell = ""
'Deliverables

	Response.Write "<span id=spnDeliverables><TABLE ID=TableDeliverables style=""width:100%;Display:none"" border=0><TR><TD>"
	if strDelFilter = "Targeted" or strDelFilter = "" then
		strStatusID = "1"
	else
		strStatusID = "0"
	end if

	if strID = 349 and (strDelType = "SW" or strDelType = "") then
		Response.Write "<font size=1 face=verdana color=red>Titan and Altima are using shared image strategy. Please refer to Titan Excalibur for the up-to-date Software and Image deliverables targets.</font>"
	end if

    if Request("DelFilter") = "Roots" then

        if strDelType = "HW" then
            strDelTypeID = 1
        elseif strDelType = "SW"  or strDelType = "" then
            strDelTypeID = 2
        elseif strDelType = "FW" then
            strDelTypeID = 3
        elseif strDelType = "DOC" then
            strDelTypeID = 4
        else
            strDelTypeID = 2 
        end if
        if trim(strDelTeam) <> "" then
            rs.Open "spListDeliverableMatrixRoots " & PVID & "," & strDelTypeID & "," & clng(strDelTeam),cn,,adOpenStatic
        else
            rs.Open "spListDeliverableMatrixRoots " & PVID & "," & strDelTypeID,cn,,adOpenStatic
        end if

        if rs.EOF and rs.BOF then
		    Response.Write "<TR><TD><font face=Verdana size=2>No deliverable roots found for this product.</font>"
        else
            Response.Write "<table ID=DeliverableTable bgColor=Ivory border=1 borderColor=tan cellPadding=2 cellSpacing=1 width=""100%"">"
	        Response.Write "<THEAD><tr bgcolor=cornsilk>"
            Response.write "<TD onclick=""SortTable( 'DeliverableTable', 0,1,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>ID</b></font></TD>"
            Response.write "<TD onclick=""SortTable( 'DeliverableTable', 1,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Name</b></font></TD>"
            if strFusionRequirements = 1 then
                Response.write "<TD onclick=""SortTable( 'DeliverableTable', 1,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Release</b></font></TD>"
            end if
            Response.write "<TD onclick=""SortTable( 'DeliverableTable', 2,1,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Versions</b></font></TD>"
            Response.write "<TD onclick=""SortTable( 'DeliverableTable', 3,1,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Targeted</b></font></TD>"
            Response.write "<TD onclick=""SortTable( 'DeliverableTable', 4,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Dev.&nbsp;Manager</b></font></TD>"
            'Response.write "<TD onclick=""SortTable( 'DeliverableTable', 5,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>Description</b></font></TD>"
	        Response.Write "</tr></THEAD>"

            Dim strFilterProductRelease
            strFilterProductRelease = trim(Request("ProductRelease"))

            do while not rs.EOF

                if not (strFusionRequirements = 1 and strFilterProductRelease <> "" and LCase(strFilterProductRelease) <> "all" and trim(rs("Releases")) <> "" and InStr( rs("Releases") & "" ,strFilterProductRelease) = 0 ) then   
                'for ProductRelease filter
                
    		    'Need to choose which deliverables the HW teams can access.
			    if trim(strHardwareAccessGroup) = "2" then 'HardwarePM
				    if (blnPlatformDevelopmentPM or blnsuperuser) and (rs("TeamID")= 1 or rs("TeamID")= 13) then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnCommodityPM or blnPlatformDevelopmentPM or blnsuperuser) and rs("TeamID")= 2 then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnProcessorPM or blnPlatformDevelopmentPM or blnsuperuser)  and (rs("TeamID") = 9 or rs("TeamID") = 7)  then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnCommPM or blnPlatformDevelopmentPM or blnsuperuser)  and rs("TeamID") = 3 then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnAccessoryPM or blnServicePM or blnPlatformDevelopmentPM or blnsuperuser)  and rs("TeamID") = 4 then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnServicePM or blnsuperuser or blnPlatformDevelopmentPM)  and rs("TeamID") = 8 then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnVideomemoryPM  or blnPlatformDevelopmentPM or blnsuperuser) and  rs("TeamID") = 10 then
					    strHardwareAccessGroupLocal = "2"
				    elseif (blnGraphicsControllerPM  or blnPlatformDevelopmentPM or blnsuperuser) and  rs("TeamID") = 11 then
					    strHardwareAccessGroupLocal = "2"
                    elseif (blnODMHWPM or blnHWPC) then
                        strHardwareAccessGroupLocal = "2"
                    elseif (blnHWPMRole) then
                        strHardwareAccessGroupLocal = "2"
				    else
					    strHardwareAccessGroupLocal = "0"
				    end if
			    else 'Pilot, Admin, and Accessory are not impacted
				    strHardwareAccessGroupLocal = strHardwareAccessGroup
			    end if
                if rs("TargetCount") = 0 then
                    response.Write "<TR bgcolor=mistyrose"
                else
                    response.Write "<TR"
                end if
                response.write " valign=top class=""ProdID=" & PVID & "&RootID=" & rs("RootID") & "&ID=0"" id=""DelRow" & trim(PVID) & "_" & trim(rs("RootID")) & """ onmouseover=""return Delrows_onmouseover()"" onmouseout=""return Delrows_onmouseout()"" onclick=""return DelMenu(" & PVID & "," & rs("RootID")&",0,0,0," & rs("CategoryID")& "," & rs("TypeID") & ",0," & strHardwareAccessGroupLocal & ",0,0,0,0," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & ", 1," & intReleaseID & ",0,'" & Request("DelFilter") & "')"" oncontextmenu=""DelMenu(" & PVID & "," & rs("RootID")& ",0,0,0," & rs("CategoryID") & "," & rs("TypeID") & ",0," & strHardwareAccessGroupLocal & ",0,0,0,0," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & ", 1, " & intReleaseID & ",0,'" & Request("DelFilter") & "');return false;"">"
                response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("RootID") & "</font></TD>"
                response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("name") & "</font></TD>"

                if strFusionRequirements = 1 then
                    response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("Releases") & "</font></TD>"
                end if
    
                response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("VersionCount") & "</font></TD>"
                response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("TargetCount") & "</font></TD>"
                response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("Devmanager") & "</font></TD>"
               ' response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & rs("Description") & "</font></TD>"
                response.Write "</TR>"

                end if  'for ProductRelease filter

                rs.MoveNext
            loop
            response.Write "</table></td></tr></table>"
        end if
        rs.Close

    else ' = (if Request("DelFilter") <> "Roots" )

        'begin: update InImg column
        set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  4
		set cm.ActiveConnection = cn
		cm.CommandText = "spUpdateDeliverableInImage"
		Set p = cm.CreateParameter("@ProductVersioID", adInteger,  &H0001)
		p.Value = PVID
		cm.Parameters.Append p
				
		Set p = cm.CreateParameter("@CompType", adVarChar,  &H0001, 6)

		if (strDelType="") then
			p.Value = "SW"
		else
			p.Value = strDelType
		end if
		cm.Parameters.Append p
			
		cm.Execute rowschanged
			
		set p = nothing
		set cm = nothing
        
        if (Request("ProductRelease") = "") then
            strProductRelease = ""
        else
            strProductRelease = Request("ProductRelease")
        end if
    'end: update InImg column    
        
	    if strDelTeam <> "" then
		    rs.Open "spListDeliverableMatrix " & PVID & "," & clng(strStatusID) & ",'" & strProductRelease & "'" & "," & clng(strDelTeam),cn,adOpenStatic
	    else
		    rs.Open "spListDeliverableMatrix " & PVID & "," & clng(strStatusID) & ",'" & strProductRelease & "'",cn,adOpenStatic
	    end if
	    if rs.EOF and rs.BOF then
		    Response.Write "<TR><TD><font face=Verdana size=2>No deliverables found for this program.</font>"
	    else
    '		LastBucket = ""

		    Response.Write "<table ID=DeliverableTable bgColor=Ivory border=1 borderColor=tan cellPadding=2 cellSpacing=1 width=""100%"">"
		    Response.Write "<THEAD>"
		    if strStatusID <> 1 then
			    'if strDelType = "HW" then
				    Response.Write "<tr bgcolor=cornsilk><TD width=10 align=middle>&nbsp;</TD><TD onclick=""SortTable( 'DeliverableTable', 1,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" align=middle><font color=black size=1><b>Target</b></font></TD>"
			    'else
			    '	Response.Write "<tr bgcolor=cornsilk><TD width=10 align=middle>&nbsp;</TD><TD align=middle><font color=black size=1><b>Target</b></font></TD>"
			    'end if
			    StartCol = 2
		    else
			    Response.Write "<tr bgcolor=cornsilk><TD width=10 align=middle>&nbsp;</TD>"
			    StartCol = 1
		    end if
		    if strDelType = "HW" then
			
			    Response.write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol & ",1,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10 align=middle><font color=black size=1><b>ID</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 1 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Name</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 2 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=80><font color=black size=1><b>Vendor</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 3 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" align=middle width=40><font color=black size=1><b>HW,FW,Rev</b></font></TD>"
                if strFusionRequirements = 1 then 
					Response.Write "<TD><font color=black size=1><b>Release</b></font></TD>"
					StartCol = StartCol + 1 
				end if
			    Response.Write "<TD  onclick=""SortTable( 'DeliverableTable', " & StartCol + 4 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" align=middle width=80><font color=black size=1><b>Model&nbsp;Number</b></font></TD>"
			    Response.Write "<TD  onclick=""SortTable( 'DeliverableTable', " & StartCol + 5 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" align=middle width=40><font color=black size=1><b>Part&nbsp;Number</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 6 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Restrict</b></font></TD>"
			    'if trim(strTeamName) = "Commodities" then
				    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 7 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Qual&nbsp;Status</b></font></TD>"
				    
				    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 8 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Pilot&nbsp;Status</b></font></TD>"
			    'else
			    if trim(strTeamName) = "Accessories" then
    '				Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 7 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Qual&nbsp;Status</b></font></TD>"
				    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 9 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Accessory&nbsp;Status</b></font></TD>"
			    'else
			    '	Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 7 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Status</b></font></TD>"
			    end if
			    Response.write "</tr>"
		    else
			    if (blnPreinstall or blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM) then
				    strPartNumberCell = "<TD style=""display:"" onclick=""SortTable( 'DeliverableTable', " & StartCol + 3 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=40><font color=black size=1><b>Part&nbsp;Number</b></font></TD>"
				    intPartColumnCount = 1
			    else
				    intPartColumnCount = 0
			    end if

			    Response.write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol & ",1,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10><font color=black size=1><b>ID</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 1 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Name</b></font></TD>"
                Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 2 & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>Language</b></font></TD>"
			    Response.Write strPartNumberCell
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 3 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=50><font color=black size=1><b>Version</b></font></TD>"
                if((strFusionRequirements = 0 and ((strDelFilter = "") or (strDelFilter = "Targeted"))) or (strFusionRequirements = 1 and ((strDelFilter = "") or (strDelFilter = "Targeted")) and (Request("ProductRelease") <> "" ))) then
                    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 4 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=50><font color=black size=1><b>Latest Version</b></font></TD>"
                    StartCol = StartCol + 1
                end if

                if strFusionRequirements = 1 then 
					Response.Write "<TD><font color=black size=1><b>Release</b></font></TD>"
					StartCol = StartCol + 1 
				end if
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 4 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=10><font color=black size=1><b>PIN</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 5 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=40><font color=black size=1><b>Alerts</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 6 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" align=middle width=20><font color=black size=1><b>Img</b></font></TD>"
			    Response.Write "<TD style=""white-space:nowrap;"" onclick=""SortTable( 'DeliverableTable', " & StartCol + 7 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font color=black size=1><b>"
                if Request("ProductRelease") <> "" then
                    Response.Write Request("ProductRelease") & "&nbsp;Target&nbsp;Notes</b></font></TD>"
                else
                    Response.Write "Target&nbsp;Notes</b></font></TD>"
                end if
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 8 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=80><font color=black size=1><b>Distribution</b></font></TD>"
			    Response.Write "<TD onclick=""SortTable( 'DeliverableTable', " & StartCol + 9 + intPartColumnCount & ",0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" width=80><font color=black size=1><b>Images</b></font></TD>"
			    Response.Write "</tr>"
		    end if
		    Response.write "</THEAD>"

		    i=0
		    j=0

            if strDelType <> "HW" then
                rsOTS.CursorType = adOpenStatic '3
                rsOTS.CursorLocation = adUseClient '3
		        rsOTS.open "spListObservationCountByRoot",cn
                rsOTS.Fields("RootID").Properties("Optimize") = True
            end if

            if (strDelType = "HW" and strFusionRequirements = 1) then
                strSql = "select pd.DeliverableVersionID, pvr.Name, pdr.TestStatusID, TestDate = isnull(pdr.TestDate,''), RiskRelease = isnull(pdr.RiskRelease, 0), pdr.PilotStatusID, pdr.PilotDate, SupplyChainRestriction=isnull(pdr.SupplyChainRestriction,0), ConfigurationRestriction=isnull(pdr.ConfigurationRestriction,0), " &_
                         "pdr.AccessoryStatusID, pdr.accessorydate, pdr.accessoryleveraged " &_        
                         "from Product_Deliverable pd " &_
                         "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID and pdr.targeted = pd.targeted " &_
                         "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                         "where pd.ProductVersionID= " & PVID & " order by pd.DeliverableVersionID, pvr.id desc"                                  
                rs2.CursorType = adOpenStatic '3
                rs2.CursorLocation = adUseClient '3
				rs2.open strSql, cn
                rs2.Fields("DeliverableVersionID").Properties("Optimize") = True
            end if

		    do while not rs.EOF
			    if ((trim(rs("Type")) = "Hardware" and strDelType="HW") or (trim(rs("Type")) = "Software" and (strDelType="SW" or strDelType="")) or (trim(rs("Type")) = "Firmware" and strDelType="FW") or (trim(rs("Type")) = "Documentation" and strDelType="DOC")) and ((strDelFilter = "Image" and rs("InImage")) or strDelFilter <> "Image") and ((strDelFilter = "PINImage" and rs("InPINImage")) or strDelFilter <> "PINImage") then '****
				   
				    strDistribution = ""
				    strImages = ""
				    if rs("ID") = 0 then
					    strDistribution = "TBD"
					    'i=i+1
					    j=j + 1
				    else
					    i=i+1
					    if rs("Preinstall") then
						    strDistribution = ",Preinstall"
					    end if
					    if rs("Preload") then
						    strDistribution = strDistribution & ",Preload"
					    end if
					    if rs("DropInBox") then
						    strDistribution = strDistribution & ",DIB"
					    end if
					    if rs("Web") then
						    strDistribution = strDistribution & ",Web"
					    end if
					    if rs("SelectiveRestore") then
						    strDistribution = strDistribution & ",Selective Restore"
					    end if
					    if rs("ARCD") then
						    strDistribution = strDistribution & ",DRCD"
					    end if
					    if rs("DRDVD") then
						    strDistribution = strDistribution & ",DRDVD"
					    end if
					    if rs("RACD_Americas") then
						    strDistribution = strDistribution & ",RACD-Americas"
					    end if
					    if rs("RACD_APD") then
						    strDistribution = strDistribution & ",RACD-APD"
					    end if
					    if rs("RACD_EMEA") then
						    strDistribution = strDistribution & ",RACD-EMEA"
					    end if
					    if rs("OSCD") then
						    strDistribution = strDistribution & ",OSCD"
					    end if
					    if rs("DocCD") then
						    strDistribution = strDistribution & ",DocCD"
					    end if
					    if trim(rs("Patch") & "") <> "0" and trim(rs("Patch") & "") <> "" then
						    strDistribution = strDistribution & ",Patch"
					    end if
                        if rs("RCDOnly") then
						    strDistribution = strDistribution & ",RCDOnly"
					    end if

					    if strDistribution <> "" then
						    strDistribution = mid(strDistribution,2)
					    else
						    strDistribution = "&nbsp;"
					    end if
				    end if

				    if rs("Type") = "Hardware" then
					    strImages = "<b>-</b>"
					    strDistribution = "<b>-</b>"
				    else
					    strImages = rs("ImageSummary") & ""
					    if strImages = "" then
						    strImages = "ALL"
					    end if
				    end if

				    if  (rs("Preinstall") or rs("Preload") or rs("SelectiveRestore")) and rs("InImage") then
					    strInImageDisplay = "Yes"
				    elseif (rs("Preinstall") or rs("Preload") or rs("SelectiveRestore")) then
					    strInImageDisplay = "&nbsp;"
				    else
					    strInImageDisplay = "<b>-</b>"
				    end if
				    if instr(rs("Location") & "","Workflow Complete")> 0 then
					    intWorkflowComplete=1
				    else
					    intWorkflowComplete=0
				    end if

				    Version = rs("Version") & ""
				    if rs("Revision") <> "" then
					    Version  = Version & "," &  rs("Revision")
				    end if
				    if rs("Pass") <> "" then
					    Version  = Version & "," &  rs("Pass")
				    end if
				    if rs("Targeted") then
					    strTargeted = "1"
				    else
					    strTargeted = "0"
				    end if
				    if rs("InImage") then
					    strInImage = "1"
				    else
					    strInImage = "0"
				    end if
				    'Need to choose which deliverables the HW teams can access.
				    if trim(strHardwareAccessGroup) = "2" then 'HardwarePM
					    if (blnPlatformDevelopmentPM or blnsuperuser) and (rs("TeamID")= 1 or rs("TeamID")= 13) then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnCommodityPM or blnPlatformDevelopmentPM or blnsuperuser) and rs("TeamID")= 2 then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnProcessorPM or blnPlatformDevelopmentPM or blnsuperuser)  and (rs("TeamID") = 9 or rs("TeamID") = 7)  then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnCommPM or blnPlatformDevelopmentPM or blnsuperuser)  and rs("TeamID") = 3 then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnAccessoryPM or blnServicePM or blnPlatformDevelopmentPM or blnsuperuser)  and rs("TeamID") = 4 then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnServicePM or blnsuperuser or blnPlatformDevelopmentPM)  and rs("TeamID") = 8 then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnVideomemoryPM  or blnPlatformDevelopmentPM or blnsuperuser) and  rs("TeamID") = 10 then
						    strHardwareAccessGroupLocal = "2"
					    elseif (blnGraphicsControllerPM  or blnPlatformDevelopmentPM or blnsuperuser) and  rs("TeamID") = 11 then
						    strHardwareAccessGroupLocal = "2"
                        elseif (blnODMHWPM or blnHWPC) then
                            strHardwareAccessGroupLocal = "2"
                        elseif(blnHWPMRole) then
                            strHardwareAccessGroupLocal = "2"
					    else
						    strHardwareAccessGroupLocal = "0"
					    end if
					    'Response.Write strHardwareAccessGroupLocal
                        'Response.Write "<tr><td>TeamID: " & rs("TeamID") & "</td></tr>"
				    else 'Pilot, Admin, and Accessory are not impacted
					    strHardwareAccessGroupLocal = strHardwareAccessGroup
				    end if

				     'Response.Write ">" & strHardwareAccessGroup & "_" & strHardwareAccessGroupLocal
				    if rs("ID") = 0 then
					    if strDelType = "HW" then
						    Response.Write   "<tr bgcolor=mistyrose valign=top class=""ProdID=" & PVID & "&RootID=" & rs("RootID") & "&ID=" & rs("ID") & ", deliverable" & trim(PVID) & "_" & trim(rs("RootID")) & "_" & rs("ID") & """ id=""DelRow" & trim(PVID) & "_" & trim(rs("RootID")) & """ LANGUAGE=javascript onmouseover=""return Delrows_onmouseover()"" onmouseout=""return Delrows_onmouseout()"" onclick=""return DelMenu(" & PVID & "," & rs("RootID")&"," & rs("ID") & "," & strTargeted &   "," & strInImage & "," & rs("CategoryID")& "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & "," & blnDEVTestLead & "," & ServicePMAccess& "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "')"" oncontextmenu=""DelMenu(" & PVID & "," & rs("RootID")& "," & rs("ID") &  "," & strTargeted & "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & ","& blnDEVTestLead & "," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "');return false;"">"
					    else
						    Response.Write   "<tr bgcolor=mistyrose valign=top class=""ProdID=" & PVID & "&RootID=" & rs("RootID") & "&ID=" & rs("ID") & ", deliverable" & trim(PVID) & "_" & trim(rs("RootID")) & "_" & rs("ID") & """ id=""DelRow" & trim(PVID) & "_" & trim(rs("RootID")) & """ LANGUAGE=javascript onmouseover=""return Delrows_onmouseover()"" onmouseout=""return Delrows_onmouseout()"" onclick=""return DelMenu(" & PVID & "," & rs("RootID")&"," & rs("ID") & "," & strTargeted &   "," & strInImage & "," & rs("CategoryID")& "," & rs("TypeID") & "," & intWorkflowComplete & ",0,0,0,0," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & request("DelFilter") & "')"" oncontextmenu=""DelMenu(" & PVID & "," & rs("RootID")& "," & rs("ID") &  "," & strTargeted & "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & ",0,0,0,0," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "');return false;"">"
					    end if
					    Response.Write "<TD align=center class=""cell"">&nbsp;</TD>"
				    else
                        if (rs("VersionActive")) then
					        Response.Write   "<tr bgcolor=ivory valign=top class=""ProdID=" & PVID & "&RootID=" & rs("RootID") & "&ID=" & rs("ID") & ", deliverable" & trim(PVID) & "_" & trim(rs("RootID")) & "_" & rs("ID") & """ id=""DelRow"" LANGUAGE=javascript onmouseover=""return Delrows_onmouseover()"" onmouseout=""return Delrows_onmouseout()"" onclick=""return DelMenu(" & PVID & "," & rs("RootID")&"," & rs("ID") & "," & strTargeted &   "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & "," & blnDEVTestLead & "," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "')"" oncontextmenu=""DelMenu(" & PVID & "," & rs("RootID")& "," & rs("ID") &  "," & strTargeted &   "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & ","& blnDEVTestLead & "," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "');return false;"">"
                        else
                            Response.Write   "<tr bgcolor=Gainsboro valign=top class=""ProdID=" & PVID & "&RootID=" & rs("RootID") & "&ID=" & rs("ID") & ", deliverable" & trim(PVID) & "_" & trim(rs("RootID")) & "_" & rs("ID") & """ id=""DelRow"" LANGUAGE=javascript onmouseover=""return Delrows_onmouseover()"" onmouseout=""return Delrows_onmouseout()"" onclick=""return DelMenu(" & PVID & "," & rs("RootID")&"," & rs("ID") & "," & strTargeted &   "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & "," & blnDEVTestLead & "," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "')"" oncontextmenu=""DelMenu(" & PVID & "," & rs("RootID")& "," & rs("ID") &  "," & strTargeted &   "," & strInImage & "," & rs("CategoryID") & "," & rs("TypeID") & "," & intWorkflowComplete & "," & strHardwareAccessGroupLocal & "," & blnSETestLead & "," & blnODMTestLead & "," & blnWWANTestLead & "," & blnDEVTestLead & "," & ServicePMAccess & "," & strFusion & "," & strFusionRequirements & "," & Abs(CInt(rs("VersionActive"))) & "," & intReleaseID & "," & rs("BusinessSegmentID") & ",'" & Request("DelFilter") & "');return false;"">"
                        end if
					    if instr(rs("Location")&"","Core Team")>0 or instr(rs("Location")&"","Development")>0 or (rs("VersionActive")=0) then
						    Response.Write "<TD align=center class=""cell"">&nbsp;</TD>"
					    else
						    Response.Write "<TD align=center class=""cell""><INPUT class=""check"" style=""WIDTH:16;HEIGHT:16"" type=""checkbox"" id=chkVersion name=chkVersion value=""" & rs("ID") & """></TD>"
					    end if
				    end if


				    if rs("targeted") and strStatusID <> 1 then
					    response.write "<TD class=""cell"" align=middle><Font size=1 face=verdana  class=""text"" ID=""HWTargetCell" & trim(rs("ID")) & """>Yes</font></TD>"
				    '	response.write "<TD class=""cell""><IMG class=""text"" SRC=""images/target.gif""></TD>"
				    elseif strStatusID <> 1 then
				    response.write "<TD class=""cell"" align=middle><Font size=1 face=verdana  class=""text"" ID=""HWTargetCell" & trim(rs("ID")) & """>&nbsp;</font></TD>"
				    '	response.write "<TD class=""cell"">&nbsp;<IMG style=""Display:none"" class=""text"" SRC=""images/target.gif""></TD>"
				    end if
				    if rs("ID") & "" = "" or rs("ID") & "" = "0" then
					    strDisplayID = "0"
				    else
					    strDisplayID = rs("ID") & ""
				    end if

				    Response.Write "<TD align=center class=""cell""><font size=1 face=verdana class=""text"">" & strDisplayID & "</font>"
				    Response.Write "<INPUT type=""hidden"" id=Path" & rs("ID") & " name=Path" & rs("ID") & " value=""" & rs("ImagePath") & """></TD>"
				    Response.Write "<TD class=""cell""><font size=1 face=verdana class=""text"">" & server.htmlencode(rs("DeliverableName")) & "<INPUT type=""hidden"" id=txtDelName" & trim(PVID & "_" & rs("ID")) & " name=txtDelName value=""" & server.htmlencode(rs("DeliverableName")) & " " & Version & """></font>"
				    if strDelType = "HW" then
					    Response.Write  midrow & "<font size=1 face=verdana  class=""text"">" & rs("Vendor") & "&nbsp;"
					    Response.Write midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"
                             
                        statusBGColor = ""
                        TestStatusText = ""
                        PilotStatusText = ""
						statusPilotBGColor = ""
                        strSCRestrict = ""
                        statusAccessoryBGColor = ""
					    AccessoryStatusText = ""
                        if strFusionRequirements = 1 then 
                            Response.Write  midrow & "<font size=1 face=verdana  class=""text"">" & rs("Releases") & "</font>&nbsp;"

                            if instr(rs("Location")&"","Development")>0 then
						        TestStatusText = "Development"
						        statusBGColor = "bgcolor=yellow"
                                PilotStatusText = "Development"
						        statusPilotBGColor = "bgcolor=yellow"
                                AccessoryStatusText = "Development"
						        statusAccessoryBGColor = "bgcolor=yellow"
					        elseif instr(rs("Location")&"","Core Team")>0 then
						        TestStatusText = "Core Team"
						        statusBGColor = "bgcolor=yellow"
                                PilotStatusText = "Core Team"
						        statusPilotBGColor = "bgcolor=yellow"
                                AccessoryStatusText = "Core Team"
						        statusAccessoryBGColor = "bgcolor=yellow"
					        elseif trim(rs("Location")&"")="" or trim(rs("Location")&"")="TBD" then
						        TestStatusText = "TBD"
                                PilotStatusText = "TBD"
                                AccessoryStatusText = "TBD"
                            else 
			                    
                                rs2.filter = ""
                                rs2.filter = "DeliverableVersionID=" & rs("ID")
                                    
                                TestStatusText = "" 
                                PilotStatusText = ""
                                AccessoryStatusText = ""

                                dim ReleasePilotStatusText
                                dim ReleaseStatusText 
                                dim ReleaseAccessoryStatusText
                                dim RestrictionText            
				                do while not rs2.EOF
                                    if rs2("TestStatusID") = 1 and rs("Commodity") then
						                ReleaseStatusText = "Investigating"
					                elseif rs2("TestStatusID") = 1 and not rs("Commodity") then
						                ReleaseStatusText = "N/A"
					                elseif rs2("TestStatusID") = 3 then
						                ReleaseStatusText = rs2("TestDate") & "&nbsp;"
					                elseif rs2("TestStatusID") = 5 then
						                if rs2("RiskRelease") then
							                ReleaseStatusText = "Risk&nbsp;Release"
						                else
							                ReleaseStatusText = "QComplete"
						                end if
					                elseif rs2("TestStatusID") = 4 then
						                ReleaseStatusText = "Leverage"
					                elseif rs2("TestStatusID") = 6 then
						                ReleaseStatusText = "Dropped"
					                elseif rs2("TestStatusID") = 7 then
						                ReleaseStatusText = "QHold"
					                elseif rs2("TestStatusID") = 8 then
						                ReleaseStatusText = "Validate&nbsp;Only"
					                elseif rs2("TestStatusID") = 10 then
						                ReleaseStatusText = "Fail"
					                elseif rs2("TestStatusID") = 11 then
						                ReleaseStatusText = "OOC"
					                elseif rs2("TestStatusID") = 15 then
						                ReleaseStatusText = "Planning"
					                elseif rs2("TestStatusID") = 16 then
						                ReleaseStatusText = "FCS"
					                elseif rs2("TestStatusID") = 17 then
						                ReleaseStatusText = "Validated"
					                elseif rs2("TestStatusID") = 0 then
						                ReleaseStatusText = "Not&nbsp;Used"
					                elseif rs2("TestStatusID") = 18 then
						                ReleaseStatusText = "Service&nbsp;Only"
					                else
						                ReleaseStatusText = "TBD" 
					                end if  
    
                                    if TestStatusText <> "" then
                                        TestStatusText = TestStatusText & " <br />"
                                    end if
                                        
                                    TestStatusText = TestStatusText & " [" & rs2("Name") & "]: " & ReleaseStatusText  
    
                                    if rs2("PilotStatusID") = 0 then
						                ReleasePilotStatusText = "Not&nbsp;Required"
					                elseif rs2("PilotStatusID") = 1 then
						                ReleasePilotStatusText = "Planning"
					                elseif rs2("PilotStatusID") = 2 then
						                if isnull(rs2("PilotDate")) then
							                ReleasePilotStatusText = "Scheduled"
						                else
							                ReleasePilotStatusText = rs2("PilotDate") & "&nbsp;"
						                end if
					                elseif rs2("PilotStatusID") = 3 then
						                ReleasePilotStatusText = "On&nbsp;Hold"
					                elseif rs2("PilotStatusID") = 4 then
						                ReleasePilotStatusText = "Canceled"
					                elseif rs2("PilotStatusID") = 5 then
						                ReleasePilotStatusText = "Failed"
					                elseif rs2("PilotStatusID") = 6 then
						                ReleasePilotStatusText = "Complete"
					                elseif rs2("PilotStatusID") = 7 then
						                ReleasePilotStatusText = "Factory&nbsp;Hold"
					                else
						                ReleasePilotStatusText = "TBD"
					                end if

                                    if PilotStatusText <> "" then
                                        PilotStatusText = PilotStatusText & " <br />"
                                    end if
                                                                                      
                                    PilotStatusText = PilotStatusText & " [" & rs2("Name") & "]: " & ReleasePilotStatusText 
                         
                                    ' Added by PMV Pandian on 11-May 2017 for the Task 129265 to populate restrictions based on the releases       - Begin
                                        
                                        if RestrictionText <> "" then
                                            RestrictionText = RestrictionText & " <br />"
                                        end if   

                                        if rs2("ConfigurationRestriction") or rs2("SupplyChainRestriction") then
						                    RestrictionText = RestrictionText & " [" & rs2("Name") & "]: " & "Yes"
                                        end if
                                 ' Added by PMV Pandian on 11-May 2017 for the Task 129265 to populate restrictions based on the releases       - End 
                                   

                                 'Accessory Status
                                    if trim(strTeamName) = "Accessories" then
					                    
					                    if rs2("AccessoryStatusID") = 0 then
						                    ReleaseAccessoryStatusText = "Not&nbsp;Used"
					                    elseif rs2("AccessoryStatusID") = 1 then
						                    ReleaseAccessoryStatusText = "Planning"
					                    elseif rs2("AccessoryStatusID") = 2 then
						                    if isnull(rs2("AccessoryDate")) or (trim(rs2("AccessoryLeveraged") & "") = "True" and rs2("TestStatusID") <> 3) then
							                    ReleaseAccessoryStatusText = "Scheduled"
						                    else
							                    ReleaseAccessoryStatusText = rs2("AccessoryDate") & "&nbsp;"
						                    end if
					                    elseif rs2("AccessoryStatusID") = 3 then
						                    ReleaseAccessoryStatusText = "On&nbsp;Hold"
					                    elseif rs2("AccessoryStatusID") = 4 then
						                    ReleaseAccessoryStatusText = "Canceled"
					                    elseif rs2("AccessoryStatusID") = 5 then
						                    ReleaseAccessoryStatusText = "Failed"
					                    elseif rs2("AccessoryStatusID") = 6 then
						                    ReleaseAccessoryStatusText = "Complete"
                                        elseif rs2("AccessoryStatusID") = 7 then
						                    ReleaseAccessoryStatusText = "Leveraged"
					                    else
						                    ReleaseAccessoryStatusText = "TBD" '& rs("location")
					                    end if

                                        if AccessoryStatusText <> "" then
                                            AccessoryStatusText = AccessoryStatusText & " <br />"
                                        end if
                                                                                      
                                        AccessoryStatusText = AccessoryStatusText & " [" & rs2("Name") & "]: " & ReleaseAccessoryStatusText 
                                    end if


                                    rs2.MoveNext
                                    
		                        loop

                                if RestrictionText ="" then
                                    RestrictionText = "&nbsp;"
                                end if
                                strSCRestrict = RestrictionText
                                
                                RestrictionText = ""
                            end if
                        else
                            if instr(rs("Location")&"","Development")>0 then
						        TestStatusText = "Development"
						        statusBGColor = "bgcolor=yellow"
					        elseif instr(rs("Location")&"","Core Team")>0 then
						        TestStatusText = "Core Team"
						        statusBGColor = "bgcolor=yellow"
					        elseif trim(rs("Location")&"")="" or trim(rs("Location")&"")="TBD" then
						        TestStatusText = "TBD"                        
					        elseif rs("TestStatusID") = 1 and rs("Commodity") then
						        TestStatusText = "Investigating"
					        elseif rs("TestStatusID") = 1 and not rs("Commodity") then
						        TestStatusText = "N/A"
					        elseif rs("TestStatusID") = 3 then
						        TestStatusText = rs("TestDate") & "&nbsp;"
					        elseif rs("TestStatusID") = 5 then
						        if rs("RiskRelease") then
							        TestStatusText = "Risk&nbsp;Release"
						        else
							        TestStatusText = "QComplete"
						        end if
					        elseif rs("TestStatusID") = 4 then
						        TestStatusText = "Leverage"
					        elseif rs("TestStatusID") = 6 then
						        TestStatusText = "Dropped"
					        elseif rs("TestStatusID") = 7 then
						        TestStatusText = "QHold"
					        elseif rs("TestStatusID") = 8 then
						        TestStatusText = "Validate&nbsp;Only"
					        elseif rs("TestStatusID") = 10 then
						        TestStatusText = "Fail"
					        elseif rs("TestStatusID") = 11 then
						        TestStatusText = "OOC"
					        elseif rs("TestStatusID") = 15 then
						        TestStatusText = "Planning"
					        elseif rs("TestStatusID") = 16 then
						        TestStatusText = "FCS"
					        elseif rs("TestStatusID") = 17 then
						        TestStatusText = "Validated"
					        elseif rs("TestStatusID") = 0 then
						        TestStatusText = "Not&nbsp;Used"
					        elseif rs("TestStatusID") = 18 then
						        TestStatusText = "Service&nbsp;Only"
					        else
						        TestStatusText = "TBD" '& rs("location")
					        end if    
    
                            'Pilot Status
					        if instr(rs("Location")&"","Development")>0 then
						        PilotStatusText = "Development"
						        statusPilotBGColor = "bgcolor=yellow"
					        elseif instr(rs("Location")&"","Core Team")>0 then
						        PilotStatusText = "Core Team"
						        statusPilotBGColor = "bgcolor=yellow"
					        elseif trim(rs("Location")&"")="" or trim(rs("Location")&"")="TBD" then
						        PilotStatusText = "TBD"
					        elseif rs("PilotStatusID") = 0 then
						        PilotStatusText = "Not&nbsp;Required"
					        elseif rs("PilotStatusID") = 1 then
						        PilotStatusText = "Planning"
					        elseif rs("PilotStatusID") = 2 then
						        if isnull(rs("PilotDate")) then
							        PilotStatusText = "Scheduled"
						        else
							        PilotStatusText = rs("PilotDate") & "&nbsp;"
						        end if
					        elseif rs("PilotStatusID") = 3 then
						        PilotStatusText = "On&nbsp;Hold"
					        elseif rs("PilotStatusID") = 4 then
						        PilotStatusText = "Canceled"
					        elseif rs("PilotStatusID") = 5 then
						        PilotStatusText = "Failed"
					        elseif rs("PilotStatusID") = 6 then
						        PilotStatusText = "Complete"
					        elseif rs("PilotStatusID") = 7 then
						        PilotStatusText = "Factory&nbsp;Hold"
					        else
						        PilotStatusText = "TBD" '& rs("location")
					        end if 
    
                            if rs("ConfigurationRestriction") or rs("SupplyChainRestriction") then
						        strSCRestrict = "Yes"
                            else
                                strSCRestrict = "&nbsp;"
					        end if 
    
                            'Accessory Status
                            if trim(strTeamName) = "Accessories" then
                                statusAccessoryBGColor = ""
					            AccessoryStatusText = ""
					            if instr(rs("Location")&"","Development")>0 then
						            AccessoryStatusText = "Development"
						            statusAccessoryBGColor = "bgcolor=yellow"
					            elseif instr(rs("Location")&"","Core Team")>0 then
						            AccessoryStatusText = "Core Team"
						            statusAccessoryBGColor = "bgcolor=yellow"
					            elseif trim(rs("Location")&"")="" or trim(rs("Location")&"")="TBD" then
						            AccessoryStatusText = "TBD"
					            elseif rs("AccessoryStatusID") = 0 then
						            AccessoryStatusText = "Not&nbsp;Used"
					            elseif rs("AccessoryStatusID") = 1 then
						            AccessoryStatusText = "Planning"
					            elseif rs("AccessoryStatusID") = 2 then
						            if isnull(rs("AccessoryDate")) or (trim(rs("AccessoryLeveraged") & "") = "True" and rs("TestStatusID") <> 3) then
							            AccessoryStatusText = "Scheduled"
						            else
							            AccessoryStatusText = rs("AccessoryDate") & "&nbsp;"
						            end if
					            elseif rs("AccessoryStatusID") = 3 then
						            AccessoryStatusText = "On&nbsp;Hold"
					            elseif rs("AccessoryStatusID") = 4 then
						            AccessoryStatusText = "Canceled"
					            elseif rs("AccessoryStatusID") = 5 then
						            AccessoryStatusText = "Failed"
					            elseif rs("AccessoryStatusID") = 6 then
						            AccessoryStatusText = "Complete"
					            else
						            AccessoryStatusText = "TBD" '& rs("location")
					            end if
                            end if                   					    
                        end if      
                        
                        if strSCRestrict = "" then 
                            strSCRestrict = "&nbsp;"
                        end if
           
					    Response.Write  midrow & "<font size=1  face=verdana  class=""text"">" & rs("ModelNUmber") & "&nbsp;" & midrow  & "<font ID=""PartCell" & rs("ID") & "_" &  PVID & """ size=1 face=verdana  class=""text"">" & rs("VersionPartNumber") & "&nbsp;" & "</td><TD width='7%' style='font-size:1em; font-family:verdana' id=""RestrictedCell" & trim(rs("ID")) & """>" & strSCRestrict & "</TD>"
				    
					    if  trim(rs("DeveloperNotificationStatus")) = "2" then
						    StatusBGColor = " bgcolor=#ff3333 "
					    end if

					    Response.write "<TD " & StatusBGColor & " valign=top  class=""cell"" width='10%'><font size=1 face=verdana  class=""text"" ID=""HWStatusCell" & trim(rs("ID")) & """>" & TestStatusText & "</font>"
					    Response.Write "<font size=1 face=verdana ID=""IMGCell" & rs("ID") & "_" &  PVID & """ style=""Display:none"" class=""text"">-</font></td>"

					    'if trim(strTeamName) = "Commodities" then
					    Response.write "<TD " & StatusPilotBGColor & " valign=top  class=""cell"" width='10%'><font size=1 face=verdana  class=""text"" ID=""PilotStatusCell" & trim(rs("ID")) & """>" & PilotStatusText & "</font></td>"
					    'else

					    if trim(strTeamName) = "Accessories" then
						    Response.write "<TD " & StatusAccessoryBGColor & " valign=top  class=""cell"" width='10%'><font size=1 face=verdana  class=""text"" ID=""AccessoryStatusCell" & trim(rs("ID")) & """>" & AccessoryStatusText & "</font></td>"
					    end if

					    Response.write postrow

    '				Response.Write midrow & "<font ID=""PartCell" & rs("ID") & "_" &  PVID & """size=1 face=verdana  class=""text"">" & rs("PartNumber") & "</font><font size=1>&nbsp;"
				    else
                        if trim(rs("Location") & "") <> "Workflow Complete" then
						    strLocation = replace(replace(rs("Location")& "","Workflow Complete","Complete")," ","&nbsp;")
						    statusBGColor = "bgcolor=yellow"
                        else
                            strLocation = "&nbsp;"
					    end if
					    if isdate(rs("EOL")) and strStatus <> "Post-Production"  and strStatus <> "Inactive" then
						    if datediff("d",rs("EOL"),now) < 365 then
						        if strLocation = "&nbsp;" then
    						        strLocation = "Use Until: " & rs("EOL")
					            else
    						        strLocation = strLocation & "<BR>Use Until: " & rs("EOL")
					            end if
					       end if
					    end if
					    if trim(rs("DeveloperNotificationStatus") & "") = "2" then
						    if strLocation = "&nbsp;" then
						        strLocation = "Dev:&nbsp;Disapproved"
					        else
						        strLocation = strLocation & "<BR>Dev:&nbsp;Disapproved"
					        end if
					    end if
					    if trim(rs("DeveloperNotification") & "") = "1" and trim(rs("DeveloperNotificationStatus") & "") = "0" and trim(strLocation) <> "TBD" then
						    if strLocation = "&nbsp;" then
						        strLocation = "Dev:&nbsp;Awaiting&nbsp;Approval"
					        else
						        strLocation = strLocation & "<BR>Dev:&nbsp;Awaiting&nbsp;Approval"
					        end if
					    end if
					    if rs("Targeted") and (trim(strDistribution) = "" or trim(strDistribution) = "&nbsp;") then
						    if strLocation = "&nbsp;" then
						        strLocation = "No&nbsp;Distributions"
					        else
						        strLocation = strLocation & "<BR>No&nbsp;Distributions"
					        end if
					    end if
					    if trim(rs("LevelID") & "") = "3" or trim(rs("LevelID") & "") = "9" or trim(rs("LevelID") & "") = "10" or trim(rs("LevelID") & "") = "11" then
						    if strLocation = "&nbsp;" then
						        strLocation = "Alpha"
					        else
						        strLocation = strLocation & "<BR>Alpha"
					        end if
					    elseif trim(rs("LevelID") & "") = "4" or trim(rs("LevelID") & "") = "12" or trim(rs("LevelID") & "") = "13" or trim(rs("LevelID") & "") = "14"   then
						    if strLocation = "&nbsp;" then
						        strLocation = "Beta"
					        else
						        strLocation = strLocation & "<BR>Beta"
					        end if
					    elseif trim(rs("CertificationStatus") & "") <> "2" and trim(rs("CertificationStatus") & "") <> "4" and trim(rs("CertRequired") & "") = "1" and (trim(rs("LevelID") & "") = "7" or trim(rs("LevelID") & "") = "15" or trim(rs("LevelID") & "") = "16" or trim(rs("LevelID") & "") = "17"  or trim(rs("LevelID") & "") = "18") then 'RC or GM, Requires WHQL, WHQL Status <> 2 or 4
						    if strLocation = "&nbsp;" then
						        strLocation = "WHQL&nbsp;Issue"
					        else
						        strLocation = strLocation & "<BR>WHQL&nbsp;Issue"
					        end if
					    end if
					    if strLocation = "TBD" then
					        strLocation = "Choose&nbsp;Version"
					    end if

                        rsOTS.Filter = ""
                        rsOTS.Filter = "RootID=" & rs("RootID")

					    if not rsOTS.EOF then
						    if strLocation = "&nbsp;" then
    					        strLocation = "OTS Alerts: " & rsOTS("OTSCount")
                            else
	    				        strLocation = strLocation & "<BR>OTS Alerts: " & rsOTS("OTSCount")
		                    end if
					    end if

					    strInternalRev = trim(rs("PreinstallInternalRev") & "")
					    if strInternalRev = "" then
					        strInternalRev = "1"
					    elseif strInternalRev = "0" then
					        strInternalRev = "-"
					    end if

                        Response.Write "</TD><TD style=""display:"" valign=top  class=""cell""><font ID=""LanguageCell" & rs("ID") & "_" &  PVID & """size=1 face=verdana  class=""text"">" & server.htmlencode(rs("Language") & "") & "&nbsp;"
                        if (blnPreinstall or blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM) and (rs("ID")) > 0 and (rs("VersionActive") = 0) then
                             Response.Write "</TD><TD style=""display:"" valign=top  class=""cell""><font ID=""PartCell" & rs("ID") & "_" &  PVID & """size=1 face=verdana  class=""text"">" & server.htmlencode(rs("versionPartNumber") & "") & "&nbsp; <font size=1 color=red face=verdana>(inactive)</font>"
					    elseif blnPreinstall or blnAdministrator or blnPreinstallPM or blnOdmPreinstallPM then
						    Response.Write "</TD><TD style=""display:"" valign=top  class=""cell""><font ID=""PartCell" & rs("ID") & "_" &  PVID & """size=1 face=verdana  class=""text"">" & server.htmlencode(rs("versionPartNumber") & "") & "&nbsp;"
					    end if
					    Response.Write midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"
                     
                        if((strFusionRequirements = 0 and ((strDelFilter = "") or (strDelFilter = "Targeted"))) or (strFusionRequirements = 1 and ((strDelFilter = "") or (strDelFilter = "Targeted")) and (Request("ProductRelease") <> "" ))) then
                            if (rs("LatestVersion") = 0) then
                                strLatestVer = "N"& " (" & rs("MaxPartNo") & ")"
                            elseif (rs("LatestVersion") = 1) then
                                strLatestVer = "Y"
                            elseif (IsNull(rs("LatestVersion")) and rs("Targeted") = 1 and rs("VersionPartNumber") <> "N/A") then
                                strLatestVer="N"
                            else
                                strLatestVer = "&nbsp;"
                            end if
                            Response.Write  "</TD><TD nowrap><font size=1>" & strLatestVer & "</font>"
                        end if
                            
                        if strFusionRequirements = 1 then 
					       Response.Write midrow & "<font size=1 face=verdana  class=""text"">" & rs("Releases") & "</font>&nbsp;"
				        end if
					    Response.Write midrowcenter & "<font size=1 face=verdana class=""text"">" & strInternalRev & "</font>"
					    'Response.Write "<TD align=center class=""cell""><font size=1 face=verdana class=""text"">" & strDisplayID & "</font></TD><TD class=""cell""><INPUT type=""hidden"" id=txtDelName" & trim(PVID & "_" & rs("ID")) & " name=txtDelName value=""" & rs("DeliverableName") & " " & Version & """><font size=1 face=verdana class=""text"">" & rs("DeliverableName") & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"
					    if strLocation = "&nbsp;" then 'or strLocation = "No Version Targeted" then
					        response.Write "</TD><TD valign=top  class=""cell"">"
					    else
					        response.Write "</TD><TD nowrap bgcolor=""#ffff99"" valign=top  class=""cell"">"
					    end if
					    response.Write "<font size=1 face=verdana  class=""text"">" & strLocation
					    Response.Write  midrowcenter & "<font ID=""InIMGCell" & rs("ID") & "_" &  PVID & """ size=1 face=verdana  class=""text"">"
					    Response.write strInImageDisplay & midrow &"<font ID=""NoteCell" & rs("ID") & "_" &  PVID & """ size=1 face=verdana  class=""text"">"
                        if strFusionRequirements = 1 then 
                            if (rs("NoteExists")) = "TBD" then
                               Response.write "TBD"
                            elseif (rs("NoteExists") = "1" and Request("ProductRelease") = "") then
				                Response.Write "<a href='#' class=""check"" id=""aNotes" & rs("ID") & "_" &  PVID & """ onclick='return ChangeTargetNotes_Pulsar(" & PVID & "," & rs("ID") & "," & rs("RootID") & ")'>View/Edit</a>"
				                if rs("TargetNotes") <> "" then
                                    Response.Write "<div style=""text-align:right;float:right;""><a name=""schedule_tooltip"" href=""#"" class=""tt""><img src=""images/info.png"" alt=""Info"" /><span class=""tooltip""><span class=""top""></span><span class=""middle"">" & replace(server.HTMLEncode(rs("TargetNotes")&""),";","<br />") & "</span><span class=""bottom""></span></span></a></div>"
                                end if 
                            else
				                Response.Write "<a href='#' class=""check"" id=""aNotes" & rs("ID") & "_" &  PVID & """ style='display:none' onclick='return ChangeTargetNotes_Pulsar(" & PVID & "," & rs("ID") & "," & rs("RootID") & ")'>View/Edit</a>&nbsp;"
                            end if
                        else    
					        Response.write server.htmlencode(rs("TargetNotes") & "") & "&nbsp;"
                        end if
					    Response.write midrow
                        if strDistribution = "TBD" or strDistribution = "&nbsp;" then
	    				    Response.write "<font size=1 ID=""DistCell" & rs("ID") & "_" &  PVID & """ face=verdana  class=""text"">" & strDistribution
                        else
	    				    Response.write "<a href='#' class=""check"" id=""DistCell" & rs("ID") & "_" &  PVID & """ onclick='return ChangeDistribution(" & PVID & "," & rs("ID") & "," & rs("RootID") & ")'>"& strDistribution & "</a>&nbsp;"
                        end if
            
                        if server.htmlencode(strImages) = "TBD" then
	    				    Response.write midrow &"<font size=1 ID=""IMGCell" & rs("ID") & "_" &  PVID & """ face=verdana  class=""text"">" & server.htmlencode(strImages) & postrow
                        else
	                        Response.write midrow & "<font size=1 ID=""IMGCell" & rs("ID") & "_" &  PVID & """ face=verdana  class=""text""><a href='#' class=""check"" id=""aImages" & rs("ID") & "_" &  PVID & """ onclick='return ViewPulsarImages(" & PVID & "," & rs("ID") & "," & rs("RootID") & ")'>" & server.htmlencode(strImages) & "</a>&nbsp;"& postrow
                        end if
					    'Response.Flush
				    end if
			    end if'********
			    rs.MoveNext
		    loop

    		if (rs2.State <> adStateClosed) then
                    rs2.close
            end if

        	if (rsOTS.State <> adStateClosed) then
                    rsOTS.close
            end if

		    if i+j = 0 then
		    Response.Write "<TR><TD colspan=10>none</TD></TR>"
		    end if
		    Response.write "</table>"
		    Response.Write "<BR><BR><font size=1 face=verdana>Versions Displayed: " & i & "</font>"
	    end if

	    Response.Write "</TD></TR></TABLE></span>"

	    rs.Close
    end if
end if
'######################################
'	Action Tabs
'######################################
if strDisplayedList <> "Action" then
	Response.Write "<Table style=""Display:none"" ID=TableAction><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else

 strSQL = "spListActionItems " & PVID & ",2," & strStatusID
rs.Open strSQL,cn,adOpenForwardOnly

  If not(rs.EOF and rs.BOF) then  %>
<table ID="TableAction" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">

  <thead>
	<td onClick="SortTable( 'TableAction', 0,1,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="50" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Number</strong> </font> </td>
    <td onClick="SortTable( 'TableAction', 1,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="120" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Owner</strong></font></td>
    <td onClick="SortTable( 'TableAction', 2,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="80" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Status</strong></font></td>
    <td onClick="SortTable( 'TableAction', 3,2,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Target Date</strong></font></td>
    <td onClick="SortTable( 'TableAction', 4,0,2);" onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();" nowrap width="100%" bgColor="cornsilk"><strong><font size="1">Summary </font></strong></td>
  </thead>
  <%
  	ItemsDisplayed = 0

  do while not rs.EOF  %>
  <tr class="ID=<%=rs("ID")%>&amp;Type=<%=rs("Type")%>" id="actionrows" LANGUAGE="javascript" onMouseOver="return actionrows_onmouseover()" onMouseOut="return actionrows_onmouseout()" onClick="return actionrows_onclick()" oncontextmenu="javascript:contextMenu(<%=rs("ID")%>,<%=rs("Type")%>);return false;">
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

	Select case rs("status")
	case 1
		strStatus = "Open"
	case 2
		strStatus = "Closed"
	case 3
		strStatus = "Need Info"
	case 4
		strStatus = "Approved"
	case 5
		strStatus = "Disapproved"
	case 6
		strStatus = "Investigating"
	case else
		strStatus = "N/A"
	end select

		ItemsDisplayed = ItemsDisplayed + 1

        if trim(strTarget) = "" then
            strtarget = "&nbsp;"
        end if
  %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("ID") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Owner") & ""%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strStatus%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=strTarget%></font></td>
	<td valign="top" class="cell"><font size="1" class="text"><%=rs("summary") & "&nbsp;"%></font></td>
  </tr>

  <%	rs.MoveNext
	loop
	%>
</table>
<%else%>
<table ID="TableAction" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No open action items found for this program.</font>
</td></tr></table>

<%end if

rs.Close

end if

'##############################################################################
'#
'#  Agency Section
'#
'##############################################################################

if strDisplayedList <> "Agency" then
	Response.Write "<Table style=""Display:none"" ID=TableAgency><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
        Response.Write "<div id=""GridViewContainer"" class=""GridViewContainer"" style=""width: 100%; height: 500px;"">"
		Response.Write "<Table ID=TableAgency width=100% border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>"
		Call DrawPMViewMatrix(PVID, blnAgencyDataMaintainer)
		Response.Write "</table><p>* Country added after POR by DCR</p>"
        Response.Write "</div>"
end if
'<!-- End Agency Section -->

'##############################################################################
'#
'#  Schedule Section
'#
'##############################################################################
if strDisplayedList <> "Schedule" then
	Response.Write "<Table style=""Display:none"" ID=TableSchedule style=""margin:0px; padding:0px;""><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
 Dim strHistory
 strSQL = "usp_SelectScheduleData NULL," & clng(m_ScheduleID) & ",NULL,NULL,'Y'"
 rs.Open strSQL,cn,adOpenForwardOnly

  If not(rs.EOF and rs.BOF) then%>
  <%
  dim strLastPhase
  strLastPhase = ""
  %>
<!--<Table ID=TableSchedule style="Display:none"><tr><td>-->
<table ID="TableSchedule" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <tr>
	<td nowrap width="220" bgColor="cornsilk" vAlign="middle" rowspan="2"><font size="1"><strong>Schedule&nbsp;Item</strong> </font> </td>
    <td nowrap width="140" bgColor="cornsilk" Align="center" vAlign="middle" colspan="2"><font size="1"><strong>POR</strong><br><font color="red">For Reference Only</font></font></td>
    <td nowrap width="140" bgColor="cornsilk" Align="center" vAlign="middle" colspan="2"><font size="1"><strong>Current<br>Commitment</strong></font></td>
    <td nowrap width="140" bgColor="cornsilk" Align="center" vAlign="middle" colspan="2"><font size="1"><strong>Actual</strong></font></td>
    <td nowrap width="50" bgColor="cornsilk" rowspan="2"><strong><font size="1">Owner</font></strong></td>
    <td nowrap width="100%" bgColor="cornsilk" rowspan="2"><strong><font size="1">Comments</font></strong></td>
  </tr>
  <tr>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Start</strong></font></td>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Finish</strong></font></td>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Start</strong></font></td>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Finish</strong></font></td>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Start</strong></font></td>
    <td nowrap width="70" bgColor="cornsilk" Align="center" vAlign="middle"><font size="1"><strong>Finish</strong></font></td>
  </tr>
  <%do while not rs.EOF

	if strLastPhase <> rs("phase_name") then%>
		<tr bgcolor="lightsteelblue"><td nowrap valign="top" colspan="9"><font color="black" size="1" class="text"><%=rs("phase_name") & ""%></font></td>
		<%strLastPhase = rs("phase_name") & ""
	end if

	%>

	<%if blnAdministrator Or blnPlatformDevelopmentPM Or blnProductServiceManager Or blnSEPMProducts = 1 then%>
		<tr <% if rs("required_yn") & "" = "y" and rs("projected_end_dt") & "" = "" then response.write "bgcolor='mistyrose' " end if %> class="ID=<%=rs("schedule_data_id")%>" id="schedulerows" onMouseOver="return schedulerows_onmouseover()" onMouseOut="return schedulerows_onmouseout()" onClick="return schedulerows_onclick(<%=strFusionRequirements %>, <%=m_ScheduleID %>)">
	<%else%>
		<tr <% if rs("required_yn") & "" = "y" and rs("projected_end_dt") & "" = "" then response.write "bgcolor='mistyrose' " end if %>>
	<%end if %>
  <%

	if isnull(rs("actual_start_dt")) then
		strActualStart = "&nbsp;"
	else
		strActualStart = formatdatetime(rs("actual_start_dt"),2)
	end if

	if isnull(rs("actual_end_dt")) then
		strActualEnd = "&nbsp;"
	else
		strActualEnd = formatdatetime(rs("actual_end_dt"),2)
	end if

	if isnull(rs("por_start_dt")) then
		strPORStart = "&nbsp;"
	else
		strPORStart = formatdatetime(rs("por_start_dt"),2)
	end if

	if isnull(rs("por_end_dt")) then
		strPOREnd = "&nbsp;"
	else
		strPOREnd = formatdatetime(rs("por_end_dt"),2)
	end if

	'if not isnull(rs("actual_start_dt")) then
	'	strTargetStart = "---"
	'else
	if isnull(rs("projected_start_dt")) then
		strTargetStart = "&nbsp;"
	else
		strTargetStart = formatdatetime(rs("projected_start_dt"),2)
		If rs("actual_start_dt")&"" = "" And DateDiff("d", strTargetStart, Now()) > 0 Then
		    strTargetStartStyle = "font-weight:bold;color:red;"
		Else
		    strTargetStartStyle = ""
		End If
	end if

	'if not isnull(rs("actual_end_dt")) then
	'	strTargetEnd = "---"
	'else
	if isnull(rs("projected_end_dt")) then
		strTargetEnd = "&nbsp;"
	else
		strTargetEnd = formatdatetime(rs("projected_end_dt"),2)
        If rs("actual_end_dt")&"" = "" And DateDiff("d", strTargetEnd, Now()) > 0 Then
		    strTargetEndStyle = "font-weight:bold;color:red;"
		Else
		    strTargetEndStyle = ""
		End If

	end if

	Dim bIsMilestone
	If UCase(rs("milestone_yn")) = "Y" Then
		bIsMilestone = True
	Else
		bIsMilestone = False
	End If

  %>

	<td valign="top" class="cell"><div style="font-size:xx-small; float:left" class="text"><%=rs("item_description") & ""%></div><%if rs("item_definition") & "" <> "" then %><div style="text-align:right;float:right;"><a name="schedule_tooltip" href="#" class="tt"><img src="images/info.png" alt="Info" /><span class="tooltip"><span class="top"></span><span class="middle"><%=replace(server.HTMLEncode(rs("item_definition")&""),vbcrlf,"<br />") %></span><span class="bottom"></span></span></a></div><% end if %></td>
	<td valign="top" class="cell" align="center" <% if rs("required_yn") & "" = "y" and rs("projected_end_dt") & "" = "" then response.write "bgcolor='mistyrose' " else response.write "bgcolor='cornsilk' " end if %> <% If bIsMilestone Then Response.Write "ColSpan=2" %>><span style="font-size:xx-small" class="text"><%=strPORStart & ""%>&nbsp;</span></td>
<%	If NOT bIsMilestone Then %>
	<td valign="top" class="cell" align="center"><span style="font-size:xx-small" class="text"><%=strPOREnd & ""%>&nbsp;</span></td>
<%	End If %>
	<td valign="top" class="cell" align="center" <% if rs("required_yn") & "" = "y" and rs("projected_end_dt") & "" = "" then response.write "bgcolor='mistyrose' " else response.write "bgcolor='cornsilk' " end if %> <% If bIsMilestone Then Response.Write "ColSpan=2" %>><span style="font-size:xx-small;<%= strTargetStartStyle %>" class="text" id="ScheduleTargetStart<%=trim(rs("schedule_data_id"))%>"><%=strTargetStart%>&nbsp;</span></td>
<%	If NOT bIsMilestone Then %>
	<td valign="top" class="cell" align="center"><span style="font-size:xx-small;<%= strTargetEndStyle %>" class="text" id="ScheduleTargetEnd<%=trim(rs("schedule_data_id"))%>"><%=strTargetEnd%>&nbsp;</span></td>
<%	End If %>
	<td valign="top" class="cell" align="center" <% if rs("required_yn") & "" = "y" and rs("projected_end_dt") & "" = "" then response.write "bgcolor='mistyrose' " else response.write "bgcolor='cornsilk' " end if %> <% If bIsMilestone Then Response.Write "ColSpan=2" %>><span style="font-size:xx-small" class="text" id="ScheduleActualStart<%=trim(rs("schedule_data_id"))%>"><%=strActualStart%>&nbsp;</span></td>
<%	If NOT bIsMilestone Then %>
	<td valign="top" class="cell" align="center"><span style="font-size:xx-small" class="text" id="ScheduleActualEnd<%=trim(rs("schedule_data_id"))%>"><%=strActualEnd%>&nbsp;</span></td>
<%	End If %>
	<td valign="top" class="cell"><span style="font-size:xx-small" class="text" id="ScheduleOwner<%=trim(rs("schedule_data_id"))%>"><%=rs("RoleShortName")%>&nbsp;</span></td>
	<td valign="top" class="cell"><span style="font-size:xx-small" class="text" id="ScheduleComments<%=trim(rs("schedule_data_id"))%>"><%=rs("item_notes")%>&nbsp;</span></td>
  </tr>

  <%	rs.MoveNext
	loop
	%>
</table>
<%
for i = 0 to 15
    response.Write "<br />"
next
%>
<!--</td></tr></table>-->
<%else%>
<table ID="TableSchedule" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No Schedule Items found for this program.</font>
</td></tr></table>

<%end if
rs.Close
end if
'######################################
'	Requirements Tabs
'######################################
if strDisplayedList <> "Requirements" then
	Response.Write "<Table style=""Display:none"" ID=TableRequirements><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
 strSQL = "spListRequirementsByProduct " & PVID
rs.Open strSQL,cn,adOpenForwardOnly
 
    if strFusionRequirements = 1 then   
        %>  
        <script>
            $(function(){
                window.parent.location.href = "/Pulsar/Product/ProductRequirementList?ProductVersionID=<%=PVID%>&ReleaseID=0&ProductNameTitle="+document.getElementById("productNameTitle").innerText;
            });
        </script> 
    <% end if 
    %> 
<table ID="TableRequirements" style="Display:none" cellSpacing="0" cellPadding="0" width="100%"><tr><td>
  <%
    dim strPRLList
    dim PRLCount
    dim strSingleSCRSource
    strPRLList = ""    
    PRLCount = 0
   ' if request("Display") = "PRL" then       
    'for pulsar product, the requirement tab will combine the previous prl and new grid into one page so convert the logic request("Display" to use strFusionRequirements
   if strFusionRequirements = 1 then   

        'below is the prl list section, moved this to MVC page, so removed the code
        
       'end of requirements tab contents for pulsar products
    else
       'start the requirements tab content for legacy products
    If not(rs.EOF and rs.BOF) then
	    strLastcategory=""
	    do while not rs.EOF

            if trim(strLastCategory) <> trim(rs("Category") &"") then
                if strLastCategory <> "" then
                    response.write "</table><BR>"
                end if
                %>
                    <font size=2 face=verdana><b><%=rs("category")%></b></font><br>
                    <table cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
                  <tr>
                    <%if rs("RequiresDeliverables") then%>
                        <td width="50%" bgColor="wheat" vAlign="center"><font size="1"><strong>Product Requirement</strong></font></td>
                                <td width="50%" bgColor="wheat" vAlign="center"><font size="1"><strong>Related Components</strong></font></td>
                    <%else%>
                        <td colspan=2 width=100% bgColor="wheat" vAlign="center"><font size="1"><strong>Product Requirement</strong></font></td>
                    <%end if%>
                  </tr>

                <%
                strLastCategory = trim(rs("Category") &"")
            end if

		    if strDeliverables = "" then
			    strDeliverables = "&nbsp;"
		    end if
            %>
			    <tr bgcolor="cornsilk"><td colspan="2"><font size="1" face="verdana"><%=rs("requirement")%></font></td></tr>
			    <%if blnAdministrator or  blnMArketingAdmin or (strdevcenter="2" and (blnPreinstallPM or blnOdmPreinstallPM))then%>
				    <tr class="ID=<%=rs("ID")%>&amp;ProdID=<%=PVID%>" id="Req:<%=rs("ID")%>" LANGUAGE="javascript" onMouseOver="return requirementrows_onmouseover('Req:<%=rs("ID")%>')" onMouseOut="return requirementrows_onmouseout('Req:<%=rs("ID")%>')" onClick="return requirementrows_onclick('ID=<%=rs("ID")%>&amp;ProdID=<%=PVID%>')">
			    <%else%>
				    <tr>
			    <%end if%>
                <%if trim(rs("specification") & "") = "" or trim(rs("specification") & "") = "&nbsp;" then %>
				    <td valign="top" class="cell"><font size="1" class="text"><%="<Label ID=SpecCell" & rs("ID") & ">" & "See PDD for Requirements." & "&nbsp;</label>"%></font></td> <!--rs("specification")-->
                <%else%>
				    <td valign="top" class="cell"><font size="1" class="text"><%="<Label ID=SpecCell" & rs("ID") & ">" & trim(rs("specification") & "")%></label></font></td>
                <%end if%>
                <%if rs("RequiresDeliverables") then%>
				    <td valign="top" class="cell"><font size="1" class="text" ID="DellCell<%=rs("ID")%>">
				    <%
			         rs2.Open "splistdeliverablesbyrequirement " & clng(rs("ID")) & "," & PVID, cn, adOpenForwardOnly
				    Response.Write "<table width=100% border=1 cellspacing=0 cellpadding=2>"
				    Do While Not rs2.EOF
					    Response.Write "<TR><TD><font size=1>" & rs2("Name") &  "</font></TD></TR>"
					    rs2.MoveNext
				    Loop
				    rs2.Close
				    Response.Write  "</table>"


				     %></font></td>
                <%end if%>
			    </tr>

			    <%
		    rs.MoveNext
	    loop
	    %>
        </table></td></tr>
    <%else%>
        <font face="Verdana" size="2">No requirements found for this program.</font></td></tr>
    <%end if
end if %>
</table>
<%
rs.Close
end if


'**************END Requirements Section*************************

'**************BEGIN SMR Section*************************
if strDisplayedList <> "PMR" then
	Response.Write "<Table style=""Display:none"" ID=""TablePMR""><TR><TD></td></tr></table>"
else

dim smroldLink
smroldLink = Application("Release_Houston_ServerName") + "/SEPMRequestReport/"
%>
<div id="TablePMR" style="display:none;">
<h4>We are in the process of updating the SMR application. The information about Softpaqs can be found <a href="<%=smroldLink %>" target="_blank">HERE</a>.<h4/>
<h3 id="newSMRLoading">Loading SMR widget...</h3>
<iframe id="newSMRiFrame" style="width:100%;height:800px;border:0;" frameborder="0"></iframe>
</div>
<%
end if
'**************END SMR Section*************************

'**************BEGIN Localization Section*************************
'######################################
' Images Tabs	Local Tabs
'######################################
    dim strImageTypeFilter
	dim strSKUCount
	dim strImageCount
	dim ImageCountTotal
	dim SKUCountTotal
	dim ImageDefinitionCount
    dim strImageActiveTypeFilter
	strImageActiveTypeFilter = 1

if strDisplayedList <> "Local" then
	Response.Write "<Table style=""Display:none"" ID=TableLocal><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else%>

<%if strImageTool = "IRS" then %>
    <table ID="TableLocal" style="Display:none" width="100%">
    <tr><td>
    <%
    if trim(request("ImageType")) = "0" then
        strImageTypeFilter  = ""
    elseif trim(request("ImageType")) = "" then
        strImageTypeFilter  = ",1"
    elseif isnumeric(request("ImageType")) then
        strImageTypeFilter = "," & clng(request("ImageType"))
    else
        strImageTypeFilter  = ""
    end if

    if request("ImageActiveType") = "Active" or request("ImageActiveType") = "" then
			strImageActiveTypeFilter = 1
	elseif request("ImageActiveType") = "Inactive" then
			strImageActiveTypeFilter = 2 
	elseif request("ImageActiveType") = "NotReleased" then
			strImageActiveTypeFilter = 3 
	elseif request("ImageActiveType") = "InFactory" then
			strImageActiveTypeFilter = 4
	elseif request("ImageActiveType") = "All" then
			strImageActiveTypeFilter = 5
	End if
    
    dim releaseIDForImage
	releaseIDForImage = request("ProductReleaseID")
	if releaseIDForImage="" or IsNull(releaseIDForImage) then
		releaseIDForImage = 0
	else
		releaseIDForImage = cLng(releaseIDForImage)
	end if

	dim OSRIDForImage
	OSRIDForImage = request("ProductOSReleaseID")
	if OSRIDForImage="" or IsNull(OSRIDForImage) then 
		OSRIDForImage = 0
	else
		OSRIDForImage = cLng(OSRIDForImage)
	end if

    if strFusionRequirements = 0 then
        if request("ImageActiveType") = "Active"  or request("ImageActiveType") = "" then
	        rs.Open "spListImageDefinitionsFusion " & clng(strID) & ",1" & strImageTypeFilter & "," & OSRIDForImage,cn,adOpenForwardOnly
        else
            rs.Open "spListImageDefinitionsFusion " & clng(strID) & "," & strImageActiveTypeFilter & strImageTypeFilter & "," & OSRIDForImage,cn,adOpenForwardOnly 
        end if
    else
		rs.Open "usp_Image_GetImageDefinitionList " & clng(strID) & "," & strImageActiveTypeFilter & strImageTypeFilter & ",0," & releaseIDForImage & "," & OSRIDForImage ,cn,adOpenForwardOnly
    end if


    if rs.EOF and rs.BOF then
	    Response.Write "<font size=2 face=verdana>No IRS Images definitions match the selected criteria.</font>"
    else
        %>
            <table ID=TableImage cellpadding="2" border="1" bordercolor="tan" bgcolor="ivory">
             <thead><tr bgcolor="cornsilk">
	            <td  nowrap onClick="SortTable( 'TableImage', 0 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product Drop</font></b></td>
	            <td  onclick="SortTable( 'TableImage', 1 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Brands</font></b></td>
	            <td  onclick="SortTable( 'TableImage', 2 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">OS</font></b></td>
                <% if  strFusionRequirements = 1 then %>
					<td  onclick="SortTable( 'TableImage', 3 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Release</font></b></td>
                    <td  onclick="SortTable( 'TableImage', 4 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Status</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 5 ,1,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Images</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 6 ,2,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">EOL&nbsp;Date</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 7 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Comments</font></b></td>
                    <td  onclick="SortTable( 'TableImage', 8 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Releases&nbsp;for&nbsp;Operating&nbsp;System</font></b></td>
                <% else %>
                    <td  onclick="SortTable( 'TableImage', 3 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Status</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 4 ,1,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Images</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 5 ,2,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">EOL&nbsp;Date</font></b></td>
	                <td  onclick="SortTable( 'TableImage', 6 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Comments</font></b></td>
                    <td  onclick="SortTable( 'TableImage', 7 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Releases&nbsp;for&nbsp;Operating&nbsp;System</font></b></td>
   				<%end if%>
                    
            </tr>
             </thead>
             <%

	            ImageCountTotal = 0
	            ImageDefinitionCount=0
                IRSProductDropID = 0

	            do while not rs.EOF
                    if strFusionRequirements = 0 then
		                rs2.open "spCountImagesinDefinition " & rs("ID"),cn,adOpenForwardOnly
                    else
                        rs2.open "usp_Image_CountImagesinDefinition " & rs("ID"),cn,adOpenForwardOnly
                    end if
		            stsImageCount = "1"
		            if not(rs2.eof and rs2.bof) then
			            strImageCount=rs2("ImageCount") & ""
		            end if
		            rs2.close
		            if not isnumeric(strImageCount) then
			            strImageCount = "0"
		            end if
		            ImageCountTotal = ImageCountTotal + strImageCount

                    if trim(rs("eoldate") & "") = "" then
                        strImageEOL = "&nbsp;"
                    else
                        strImageEOL = formatdatetime(rs("eoldate"),vbshortdate)
                    end if

                    if trim(rs("ProductDrop") & "") <> "" then
                        rs2.open "spGetProductDropIDFusion '" & rs("ProductDrop") & "'",cn,adOpenForwardOnly
                        if not (rs2.EOF and rs2.BOF ) then
                            IRSProductDropID = rs2("ProductDropID") & ""
                        end if
                        rs2.close
                        if IRSProductDropID = "" then
                            IRSProductDropID = 0
                        end if
                    end if

                    strImageBrandSummary = ""
                     if strFusionRequirements = 0 then
		                rs2.open "spListImageDefinitionBrands " & rs("ID"),cn,adOpenForwardOnly
                     else
                        rs2.open "usp_Image_ListImageDefinitionBrands " & rs("ID"),cn,adOpenForwardOnly
                    end if
                 
                    do while not rs2.eof
                        strImageBrandSummary = strImageBrandSummary & ", " & rs2("Brand")
                        rs2.movenext
                    loop
		            rs2.close
                    if trim(strImageBrandSummary) <> "" then
                        strImageBrandSummary = mid(strImageBrandSummary,3)
                    end if

	            %>

		            <tr class="ID=<%=rs("ID")%>" id="Imagerows" LANGUAGE="javascript" onMouseOver="return imagerows_onmouseover()" onMouseOut="return imagerows_onmouseout()" oncontextmenu="localizationMenuFusion(<%=rs("ID")%>,<%=IRSProductDropID%>);return false;" onClick="return localizationMenuFusion(<%=rs("ID")%>,<%=IRSProductDropID%>)">
			            <td nowrap class="cell"><font size="1" face="verdana" class="text"><%=rs("ProductDrop") & "&nbsp;" %></font></td>
			            <td nowrap class="cell"><font size="1" face="verdana" class="text"><%=strImageBrandSummary & "&nbsp;"%></font></td>
			            <td class="cell"><font size="1" face="verdana" class="text"><%=rs("OS") & "&nbsp;"%></font></td>
                        <% if  strFusionRequirements = 1 then %>
							<td class="cell"><font size="1" face="verdana" class="text"><%=rs("ReleaseName") & "&nbsp;"%></font></td>
						<%end if%>
			            <td class="cell"><font size="1" face="verdana" class="text"><%=rs("Status") & "&nbsp;"%></font></td>
			            <td class="cell"><font size="1" face="verdana" class="text"><%=strImageCount%></font></td>
			            <td class="cell"><font size="1" face="verdana" class="text"><%=strImageEOL%></font></td>
			            <td class="cell"><font size="1" face="verdana" class="text"><%=rs("Comments") & "&nbsp;"%></font></td>
                        <td class="cell"><font size="1" face="verdana" class="text"><%=rs("OSReleaseName") & "&nbsp;"%></font></td>
		            </tr>
	            <%
	                ImageDefinitionCount = ImageDefinitionCount + 1
		            rs.Movenext
	            loop


             %>
            </table>
        <%
    end if
    rs.close
    %>
    </td></tr>
</table>
<%else%>
<table ID="TableLocal" style="Display:none" width="100%">
<tr><td>
<%
if trim(request("ImageType")) = "0" then
    strImageTypeFilter  = ""
elseif trim(request("ImageType")) = "" then
    strImageTypeFilter  = ",0,1"
elseif isnumeric(request("ImageType")) then
    strImageTypeFilter = ",0," & clng(request("ImageType"))
else
    strImageTypeFilter  = ""
end if

if request("ImageActiveType") = "Active"  or request("ImageActiveType") = "" then
	rs.Open "spListImageDefinitionsByProduct " & clng(strID) & ",1" & strImageTypeFilter,cn,adOpenForwardOnly
elseif request("ImageActiveType") = "Inactive" then
	rs.Open "spListImageDefinitionsByProduct " & clng(strID) & ",2" & strImageTypeFilter,cn,adOpenForwardOnly
elseif request("ImageActiveType") = "NotReleased"  then
	rs.Open "spListImageDefinitionsByProduct " & clng(strID) & ",3"  & strImageTypeFilter,cn,adOpenForwardOnly
elseif request("ImageActiveType") = "InFactory"  then
	rs.Open "spListImageDefinitionsByProduct " & clng(strID) & ",4" & strImageTypeFilter ,cn,adOpenForwardOnly
elseif request("ImageActiveType") = "All" then
	rs.Open "spListImageDefinitionsByProduct " & clng(strID) & ",0" & strImageTypeFilter,cn,adOpenForwardOnly
end if

if rs.EOF and rs.BOF then
	Response.Write "<font size=2 face=verdana>No Images definitions match the selected criteria.</font>"
else
if strID = 349 then
	Response.Write "<font size=1 face=verdana color=red>Titan and Altima are using shared image strategy. Please refer to Titan Excalibur for the up-to-date Software and Image deliverables targets.</font>"
end if

%>

<table ID=TableImage cellpadding="2" border="1" bordercolor="tan" bgcolor="ivory">
 <thead><tr bgcolor="cornsilk">
	<td  nowrap onClick="SortTable( 'TableImage', 0 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Image Number</font></b></td>
	<td  onclick="SortTable( 'TableImage', 1 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Brand</font></b></td>
	<td  onclick="SortTable( 'TableImage', 2 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Software</font></b></td>
	<td  onclick="SortTable( 'TableImage', 3 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">OS</font></b></td>
	<td  onclick="SortTable( 'TableImage', 4 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Type</font></b></td>
	<td  onclick="SortTable( 'TableImage', 5,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Status</font></b></td>
	<td  onclick="SortTable( 'TableImage', 6 ,1,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Dashes</font></b></td>
	<td  onclick="SortTable( 'TableImage', 7 ,1,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">OS&nbsp;Images</font></b></td>
	<td  onclick="SortTable( 'TableImage', 8 ,2,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">EOL&nbsp;Date</font></b></td>
	<td  onclick="SortTable( 'TableImage', 9 ,0,2);"><b><font size="1" face="verdana" onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Comments</font></b></td>
</tr>
 </thead>
<%
	if strID = 268 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>SMB images are for Thurman 1.1 only</b></td></TR>"
	elseif strID = 267 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>These images are shared with Thurman 1.1</b></td></TR>"
	end if
	ImageCountTotal = 0
	SKUCountTotal=0
	ImageDefinitionCount=0

	do while not rs.EOF
		rs2.open "spCountImagesinDefinition " & rs("ID"),cn,adOpenForwardOnly
		strSKUCount = "0"
		stsImageCount = "1"
		if not(rs2.eof and rs2.bof) then
			strSKUCount=rs2("SKUCount") & ""
			strImageCount=rs2("ImageCount") & ""
		end if
		rs2.close
		if not isnumeric(strSKUCount) then
			strSKUCount = "0"
		end if
		if not isnumeric(strImageCount) then
			strImageCount = "0"
		end if
		SKUCountTotal = SKUCountTotal + strSKUCount
		ImageCountTotal = ImageCountTotal + strImageCount

        if trim(rs("eoldate") & "") = "" then
            strImageEOL = "&nbsp;"
        else
            strImageEOL = formatdatetime(rs("eoldate"),vbshortdate)
        end if
	%>

		<tr class="ID=<%=rs("ID")%>" id="Imagerows" LANGUAGE="javascript" onMouseOver="return imagerows_onmouseover()" onMouseOut="return imagerows_onmouseout()" oncontextmenu="localizationMenu(<%=rs("ID")%>);return false;" onClick="return localizationMenu(<%=rs("ID")%>)">
			<td nowrap class="cell"><font size="1" face="verdana" class="text"><%=rs("SKUNumber") & "&nbsp;" %></font></td>
			<td nowrap class="cell"><font size="1" face="verdana" class="text"><%=rs("Brand") & "&nbsp;"%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=rs("SWType") & "&nbsp;"%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=rs("OS") & "&nbsp;"%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=rs("ImageType") & "&nbsp;"%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=rs("Status") & "&nbsp;"%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=strSKUCount%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=strImageCount%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=strImageEOL%></font></td>
			<td class="cell"><font size="1" face="verdana" class="text"><%=rs("Comments") & "&nbsp;"%></font></td>
		</tr>
	<%
	    ImageDefinitionCount = ImageDefinitionCount + 1
		rs.Movenext
	loop

	if strID = 268 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>Consumer images are shared with Ford 1.1</b></td></TR>"
		rs.Close
		rs.Open "spListImageDefinitionsByProduct " & 267,cn,adOpenForwardOnly

		do while not rs.EOF
		%>
			<tr class="ID=<%=rs("ID")%>" id="Imagerows" LANGUAGE="javascript" onMouseOver="return imagerows_onmouseover()" onMouseOut="return imagerows_onmouseout()" oncontextmenu="localizationMenu(<%=rs("ID")%>);return false;" onClick="return localizationMenu(<%=rs("ID")%>)">
				<td nowrap class="cell"><font size="1" face="verdana" class="text"><%=rs("SKUNumber") & "&nbsp;" %></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("Brand") & "&nbsp;"%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("SWType") & "&nbsp;"%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("OS") & "&nbsp;"%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("ImageType") & "&nbsp;"%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("Status") & "&nbsp;"%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=strSKUCount%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=strImageCount%></font></td>
				<td class="cell"><font size="1" face="verdana" class="text"><%=rs("Comments") & "&nbsp;"%></font></td>
			</tr>
		<%
    	    ImageDefinitionCount = ImageDefinitionCount + 1
			rs.Movenext
		loop

	end if


%>


<%end if%>
</table>
<%
    if ImageDefinitionCount <> 0 then
        response.Write "<BR><BR><font size=1 face=verdana>Image Definitions Displayed: " & ImageDefinitionCount & "</font>"
    end if
%>


</td></tr>
</table>
<%
rs.Close
    end if
end if

'######################################
'	Country Tabs
'######################################

if strDisplayedList <> "Country" then
	Response.Write "<Table style=""Display:none"" ID=TableCountry><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else
  If Trim(m_BrandID) <> "" Then
    if (strFusionRequirements = 1) then        
        Set dw = New DataWrapper
		Set cm = dw.CreateCommandSP(cn, "usp_SelectBrandLocalizationData_Pulsar")
		dw.CreateParameter cm, "@p_intProductBrandID", adInteger, adParamInput, 8, m_BrandID
        If request("ProductRelease") <> "" Then
            dw.CreateParameter cm, "@p_chrRelease", adVarchar, adParamInput, 30, Request("ProductRelease")
        Else
            dw.CreateParameter cm, "@p_chrRelease", adVarchar, adParamInput, 30, ""
        End If 
		Set rs = dw.ExecuteCommandReturnRS(cm)

    else
        strSQL = "usp_SelectBrandLocalizationData2 " & m_BrandID 
        rs.Open strSQL,cn,adOpenForwardOnly
    end if

  If not(rs.EOF and rs.BOF) then%>
  <%
    dim blnPowerCord
    dim blnDuckheadPowerCord
    dim blnDuckhead

    blnPowerCord = rs.Fields("PowerCord").Value
    blnDuckheadPowerCord = rs.Fields("DuckheadPowerCord").Value
    blnDuckhead = rs.Fields("Duckhead").Value

    dim iColumnCount : iColumnCount = 12
    if blnPowerCord = true then
       iColumnCount = iColumnCount + 1
    end if

    if blnDuckheadPowerCord = true then
       iColumnCount = iColumnCount + 1
    end if 
    
    if blnDuckhead = true then
       iColumnCount = iColumnCount + 1
    end if
    
    if strFusionRequirements = 1 then
       iColumnCount = iColumnCount + 1
    end if 
    dim strLastRegion
  strLastRegion = ""
  %>
<table ID="TableCountry" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <tr bgcolor="cornsilk">
	<td><b><font color="black" size="1" class="text">Country</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">HP Code</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Country Code</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Product Dash</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Languages</font></b></td>
    <% if (strFusionRequirements = 1) then %>
        <td style="white-space:nowrap"><b><font color="black" size="1" class="text">Release(s)</font></b></td>
    <% End if %>
	<td style="display:"><b><font color="black" size="1" class="text">MUI</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Keyboard</font></b></td>
    <% if (strFusionRequirements = 1) then %>
        <td style="white-space:nowrap"><b><font color="black" size="1" class="text">Keyboard Layout</font></b></td>
    <% End if %>
	<% if (blnPowerCord = True) then %>
        <td id="PowerCord" style="white-space:nowrap"><b><font color="black" size="1" class="text">Power Cord</font></b></td>
    <% End if %>
    <% if (blnDuckheadPowerCord = True) then %>
        <td style="white-space:nowrap"><b><font color="black" size="1" class="text">Duckhead Power Cord</font></b></td>
    <% End if %>
    <% if (blnDuckhead = True) then %>
        <td style="white-space:nowrap"><b><font color="black" size="1" class="text">Duckhead</font></b></td>
    <% End if %>
	<td style="display:"><b><font color="black" size="1" class="text">Restore Solution</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Image Docs</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Printed Docs</font></b></td>
	<td style="display:"><b><font color="black" size="1" class="text">Comments</font></b></td>
  </tr>
  <%do while not rs.EOF

	if strLastRegion <> rs("Region") then%>
        <tr bgcolor="MediumAquamarine">      
             <td colspan="<%=iColumnCount%>" nowrap valign="top"><b><font color="black" size="1" class="text"><%=mid(left(rs("Region"),len(rs("Region"))-1),2) & ""%></font></b></td>        
        </tr>
		<%strLastRegion = rs("Region") & ""
	end if

	if blnAdministrator or HasPemission then
	    if rs("Active") = 0 Then 'Inactive %>
	     	<tr class="ID=<%=rs("ProdBrandCountryID")%>" id="countryrows" bgcolor="Grey" LANGUAGE="javascript" onMouseOver="return schedulerows_onmouseover()" onMouseOut="return schedulerows_onmouseout()" onClick="return countryrows_onclick()">
	  <%elseif (NOT ISNULL(rs("CustomLocalization"))) OR (ISNULL(rs("ID"))) Then %>
			<tr class="ID=<%=rs("ProdBrandCountryID")%>" id="countryrows" bgcolor="LightPink" LANGUAGE="javascript" onMouseOver="return schedulerows_onmouseover()" onMouseOut="return schedulerows_onmouseout()" onClick="return countryrows_onclick()">
	  <%else%>
			<tr class="ID=<%=rs("ProdBrandCountryID")%>" id="countryrows" LANGUAGE="javascript" onMouseOver="return schedulerows_onmouseover()" onMouseOut="return schedulerows_onmouseout()" onClick="return countryrows_onclick()">
	  <%End If
	else
	    if rs("Active") = 0 Then 'Inactive %>
	        <tr bgcolor="Grey">
	  <%elseif ISNULL(rs("CustomLocalization")) Then%>
			<tr>
	  <%else%>
			<tr bgcolor="LightPink">
	  <%End If
	end if %>

	<td valign="top" class="cell"><font size="1" class="text"><%=rs("Country") & ""%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("OptionConfig") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("CountryCode") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("DASH") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text">
	<%
		If ISNULL(rs("OtherLanguage")) Then
			Response.Write rs("OSLanguage") & "&nbsp;"
		Else
			Response.Write rs("OSLanguage") & "," & rs("OtherLanguage")
		End If
	%></font></td>
    <% if (strFusionRequirements = 1) then %>
        <td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("Releases") & "&nbsp;"%></font></td>
    <%end if %>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("MUI") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("Keyboard") & "&nbsp;"%></font></td>
    <% if (strFusionRequirements = 1) then %>
        <td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("KeyboardLayout") & "&nbsp;"%></font></td>
    <%end if %>
    <% if (blnPowerCord = True) then %>
        <td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("PowerCordGEO") & "&nbsp;"%></font></td>
    <% End if %>
    <% if (blnDuckheadPowerCord = True) then %>
        <td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("DuckheadPowerCordGEO") & "&nbsp;"%></font></td>
    <% End if %>
    <% if (blnDuckhead = True) then %>
       	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("DuckheadGEO") & "&nbsp;"%></font></td>
    <% End if %>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("RestoreMedia") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("DocKits") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("PrintedDocs") & "&nbsp;"%></font></td>
	<td style="display:" valign="top" class="cell"><font size="1" ID="LocalizationID=<%=trim(rs("ProdBrandCountryID"))%>" class="text"><%=rs("Comments") & "&nbsp;"%></font></td>
  </tr>

  <%	PreviousID=rs("ProdBrandCountryID")
		rs.MoveNext
	loop
	%>
</table>
<%else%>
<table ID="TableCountry" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No Localizations found for this brand.</font>
</td></tr></table>

<%end if
    rs.Close
    set cm = nothing
	set dw = nothing    
  Else
%>
<table ID="TableCountry" style="Display:none" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
<font face="Verdana" size="2">No Brands defined for this Product.</font>
</td></tr></table>
<%
  End If
end if


'Start Tool - New Item Section
if strProdType = "2"  then

	strSQlFilter = ""
	if trim(strListFilter) = "My" then 'request("ListFilter") = "My" or request("ListFilter") = "" then
		strSQlFilter = CurrentUserID
	else
		strSQLFilter = "0"
	end if
	strSQL  = ""
	if sList="Tool_Alerts" then
		strSQL = "spListActionAlerts " & PVID & ", " & strSQlFilter
	elseif sList="Tool_Issues" then
		strSQL = "spListActionItemsForTools " & PVID  & ",1, " & strSQlFilter
	elseif sList="Tool_Tasks" then
		strSQL = "spListActionItemsForTools " & PVID  & ",2, " & strSQlFilter
	elseif sList="Tool_Roadmap" then
		strSQL = "spListActionRoadmapSummary " & PVID
	elseif 	sList="Tool_Working" or sList="" then
		strSQL = "spListActionItemsForTools " & PVID  & ",0, " & strSQLFilter
	end if

	ItemsDisplayed=0

	if strSQL <> "" then
		rs.Open strSQL,cn,adOpenForwardOnly

		if rs.EOF and rs.BOF then
			Response.Write "<font size=1 face=verdana><b><BR>none</b></font>"
		else
		%>
			<table ID="ToolTable" cellSpacing="1" cellPadding="1" width="100%" bordercolor="tan" border="1" bgColor="ivory">
				<thead>
				<tr bgcolor="Cornsilk">
					<%if sList="Tool_Working" or sList="" then%>
						<td onClick="SortTable( 'ToolTable', 0 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">ID</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 1 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Type</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 2 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Status</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 3 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">PR</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 4 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Order</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 5 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Target</strong></font></td>
						<%if trim(strListFilter) <> "My" then%>
							<td onClick="SortTable( 'ToolTable', 6 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Owner</strong></font></td>
							<%NextColumn=7%>
						<%else
							NextColumn=6%>
						<%end if%>
						<td width="20%" onClick="SortTable( 'ToolTable', <%=NextColumn%> ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Roadmap</strong></font></td>
						<td width="100%" onClick="SortTable( 'ToolTable', <%=NextColumn+1%> ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Summary</strong></font></td>
					<%elseif sList="Tool_Roadmap" then%>
						<td class="cell" onClick="SortTable( 'ToolTable', 0 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">ID</strong></font></td>
						<td class="cell" onClick="SortTable( 'ToolTable', 1 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Status</strong></font></td>
						<td class="cell" onClick="SortTable( 'ToolTable', 2 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">PR</strong></font></td>
						<td class="cell" onClick="SortTable( 'ToolTable', 3 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Tasks</strong></font></td>
						<td class="cell" onClick="SortTable( 'ToolTable', 4 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Progress</strong></font></td>
						<td class="cell" onClick="SortTable( 'ToolTable', 5 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Target</strong></font></td>
						<td width="20%" onClick="SortTable( 'ToolTable', 6 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Notes</strong></font></td>
						<td width="80%" class="cell" onClick="SortTable( 'ToolTable', 7 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Summary</strong></font></td>
					<%else%>
						<td onClick="SortTable( 'ToolTable', 0 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">ID</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 1 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Status</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 2 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">PR</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 3 ,1,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Order</strong></font></td>
						<td onClick="SortTable( 'ToolTable', 4 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Target</strong></font></td>
						<%if trim(strListFilter) = "My" then%>
							<%NextColumn=5%>
						<%else%>
							<td onClick="SortTable( 'ToolTable', 5 ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Owner</strong></font></td>
							<%NextColumn=6%>
						<%end if%>
						<td width="20%" onClick="SortTable( 'ToolTable', <%=NextColumn%> ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Roadmap</strong></font></td>
						<td width="100%" onClick="SortTable( 'ToolTable', <%=NextColumn+1%> ,0,2);"><font size="1"><strong onMouseOut="javascript: HeaderMouseOut();" onMouseOver="javascript: HeaderMouseOver();">Summary</strong></font></td>
					<%end if%>
				</tr>
				</thead>
			<%
			do while not rs.EOF
				ItemsDisplayed=ItemsDisplayed+1
				%>

					<%


					if sList="Tool_Alerts" then
						if rs("Type") = 1 then
							strTypeName = "Issue"
						elseif rs("Type") = 2 then
							strTypeName = "Action"
						end if
						%>
						<tr LANGUAGE="javascript" onClick="return toolactionrows_onclick(<%=rs("ID")%>)" onMouseOut="return schedulerows_onmouseout()" onMouseOver="return schedulerows_onmouseover()">
						<td width="1"><input style="WIDTH:16;HEIGHT:16" type="checkbox" id="chkAllItems" name="chkAllItems"></td>
						<%
						Response.Write "<TD  class=cell nowrap><FONT class=text size=1>Assign Priority&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD class=cell><FONT class=text size=1>" & rs("Summary") & "</FONT></TD>"
					elseif sList="Tool_Roadmap" then
						if trim(rs("Timeframe") & "") = "" then
							strTargetDate = "TBD"
						elseif isdate(rs("Timeframe")) then
							if year(cdate(rs("Timeframe"))) > 1600 then
								strTargetDate = formatdatetime(rs("TimeFrame"),vbshortdate) & "&nbsp;"
							else
								strTargetDate = rs("TimeFrame") & "&nbsp;"
							end if
						else
							strTargetDate = rs("TimeFrame") & "&nbsp;"
						end if

						if rs("Tasks") = 0 then
							strPercentComplete = "0%"
						else
							set rs2 = server.CreateObject("ADODB.Recordset")
							rs2.open "Select count(ID) as ClosedActions from DeliverableIssues with (NOLOCK) where ActionRoadmapID = " & rs("ID") & " and status = 2",cn,adOpenForwardOnly
							if rs2.eof and rs2.bof then
								strClosedActions = 0
							strPercentComplete = "0%"
							else
								strClosedActions = rs2("ClosedActions")
								strPercentComplete = (int((strClosedActions / rs("Tasks")) * 100)) &  "%"
							end if
							set rs2=nothing
						end if

						if isnull(rs("Notes")) then
							strNotes = "&nbsp;"
						elseif len(rs("Notes") & "") > 21 then
							strNotes = left(rs("Notes") & "",21) & "..."
						else
							strNotes = rs("Notes") & "&nbsp;"
						end if

						if blnActionOwner or blnToolsPM then
						%>
							<tr LANGUAGE="javascript" onClick="return roadmaprows_onclick(<%=rs("ID")%>)" onMouseOut="return schedulerows_onmouseout()" onMouseOver="return schedulerows_onmouseover()">
						<%else%>
							<TR>
						<%end if%>
						<td><font class="text" size="1"><%=rs("ID")%>&nbsp;&nbsp;</font></td>
						<td><font class="text" size="1"><%=rs("Status")%>&nbsp;&nbsp;</font></td>
						<td><font class="text" size="1"><%=rs("DisplayOrder")%>&nbsp;&nbsp;</font></td>
						<td><font class="text" size="1"><%=rs("Tasks")%>&nbsp;&nbsp;</font></td>
						<td><font class="text" size="1"><%=strPercentComplete%>&nbsp;&nbsp;</font></td>
						<td nowrap><font class="text" size="1"><%=strTargetDate%>&nbsp;&nbsp;&nbsp;</font></td>
						<td class="cell"><font class="text" size="1"><%=strNotes%>&nbsp;</font></td>
						<td class="cell"><font class="text" size="1"><%=rs("Summary")%></font></td>
					<%else%>
						<%
						if rs("Type") = 1 then
							strTypeName = "Issue"
						elseif rs("Type") = 2 then
							strTypeName = "Action"
						end if

						if isnull(rs("Roadmap")) then
							strRoadmap = "TBD"
						elseif len(rs("Roadmap") & "") > 21 then
							strRoadmap = left(rs("Roadmap") & "",21) & "..."
						else
							strRoadmap = rs("Roadmap") & ""
						end if

						if isdate(rs("targetDate")) then
							strTargetDate = cdate(rs("targetDate"))
						else
							if trim(rs("Timeframe") & "") = "" then
								strTargetDate = "TBD"
							elseif isdate(rs("Timeframe")) then
								if year(cdate(rs("Timeframe"))) > 1600 then
									strTargetDate = formatdatetime(rs("TimeFrame"),vbshortdate) & "&nbsp;"
								else
									strTargetDate = rs("TimeFrame") & "&nbsp;"
								end if
							else
								strTargetDate = rs("TimeFrame") & "&nbsp;"
							end if

						end if
						if blnActionOwner or blnToolsPM then
						%>
						<tr LANGUAGE="javascript" onClick="return toolactionrows_onclick(<%=rs("ID")%>,<%=rs("Type")%>)" onMouseOut="return schedulerows_onmouseout()" onMouseOver="return schedulerows_onmouseover()">
						<%else%>
						<tr LANGUAGE="javascript" onClick="return toolactionrowDetails_onclick(<%=rs("ID")%>)" onMouseOut="return schedulerows_onmouseout()" onMouseOver="return schedulerows_onmouseover()">
						<%end if%>
						<td class="cell" nowrap><font class="text" size="1"><%=rs("ID")%>&nbsp;&nbsp;</font></td>
						<%if sList="Tool_Working" or sList=""  then%>
						<td class="cell" nowrap><font class="text" size="1"><%=strTypeName%>&nbsp;&nbsp;</font></td>
						<%end if%>
						<td class="cell" nowrap><font class="text" size="1"><%=rs("Status")%>&nbsp;&nbsp;</font></td>
						<td class="cell" nowrap><font class="text" size="1"><%=rs("Priority")%>&nbsp;&nbsp;</font></td>
						<td class="cell" nowrap><font class="text" size="1"><%=rs("DisplayOrder")%>&nbsp;&nbsp;</font></td>
						<td class="cell" nowrap><font class="text" size="1"><%=strtargetDate %>&nbsp;&nbsp;</font></td>
						<%if trim(strListFilter) <> "My" then%>
							<td class="cell" nowrap><font class="text" size="1"><%=rs("Owner")%>&nbsp;&nbsp;</font></td>
						<%end if%>
						<td nowrap class="cell"><font class="text" size="1"><%=strRoadmap%></font></td>
						<td class="cell"><font class="text" size="1"><%=rs("Summary")%></font></td>
					<%end if%>
				</tr>
				<%

			rs.MoveNext
			loop
		end if
		rs.Close

		%>
		</table>
	<%else%>
		<table ID="ToolTable" cellSpacing="1" cellPadding="1" width="100%" border="0"><tr><td>
		<font face="Verdana" size="2">None</font>
		</td></tr></table>
	<%end if

end if




'######################################
'	Documents Tabs
'######################################

if strDisplayedList <> "Documents" then
	Response.Write "<Table style=""Display:none"" ID=TableDocuments><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
else%>


<table width="100%" ID="TableDocuments" style="Display:none"><tr>

<%

 strSQL = "spGetProductVersion " & PVID
 rs.Open strSQL,cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		Response.Write "<td colspan=4 width=""100%""nowrap valign=top>"


		Response.Write "<BR><font size=2 face=verdana><b>General Documents</b></font><HR><BR></td></tr><tr><td nowrap>"

		dim strProductFileName
		strProductFileName = replace(rs("Name") + " " + rs("Version")," ", " ")
		strProductFileName = replace(strProductFileName,"/", "")
		%>
		<a target="_blank" href="SystemTeam.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a>
		<a target="_blank" href="SystemTeam.asp?ID=<%=rs("ID")%>">System Team Roster</a><br>
		<a target="_blank" href="ProductStatus.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a>
		<a target="_blank" href="ProductStatus.asp?ID=<%=rs("ID")%>">Current Status</a><br>
		<a target="_blank" href="image/DeliverableMatrix.asp?ProdID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" hspace="3" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="image/DeliverableMatrix.asp?ProdID=<%=rs("ID")%>">Deliverable Matrix</a><br>
		<a target="_blank" href="Deliverable/HardwareMatrix.asp?lstProducts=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" hspace="3" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="Deliverable/HardwareMatrix.asp?lstProducts=<%=rs("ID")%>">Hardware Qual Matrix</a><br>
		<a href="javascript:OpenStatusOptions(<%=rs("ID")%>);"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a href="javascript:OpenStatusOptions(<%=rs("ID")%>);">Changes This Week</a><br>
        </td><td nowrap valign=top>
        <%if strImageTool = "IRS" then%>
            <%if strFusionRequirements = 0 then %>
    		    <a target="_blank" href="Image/Fusion/Localization.asp?ProdID=<%=rs("ID")%>&amp;PINTest=0"><img SRC="images/ICON-DOC-HTML.GIF" border="0" hspace="3" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="Image/Fusion/Localization.asp?ProdID=<%=rs("ID")%>&amp;PINTest=0">Image Localization Matrix</a><br>
            <%else %>
                <a target="_blank" href="Image/Fusion/Localization_Pulsar.asp?ProdID=<%=rs("ID")%>&amp;PINTest=0"><img SRC="images/ICON-DOC-HTML.GIF" border="0" hspace="3" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="Image/Fusion/Localization.asp?ProdID=<%=rs("ID")%>&amp;PINTest=0">Image Localization Matrix</a><br>
            <%end if %>
	    <%else%>
        	<a target="_blank" href="image/localization.asp?ProdID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" border="0" hspace="3" WIDTH="16" HEIGHT="16"></a> <a target="_blank" href="image/localization.asp?ProdID=<%=rs("ID")%>">Localization Matrix</a><br>
		<%end if%>
        <a target="_blank" href="Deliverable/ProductDeliverableStatus.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Deliverable/ProductDeliverableStatus.asp?ID=<%=rs("ID")%>">Product Deliverable Status</a>&nbsp;&nbsp;&nbsp;<br>
        <%if strImageTool = "IRS" then%>
            <%if strFusionRequirements = 0 then %>
		        <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>">Rollout Plan</a><br>
		        <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>&Report=1"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>&Report=1">Ramp Plan</a><br>
            <%else %>
                <a target="_blank" href="Image/fusion/Buildplan_Pulsar.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>">Rollout Plan</a><br>
		        <a target="_blank" href="Image/fusion/Buildplan_Pulsar.asp?ID=<%=rs("ID")%>&Report=1"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/fusion/Buildplan.asp?ID=<%=rs("ID")%>&Report=1">Ramp Plan</a><br>
            <%end if %>
	    <%else%>
		    <a target="_blank" href="Image/Buildplan.asp?ID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/Buildplan.asp?ID=<%=rs("ID")%>">Rollout Plan</a><br>
		    <a target="_blank" href="Image/Buildplan.asp?ID=<%=rs("ID")%>&Report=1"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a target="_blank" href="Image/Buildplan.asp?ID=<%=rs("ID")%>&Report=1">Ramp Plan</a><br>
		<%end if%>
		<!--<a target="_blank" href="search/ots/default.asp?lstProduct=<%=strProductName%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" hspace="3" border="0"></a> <a target="_blank" href="search/ots/default.asp?lstProduct=<%=strProductName%>">Custom OTS Reports</a><br>-->
		</td><td nowrap valign=top>
		<a href="javascript: ShowSICommodities(<%=rs("ID")%>);"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" hspace="3" border="0"></a> <a href="javascript: ShowSICommodities(<%=rs("ID")%>);">Integration Commodity List</a><br>
		<a target="_blank" href="/iPulsar/ExcelExport/mdacompliance.aspx?pvid=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" hspace="3" border="0"></a> <a target="_blank" href="/iPulsar/ExcelExport/mdacompliance.aspx?pvid=<%=rs("ID")%>">MDA Compliance Report</a><br>
        <a target="_blank" href="PNPDevices.asp?Report=2&ProductID=<%=rs("ID")%>"><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16" hspace="3" border="0"></a> <a target="_blank" href="PNPDevices.asp?Report=2&ProductID=<%=rs("ID")%>">Device ID List</a><br>
		<%if trim(rs("Distribution")) <> "" then%>
			<a HREF="mailto:<%= rs("Distribution") %>"><img SRC="images/ICON-DOC-OUTLOOK.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"></a> <a HREF="mailto:<%= rs("Distribution") %>">Send Email to System Team</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
		<%else%>
			<img SRC="images/ICON-DOC-OUTLOOK.GIF" WIDTH="16" HEIGHT="16" border="0" hspace="3"> No Email List Defined<br>
		<%end if %>
		</td>

        <td width=100%>&nbsp;</td>
        </tr>
        <tr>
        <td valign=top nowrap colspan=4>
        <font size=2 face=verdana><b><BR>RTM Documents</b></font><HR></td></tr>
        <%
        rs.close

            Dim HasRTPPemission
             HasRTPPemission = 0

            set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
	        set rs = server.CreateObject("ADODB.recordset")
            cm.CommandType = 4
	        cm.CommandText = "usp_USR_ValidatePermission"
	
	        Set p = cm.CreateParameter("@p_intUserId", 200, &H0001, 15)
	        p.Value = CurrentUserID
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@p_PName", 200, &H0001, 100)
	        p.Value = "RTP.Edit"
	        cm.Parameters.Append p

            rs.CursorType = adOpenForwardOnly
	        rs.LockType=AdLockReadOnly
	        Set rs = cm.Execute 
	
            if not(rs.EOF and rs.BOF) then
		        HasRTPPemission = rs("HasPermission")
	        end if
            rs.Close

            response.write "<tr><td nowrap colspan=4>" _
                          & " <a href=""javascript: CreateRTM(" & pvid & ");"">Create RTM Document</a> &nbsp;&nbsp;" _
                          & " <a href=""javascript: CreateMRTM(" & pvid & ");"">Create RTM Documents for Multiple Products</a>" _
                          & "<BR></td></tr>"
                        rs.open "spListProductRTMDocuments " & clng(pvid) & ",0",cn
            if rs.eof and rs.bof then
                response.Write "<tr><td colspan=6><BR>No RTM documents have been created for this product.</td></tr>"
            else
                response.Write "<tr><td colspan=6 width=""100%""><BR><table id=TableRTMDocs bgcolor=ivory width=""100%"" border=1 bordercolor=tan>"
                response.Write "<thead><tr bgcolor=cornsilk>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 0,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>Title</font></b></td>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 1,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>RTM Date</font></b></td>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 2,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>RTP Date </font></b></td>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 3,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>RTP Status</font></b></td>"
               response.Write "<td onclick=""SortTable( 'TableRTMDocs', 4,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>RTP Actions</font></b></td></tr></thead>"
                 do while not rs.eof
                   typeId = rs("typeid") & ""
                   rtpStatus = rs("RTPStatus") & ""
                     response.Write "<tr><td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""   onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("Title") & "</font></td>"
                     response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("RTMDate") & "</font></td>"
                   If InStr(typeId, "1") > 0 or InStr(typeId, "4") or InStr(typeId, "0") > 0 Then
                      if Not IsNull(rtpStatus) and not rtpStatus = "" then
                        if cbool(rtpStatus) = true then
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("RTPDate") & " &nbsp;</font></td>"
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana> Approved on "  & rs("RTPCompletedby") & "</font></td>"
                          response.Write "<td class=""cell""  onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" ><A onclick=""UpdateRTPStatus(" & rs("ID") & ");"" href=""#"">Review RTP Status</A> </td></tr>"
                        else
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("RTPDate") & " &nbsp;</font></td>"
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>Rejected on "  & rs("RTPCompletedby") & "</font></td>"
                          response.Write "<td class=""cell""  onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  ><A onclick=""UpdateRTPStatus(" & rs("ID") & ");"" href=""#"">Review RTP Status</A> </td></tr>"  
                        end if
                      else
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  id=RTP_Date" & rs("ID") &" onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("RTPDate") & " &nbsp;</font></td>"
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  id=RTP_Status" & rs("ID") &" onclick=""javascript:RTMDocClick(" & rs("ID") & ");"">Pending</td>"
                        if HasRTPPemission  then
                          response.Write "<td class=""cell""  onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  ><A id=RTP_Action" & rs("ID") &" onclick=""UpdateRTPStatus(" & rs("ID") & ");"" href=""#"">Update RTP Status</A> </td></tr>"  
                        else
                          response.Write "<td class=""cell"" disabled  onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" ><A id=RTP_Action" & rs("ID") &" onclick=""UpdateRTPStatus(" & rs("ID") & ");"" href=""#"">Update RTP Status</A> </td></tr>"  
                        end if
                      end if
                    else
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>" & rs("RTPDate") & "&nbsp;</font></td>"
                          response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:RTMDocClick(" & rs("ID") & ");""><font size=1 face=verdana>N/A</font></td>"
                          response.Write "<td class=""cell""  onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><font size=1 face=verdana> &nbsp;</font></td></tr>"
                    end if

                    rs.movenext
                loop
                response.Write "</table></td><BR> </tr>"
            end if
            rs.close

                  response.Write "<tr> <td valign=top nowrap colspan=4><font size=2 face=verdana><b><BR>RTM Document (Draft)</b></font><HR><BR></td> </tr>"

            rs.open "spListProductRTMDocuments " & clng(pvid) & ",1",cn
            if rs.eof and rs.bof then
                response.Write "<tr><td colspan=6><BR>No RTM document (Draft) have been created for this product.</td></tr>"
            else
                response.Write "<tr><td colspan=6 width=""100%""><BR><table id=TableRTMDraftDocs bgcolor=ivory width=""100%"" border=1 bordercolor=tan>"
                response.Write "<thead><tr bgcolor=cornsilk>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 0,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>Title</font></b></td>"
                response.Write "<td onclick=""SortTable( 'TableRTMDocs', 1,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b><font size=1 face=verdana>RTM Date</font></b></td>"
                 do while not rs.eof
                   
                     response.Write "<tr><td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""   onclick=""javascript:UpdateRTMDraft(" & pvid & ","&rs("ID")&");""><font size=1 face=verdana>" & rs("Title") & "</font></td>"
                     response.Write "<td class=""cell"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""  onclick=""javascript:UpdateRTMDraft(" & pvid & ","&rs("ID")&");""><font size=1 face=verdana>" & rs("RTMDate") & "</font></td>"

                    rs.movenext
                loop
                response.Write "</table></td><BR> </tr>"
            end if
            rs.close
          
        %>

       
		<%
    else
	    rs.Close
	end if


%>

</tr>
</table>

<%
end if


		set rs = nothing
		set rs2 = nothing
		cn.close
		set cn= nothing

'----------------------------------
end if
end if
%>
<div ID="TableGeneral" style="Display:none">
<table cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
	<%'Harris, Valerie -  02/29/2016 - PBI 17178/ Task 17281 - Change General Tab System Team rows to match Product Properties' System Team Tab's Primary Team Memebers
        
			if strSMEmail <> "" then
				strSMName = "<a href=""mailto:" & strSMEmail & """>" & longname(strSMName) & "</a>"
			end if
			if strSEPMEmail <> "" then
				strSEPMName = "<a href=""mailto:" & strSEPMEmail & """>" & longname(strSEPMName) & "</a>"
			end if
			if strManagerEmail <> "" then
				strManager = "<a href=""mailto:" & strManagerEmail & """>" & longname(strManager) & "</a>"
			end if

        If bIsPulsarProduct = True Then 'PULSAR PRODUCT SYSTEM TEAM%>
        <tr>
	        <td rowspan="4" class="td-title-systemteam">
                <div class="left"><strong>System Team:</strong>&nbsp;&nbsp;&nbsp;</div><br /><br />
                <div class="center"><a target=_blank href="SystemTeam.asp?ID=<%= strID %>">Full Roster</a></div>
	        </td>
		    <td class="td-systemteam"><%="<b>System Manager</b><BR>" & strSMName%></td>
            <td class="td-systemteam"><%="<b>Platform Development PM</b><BR>" & strPlatformDevelopment%></td>
            <td class="td-systemteam"><%="<b>Supply Chain</b><BR>" & strSupplyChain%></td>
            <td class="td-systemteam"><%="<b>ODM System Engineering PM</b><BR>" & sODMSEPMName%></td>     
        </tr>
	    <tr>
			<td class="td-systemteam"><%="<b>Configuration Manager</b><BR>" & sConfigManagerName%></td>		
            <td class="td-systemteam"><%="<b>Commodity PM</b><BR>" & strCommodityPM%></td> 	    
		   	<td class="td-systemteam"><%="<b>Service</b><BR>" & strService%></td>	
            <td class="td-systemteam"><%="<b>ODM HW PM</b><BR>" & strODMHWPM%></td>
        </tr>
	    <tr>
		    <td class="td-systemteam"><%="<b>Program Office Program Manager</b><BR>" & sPOManagerName%></td>
            <td class="td-systemteam"><%="<b>Planning PM</b><BR>" & sPlanningPMName%></td>
            <td class="td-systemteam"><%="<b>Quality</b><BR>" & strQuality%></td>	
            <td class="td-systemteam"><% 
                if bIsPulsarProduct then
                    response.write "<b>BIOS Lead</b><BR>" & strBiosLead
                else 
                    response.write "&nbsp;"
                end if
                %>&nbsp;
            </td>
		</tr>
	    <tr>
            <td class="td-systemteam"><%="<b>Systems Engineering PM</b><BR>" & strSEPMName%></td>
            <td class="td-systemteam"><%="<b>Marketing/Product Mgmt</b><BR>" & strMarketing%></td>
		    <td class="td-systemteam"><%="<b>Procurement PM</b><BR>" & sProcurementPMName%></td>    
            <td class="td-systemteam">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	    </tr>
    <%  Else 'LEGACY PRODUCT SYSTEM TEAM%>
        <tr>
	        <td rowspan="5" class="td-title-systemteam">
                <div class="left"><strong>System Team:</strong></div><br /><br />
                <div class="center"><a target=_blank href="SystemTeam.asp?ID=<%= strID %>">Full Roster</a></div>
	        </td>
	        <td class="td-systemteam"><%="<b>System Manager</b><BR>" & strSMName%></td>
            <td class="td-systemteam"><%="<b>Platform Development PM</b><BR>" & strPlatformDevelopment%></td>
            <td class="td-systemteam"><%="<b>Commercial Marketing</b><BR>" & sComMarketingName%></td> 
            <td class="td-systemteam"><%="<b>Service</b><BR>" & strService%></td>
        </tr>
	    <tr>
            <td class="td-systemteam"><%="<b>Program Office Manager</b><BR>" & sPOManagerName%></td>
            <td class="td-systemteam"><%="<b>Commodity PM</b><BR>" & strCommodityPM%></td>
            <td class="td-systemteam"><%="<b>Consumer Marketing</b><BR>" & sConMarketingName%></td>
            <td class="td-systemteam"><%="<b>Quality</b><BR>" & strQuality%></td>
        </tr>
	    <tr>
			<td class="td-systemteam"><%="<b>Configuration Manager</b><BR>" & sConfigManagerName%></td>
            <td class="td-systemteam"><%="<b>ODM System Engineering PM</b><BR>" & sODMSEPMName%></td>
            <td class="td-systemteam"><%="<b>SMB Marketing</b><BR>" & sSMBMarketingName%></td>
            <td class="td-systemteam"><%="<b>Procurement PM</b><BR>" & sProcurementPMName%></td>
		</tr>
	    <tr>
            <td class="td-systemteam"><%="<b>Systems Engineering PM</b><BR>" & strSEPMName%></td>
            <td class="td-systemteam"><%="<b>Planning PM</b><BR>" & sPlanningPMName%></td>
            <td class="td-systemteam"><%="<b>Supply Chain</b><BR>" & strSupplyChain%></td>
            <td class="td-systemteam">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td> 
	    </tr>
    <%  End If %>
</table>
<table cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">

	<tr>
	    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">ODM:</font></strong></td>
		<td width="30%"><font size="1"><%=strPartnername%></font></td>
	    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Development&nbsp;Center:&nbsp;&nbsp;</font></strong></td>
		<td width="30%"><font size="1"><%=strDevCenterName%></font></td>
		</tr>
	<tr>
		<td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">System&nbsp;Board&nbsp;ID:&nbsp;&nbsp;</font></strong></td>
		<td width="30%"><font size="1"><%=FormatSystemID(strSystemboardComments)%></font></td>
		<td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Machine&nbsp;PnP&nbsp;ID:&nbsp;&nbsp;</font></strong></td>
    <td width="30%"><font size="1"><%=FormatSystemID(strMachinePNPComments)%></font></td>		</tr>
  <tr>
    <td bgColor="cornsilk" valign="top" width="100" nowrap><strong><font size="1">
		<%if strDevCenter = "2" then%>
			Reference&nbsp;Platform:&nbsp;&nbsp;
		<%else%>
			Lead&nbsp;Product:&nbsp;&nbsp;
		<%end if%>
		</font></strong></td>
      <%if strDevCenter = "2" then%>
			<td width="30%"><font size="1"><%=strReferencePlatform%>&nbsp;</font></td>
		<%else%>
			 <td width="30%"><font size="1"><%=strLeadProduct%>&nbsp;</font></td>
		<%end if%>     
    <td bgColor="cornsilk" valign="top" width="100" nowrap><strong><font size="1">Current&nbsp;BIOS&nbsp;Versions:&nbsp;&nbsp;</font></strong></td>
    <td width="30%"><font size="1"><%=strROMVersion%>&nbsp;</font></td></tr>
  <tr>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Preinstall&nbsp;Team:&nbsp;&nbsp;</font></strong></td>
    <td valign="top"><font size="1"><%=strPreinstallTeam%>&nbsp;</font></td>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Release&nbsp;Team:&nbsp;&nbsp;</font></strong></td>
    <td valign="top"><font size="1"><%=strReleaseTeam%>&nbsp;</font></td>

    </tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">Product&nbsp;Phase:</font></strong></td>
		<td><font size="1"><%=strStatus%></font></td>

		<td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Regulatory&nbsp;Model:&nbsp;&nbsp;</font></strong></td>
		<td valign="top"><font size="1"><%=strRegulatoryModel%>&nbsp;</font></td>
	</tr>
	<tr>
        <%if strFusionRequirements = 0 then%>
			<td bgColor="cornsilk"><strong><font size="1">Minimum&nbsp;RoHS&nbsp;Level:</font></strong></td>
		    <td><font size="1"><%=strMinRoHSLevel%>&nbsp;</font></td>		
		<%else%>
			<td bgColor="cornsilk"><strong><font size="1">Releases:</font></strong></td>
		    <td><font size="1"><%=strProductReleases%>&nbsp;</font></td>		
		<%end if%>
	    <td bgColor="cornsilk"><strong><font size="1">Factory:</font></strong></td>
		<td><font size="1"><%=strFactoryName%>&nbsp;</font></td>
	</tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">Product&nbsp;Line:</font></strong></td>
		<td><font size="1"><%=strProductLineName%>&nbsp;</font></td>
		<td bgColor="cornsilk"><strong><font size="1">Business&nbsp;Segment:</font></strong></td>
		<td><font size="1"><%=strBusinessSegmentName%>&nbsp;</font></td>
	</tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">End&nbsp;of&nbsp;Production:</font></strong></td>
		<td><font size="1"><%=strEndOfProductionDate%>&nbsp;</font></td>
	    <td bgColor="cornsilk"><strong><font size="1">End&nbsp;of&nbsp;Service:</font></strong></td>
		<td><font size="1"><%=strServiceLifeDate%>&nbsp;</font></td>
	</tr>
	<tr>
	    <td bgColor="cornsilk"><strong><font size="1">RCTO&nbsp;Sites:</font></strong></td>
		<td><font size="1"><%=strRCTOSites%>&nbsp;</font></td>
	    <td bgColor="cornsilk"><strong><font size="1">Product&nbsp;Groups:</font></strong></td>
		<td><font size="1"><%=strProductGroups%>&nbsp;</font></td>
	</tr>


		<%if strID = "268" then%>
			<tr>
				<td bgColor="cornsilk"><strong><font size="1">Note:</font></strong></td>
				<td colspan="4"><font size="1" color="red">Thurman 1.1 Excalibur is used to build Thurman SMB images only. For Thurman 1.1
Consumer image deliverables, please refer to Ford 1.1 Excalibur</font></td></tr>

		<%elseif strID = "349" then%>
			<tr>
				<td bgColor="cornsilk"><strong><font size="1">Note:</font></strong></td>
				<td colspan="4"><font size="1" color="red">Titan and Altima are using shared image strategy. Please refer to Titan Excalibur for the up-to-date Software and Image deliverables targets.</font></td></tr>
		<%elseif strID = "267" then%>
			<tr>
				<td bgColor="cornsilk"><strong><font size="1">Note:</font></strong></td>
				<td colspan="4"><font size="1" color="red">Ford/Thurman consumer share the same image. As such, this Excalibur page is showing all the Ford/Thurman 1.1 Consumer localizations and image deliverables are for the shared consumer image.</font></td></tr>
		<%end if%>
 <%if strFusionRequirements = 0 then %>
  <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">Marketing&nbsp;Names:&nbsp;&nbsp;</font></strong></td>
    <td valign="top" colspan="3"><%=strSeries1%></td>
  </tr>
 <% end if %>
 <%if strFusionRequirements = 1 and FollowMKTName =0 then %>
 <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">Marketing&nbsp;Names:&nbsp;&nbsp;</font></strong></td>
    <td valign="top" colspan="3"><%=strSeries1%></td>
  </tr>
  <% end if %>
  <%if strFusionRequirements = 0 then 
      'YONG made the following changes on 10/13/1016 for: Bug 28064:Production: Ticket 10460: Brands admin formulas do not update all Legacy product names
      'for what ever history reason, the similar sps: usp_GetBrands4product and spListBrandsforProducts are both called
      'the marketing names or phweb names are returned correctly in the usp_GetBrands4product; so replaced strPHWebFamily with strPHWebFamily1 for legacy product for example
      
      
      %>
  <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">PHWeb&nbsp;Names:&nbsp;&nbsp;</strong></font></td>
    <td valign="top" colspan="3"><font size="1"><%="<Table cellspacing=0 cellpadding=0><tr><td width=210 valign=top nowrap><b><font size=1 face=verdana>Family&nbsp;Name</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>Brand&nbsp;Name</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>KMAT</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>Last SCM Publish</font></b></td></tr><tr><td valign=top nowrap><font size=1 face=verdana>" & strPHWebFamily1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strBrandName1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strKMAT & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strLastScmPublish & "</font></td></tr></table>"%></td>
  </tr>
  <% end if %>

  <%if strFusionRequirements = 1 and FollowMKTName = 0 then %>
  <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">PHweb&nbsp;Names&nbsp;&nbsp;</strong></font></td>
    <td valign="top" colspan="3"><font size="1"><%="<Table cellspacing=0 cellpadding=0><tr><td width=210 valign=top nowrap><b><font size=1 face=verdana>Family&nbsp;Name</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>Brand&nbsp;Name</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>KMAT</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>Last SCM Publish</font></b></td></tr><tr><td valign=top nowrap><font size=1 face=verdana>" & strPHWebFamily1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strBrandName1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strKMAT1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strLastScmPublish1 & "</font></td></tr></table>"%></td>
  </tr>
  <% elseif strFusionRequirements = 1 and FollowMKTName = 1 then %>
     <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">KMAT/SCM</strong></font></td>
    <td valign="top" colspan="3"><font size="1"><%="<Table cellspacing=0 cellpadding=0><tr><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>KMAT</font></b></td><td nowrap width=10>&nbsp;</td><td valign=top nowrap><b><font size=1 face=verdana>Last SCM Publish</font></b></td></tr><tr><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strKMAT1 & "</font></td><td width=10 valign=top nowrap>&nbsp;</td><td valign=top nowrap><font size=1 face=verdana>" & strLastScmPublish1 & "</font></td></tr></table>"%></td>
  <% end if %>
   <%if strFusionRequirements = 1 then %>
  <tr>
    <td bgColor="cornsilk" valign="top" width="100"><strong><font size="1">Base&nbsp;Unit&nbsp;Groups:&nbsp;&nbsp;</strong></font></td>
    <td valign="top" colspan="3"><font size="1"><iframe id="pmview_PlatformFrame" frameBorder="0"marginheight="0px" marginwidth="0px" style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; border-left: steelblue 1px solid; border-bottom: steelblue 1px solid;margin-top:4px;height: 250px;width:100%;" src="/Excalibur/MobileSE/Today/PlatformList.asp?ID=<%=strID%>&Edit=0&FollowMktName=<%=FollowMktName%>&isCM=<%=isCMPermission%>"></iframe></td>
  </tr>  
   <% end if %>
  <tr>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Current&nbsp;Image<br>Part&nbsp;Number:&nbsp;&nbsp;</font></strong></td>
    <td valign="top" colspan="3"><font size="1"><%=strImagePO%></font></td></tr>
  <%if strFusionRequirements = 0 then %> 
  <tr>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Operating&nbsp;Systems:&nbsp;&nbsp;</font></strong></td>
    <td valign="top" colspan="3"><font size="1"><%=strOS%></font></td></tr>
  <%end if %>
  <tr>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Description:&nbsp;&nbsp;</font></strong></td>
    <td valign="top" colspan="3"><font size="1"><%=replace(strDescription,vbcrlf,"<BR>")%></font></td></tr>
  <tr>
    <td style="width:100px; background-color:cornsilk; vertical-align:top; ">
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">WHQL IDs:&nbsp;&nbsp;<br /><font face="Verdana" color="Red" size="1">Under Development</font></span></td>
    <td style="vertical-align:top;"><font size="1"><%= strWhqlEditIDs %>
        <% If blnWhqlTeam Or blnAdministrator Then%><br /><a href="#" onClick="RunWhqlWizard(<%= strID %>)">Add Submission</a><% End If %>
        </font></td>
    <td style="vertical-align:top; ">
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">WHQL Status:&nbsp;&nbsp;<a href="JavaScript:ShowWhqlStatus(<%= strID%>);"><span class="<%= strWhqlStatus%>"><%= strWhqlStatus%></span></a></span></td>
    <td style="vertical-align:top;"><font size="1">
        <a href="/iPulsar/ExcelExport/MDACompliance.aspx?PVID=<%= strID %>">Product Compliance Report</a>
        <% If blnMITTestLead Or blnAdministrator Then%><br /><a href="JavaScript:LeverageWhqlStatus(<%= strID%>);">Leverage Image Status</a><% End If %>
        </font></td>
    </tr>
  <tr>
    <td width="100" bgColor="cornsilk" valign="top"><strong><font size="1">Key&nbsp;Documents:&nbsp;&nbsp;</font></strong></td>
	<td valign="top" colspan="3"><font size="1">
	  <%if trim(strPddPath) <> "" then%>
		<a href="<%=strPddPath%>" target="new">PDD</a>&nbsp;|&nbsp;
	  <%else%>
		PDD&nbsp;|&nbsp;
	  <%end if%>

	  <%if trim(strScmPath) <> "" then%>
		  <a href="<%=strScmPath%>" target="new">SCM</a>&nbsp;|&nbsp;
	  <%else%>
		  SCM&nbsp;|&nbsp;
	  <%end if%>

	  <%if trim(strSTLStatusPath) <> "" then%>
		  <a href="<%=strSTLStatusPath%>" target="new">STL Status</a>&nbsp;|&nbsp;
	  <%else%>
			STL Status&nbsp;|&nbsp;
	  <%end if%>

	  <%if trim(strProgramMatrixPath) <> "" then%>
		  <a href="<%=strProgramMatrixPath%>" target="new">Product Data Matrices</a>&nbsp;|&nbsp;
	  <%else%>
		  Product Data Matrices&nbsp;|&nbsp;
	  <%end if%>

	  <%if trim(strAccessoryPath) <> "" then%>
		  <a href="<%=strAccessoryPath%>" target="new">Accessory Info</a>&nbsp;|&nbsp;
	  <%else%>
		  Accessory Info&nbsp;|&nbsp;
	  <%end if%>
	  <a href="javascript:OpenStatusOptions();">Change&nbsp;Log</a>&nbsp;|&nbsp;
  	  <a href="mailto:Charity.Harrison@hp.com?subject=Request access to Key Documents&body=The Key Documents Server Share is not managed by the Pulsar Team.  This email is to request access to the Server Share.">Request Access to Key Documents Server Share</a>&nbsp;|&nbsp;
      <%if trim(strMSPEKSExecutionPath) <> "" then%>
		  <a href="<%=strMSPEKSExecutionPath%>" target="new">MSPEKS (Execution)</a>
	  <%else%>
		  MSPEKS (Execution)
	  <%end if%>
	  <%'If CurrentUserID = 1396 Then %><!--//&nbsp;|&nbsp;<a href="reports/pdd_export.asp?PVID=<%=PVID%>">PDD Export</a>&nbsp;<font color="red">(BETA)</font>//--><%'End If %>
	    </font>
	</td></tr>


    <tr>
    <td style="width:100px; background-color:cornsilk; vertical-align:top; ">
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">Product Properties:&nbsp;&nbsp;</span></td>
    <td valign="top" colspan="3">
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">Created By:&nbsp;&nbsp;<span style="width:17%; font-weight:normal;"><%= CreatedBy%></span></span>
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">Created:&nbsp;&nbsp;<span style="width:17%; font-weight:normal;"><%= Created%></span></span>
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">Updated By:&nbsp;&nbsp;<span style="width:17%; font-weight:normal;"><%= UpdatedBy%></span></span>
        <span style="font-size:xx-small;font-weight:bold;white-space:nowrap">Updated:&nbsp;&nbsp;<span style="width:17%; font-weight:normal;"><%= Updated%></span></span>
    </tr>


</table>
</div>

<input type="hidden" id="txtID" name="txtID" value="<%=PVID%>" />
<input type="hidden" id="txtDisplayedStatus" name="txtDisplayedStatus" value="<%=strStatusID%>" />

<form id="ExportForm" target="new" method="post" action="mobilese/today/excelexport.asp">
<textarea ID="txtData" name="txtData" style="Display:none" rows="100" cols="20"></textarea>
</form>
<form id="ExportWordForm" target="new" method="post" action="exportoffice.asp?txtFormat=2">
<textarea ID="txtBody" name="txtBody" style="Display:none" rows="2" cols="20"></textarea>
</form>

<!--<form id=DetailsForm target=new method=post action="ExportDetails.asp"><TEXTAREA ID=Query name=Query style="Display:none" rows=2 cols=20 ></TEXTAREA><INPUT type="hidden" id=txtProdID name=txtProdID value="<%=PVID%>"></form>-->
<input type="hidden" id="txtClass" name="txtClass" value="<%=sClass%>">
<input type="hidden" id="txtDisplayedProduct" name="txtDisplayedProduct" value="<%=DisplayedProductName%>">
<%if blnAdministrator or blnSWPMRole then%>
	<input type="hidden" id="txtISPM" name="txtISPM" value="1">
<%else%>
	<input type="hidden" id="txtISPM" name="txtISPM" value="0">
<%end if%>

<%if blnMarketingAdmin then%>
	<input type="hidden" id="txtISMarketing" name="txtISMarketing" value="1">
<%else%>
	<input type="hidden" id="txtISMarketing" name="txtISMarketing" value="0">
<%end if%>
<% if blnAdministrator or blnPreinstall then%>
	<input type="hidden" id="txtISPreinstall" name="txtISPreinstall" value="1">
<%else%>
	<input type="hidden" id="txtISPreinstall" name="txtISPreinstall" value="0">
<%end if%>

<% if blnPreinstallPM or blnOdmPreinstallPM then%>
	<input type="hidden" id="txtISPreinstallPM" name="txtISPreinstallPM" value="1">
<%else%>
	<input type="hidden" id="txtISPreinstallPM" name="txtISPreinstallPM" value="0">
<%end if%>


<input type="hidden" id="txtDisplayedList" name="txtDisplayedList" value="<%=strDisplayedList%>">
<!--<a href="javascript: SortTable( 'TableDCR', 0,1);">Number</a>-->


<!--<a href="javascript:alert(sortedOn+ ' ' +sortDirection);">Show Sort</a>-->
<%if ItemsDisplayed <> 0 then%>
	<font size="1" face="verdana">Items Displayed: <%=ItemsDisplayed%></font>
<%end if%>

<input type="hidden" id="txtProductType" name="txtProductType" value="<%=strProdType%>">
<input type="hidden" id="txtFavs" name="txtFavs" value="<%=strFavs%>">
<input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=strFavCount%>">
<input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>">
<input type="hidden" id="txtQueryString" name="txtQueryString" value="<%=Request.QueryString%>">
<input type="hidden" id="txtServiceCommodityManager" name="txtServiceCommodityManager" value="<%=trim(lcase(blnServiceCommodityManager))%>">
<input type="hidden" id="txtUserPartner" name="txtUserPartner" value="<%=trim(CurrentUserPartner)%>">
<label ID="lblTest"></label>
<%


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
%>
<div style="display:none">  <!-- Popup Menus & Dialogs -->
<div id="localizationMenu">
    <div style="BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px">

    <% If blnAdministrator Or blnMarketingAdmin or ProductImageEdit = "1" Then %>
    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
        <font face="Arial" size="2">
        <span onClick="parent.location.href='javascript:CopyImage([ImageID])'" >&nbsp;&nbsp;&nbsp;Copy&nbsp;...</span></font>
    </div>
    <div><span><hr width="95%" /></span></div>
    <% End If %>

    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
    <font face="Arial" size="2">
    <span onClick="parent.location.href='javascript:DisplaySingleImage([ImageID])'" >&nbsp;&nbsp;&nbsp;View&nbsp;Matrix&nbsp;&nbsp;</span></font></div>
    <div><span><hr width="95%" /></span></div>

    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
    <font face="Arial" size="2">
    <span onClick="parent.location.href='javascript:DisplayImage([ImageID])'" >&nbsp;&nbsp;&nbsp;Properties</span></font></div>
    </div>
</div> <!-- localizationMenu -->

<div id="localizationMenuFusion">
    <div style="BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px">

    <% If blnAdministrator Or blnMarketingAdmin or ProductImageEdit = "1" Then %>
    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
    <font face=Arial size=2>
    <% If strFusionRequirements = 0 Then %>
        <span onClick="parent.location.href='javascript:CopyImageFusion([ImageID])'" >&nbsp;&nbsp;&nbsp;Copy&nbsp;...</span>    
    <% Else %>
        <span onClick="parent.location.href='javascript:CopyImagePulsar([ImageID])'" >&nbsp;&nbsp;&nbsp;Copy&nbsp;...</span>    
    <% End If%>
    </font></div>
        <!-- PBI 17835 / Task 18059; PBI 19513/ Task 20989 - Add Copy with Targeting to context menu -->
        <div><span><hr width="95%" /></span></div>
        <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
        <font face=Arial size=2>
            <% If strFusionRequirements = 0 Then %>
                <span onClick="parent.location.href='javascript:CopyWithTarget_Fusion([ImageID])'" >&nbsp;&nbsp;&nbsp;Copy with Targeting&nbsp;...</span> 
            <% Else %>
             <span onClick="parent.location.href='javascript:CopyWithTarget_Pulsar([ImageID])'" >&nbsp;&nbsp;&nbsp;Copy with Targeting&nbsp;...</span>  
            <% End If%> 
        </font></div>
    <div><span><hr width="95%" /></span></div>
    <% End If %>

    <div [DisplayOption1] onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
    <font face=Arial size=2>
    <span onClick="parent.location.href='javascript:LaunchProjectExplorer([ProductDropID])'" >&nbsp;&nbsp;&nbsp;Project&nbsp;Explorer&nbsp;(IRS)&nbsp;...</span></font></div>

    <div [DisplayOption1]><span><hr width="95%" /></span></div>

    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
        <font face="Arial" size="2">
            <% If strFusionRequirements = 0 Then %>
                <span onClick="parent.location.href='javascript:DisplaySingleImageFusion([ImageID])'" >&nbsp;&nbsp;&nbsp;Image&nbsp;Matrix&nbsp;&nbsp;</span>
            <% Else %>
                <span onClick="parent.location.href='javascript:DisplaySingleImagePulsar([ImageID])'" >&nbsp;&nbsp;&nbsp;Image&nbsp;Matrix&nbsp;&nbsp;</span>
            <% End If%>
        </font>
    </div>
    <div><span><hr width="95%" /></span></div>
    <div onMouseOver="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onMouseOut="this.style.background='white';this.style.color='black'">
    <font face="Arial" size="2">
    <% If strFusionRequirements = 0 Then %>
        <span onClick="parent.location.href='javascript:DisplayImageFusion([ImageID])'" >&nbsp;&nbsp;&nbsp;Properties</span>
    <% Else %>
        <span onClick="parent.location.href='javascript:DisplayImagePulsar([ImageID])'" >&nbsp;&nbsp;&nbsp;Properties</span>
    <% End If%>
    </font></div>
    </div>
</div> <!-- localizationMenuFusion -->

</div>
    <div style="display: none;">
        <div id="iframeDialog" title="Coolbeans">
            <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
        </div>
    </div>
    <div style="display: none;">
        <div id="divOpenMarketingNameUpdate" title="Coolbeans">
            <iframe frameborder="0" name="ifOpenMarketingNameUpdate" id="ifOpenMarketingNameUpdate"></iframe>
        </div>
    </div>

<%
    function Val(strText)
        dim strOutput
        dim i

        strOutput = ""
        for i = 1 to len(trim(strText))
            if isnumeric(mid(strText,i,1)) then
                strOutput = strOutput & mid(trim(strText),i,1)
            else
                exit for
            end if
        next
        Val = strOutput
    end function

    function createImageReleaseFiltrer(TabName)
		Dim rsImageReleases, strCmd, cn2, releaseID
		set cn2 = server.CreateObject("ADODB.Connection")
		cn2.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn2.Open
		Set rsImageReleases = server.CreateObject("ADODB.recordset")
		strCmd = "SELECT pvr.ID as ProductVersionReleaseId, r.name as ReleaseName From ProductVersion pv INNER JOIN ProductVersion_Release pvr on pv.id = pvr.productversionid  INNER JOIN ProductVersionRelease r on pvr.ReleaseID = r.id  where pv.id=" & PVID & " order by r.ReleaseYear desc, r.ReleaseMonth desc"
		
		rsImageReleases.open strCmd, cn2,  adOpenForwardOnly
		

		bFirstWrite = false                  
		releaseID = 0
		response.write "<tr>"
		response.write "<td style=""width:20%""><b>Release(s):</b></td>"
		response.write "<td style=""width:80%"">"
				
		if (Request("ProductReleaseID") = "" or IsNull(Request("ProductReleaseID"))) then 
			Response.Write "&nbsp;&nbsp;All"
		else
			Response.Write "&nbsp;&nbsp;<a href=""javascript:imageReleaseLink_onClick('','" & TabName & "')"">All</a>"
			releaseID = clng(Request("ProductReleaseID"))
		end if
		Do until rsImageReleases.EOF
			If Not bFirstWrite Then
				Response.Write "&nbsp;|&nbsp;"
			End If
			
			
		   if (releaseID > 0 and clng(rsImageReleases("ProductVersionReleaseId")) = releaseID) then 
				Response.Write server.HTMLEncode(rsImageReleases("ReleaseName"))			
		    else
				Response.Write "<a href=""javascript:imageReleaseLink_onClick('" & rsImageReleases("ProductVersionReleaseId") & "','" & TabName & "')"">" & server.HTMLEncode(rsImageReleases("ReleaseName")) & "</a>"
			end if 
			bFirstWrite = False
			rsImageReleases.MoveNext
		Loop
				
		response.write "</td>"
		response.write "</tr>"
		 
		rsImageReleases.Close
		cn2.close
		set rsImageReleases = nothing 
		set cn2 = nothing
		
	end function
	
	function createImageOSReleaseFiltrer(TabName, Fusion)
		'add more filter: OSRelease
		Dim rsImageOSReleases, strCmd3, cn3, osrID
		set cn3 = server.CreateObject("ADODB.Connection")
		cn3.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn3.Open
		Set rsImageOSReleases = server.CreateObject("ADODB.recordset")
		if Fusion = 0 then
			strCmd3 = "SELECT osr.id AS osrID, osr.Description AS osrName FROM Imagedefinitions d INNER JOIN ProductDrop1 pd ON pd.id = d.productdropid INNER JOIN Productversion_productdrop pvpd ON pvpd.productdropid = pd.id INNER JOIN OSRelease osr ON d.OSReleaseId = osr.Id WHERE pvpd.productversionid =" & PVID & " ORDER BY osr.id"
		else
			strCmd3 = "SELECT DISTINCT osr.id AS osrID, osr.Description AS osrName FROM ProductVersion pv INNER JOIN ProductVersion_Release pv_r ON pv_r.ProductVersionID = pv.id INNER JOIN ImageDefinitions im ON im.ProductVersionReleaseId = pv_r.ID INNER JOIN OSRelease osr ON osr.id = im.OSReleaseId WHERE pv.id=" & PVID & " ORDER BY osr.id"
		end if
	
		rsImageOSReleases.open strCmd3, cn3,  adOpenForwardOnly

		bFirstRound = false                  
		strosrID = 0
		response.write "<tr>"
		response.write "<td style=""width:20%""><b>Releases for Operating System:</b></td>"
		response.write "<td style=""width:80%"">"
				
		if (Request("ProductOSReleaseID") = "" or IsNull(Request("ProductOSReleaseID"))) then 
			Response.Write "&nbsp;&nbsp;All"
		else
			Response.Write "&nbsp;&nbsp;<a href=""javascript:imageOSReleaseLink_onClick('','" & TabName & "')"">All</a>"
			strosrID = clng(Request("ProductOSReleaseID"))
		end if
		Do until rsImageOSReleases.EOF
			If Not bFirstRound Then
				Response.Write "&nbsp;|&nbsp;"
			End If
			
			
		    if (strosrID > 0 and clng(rsImageOSReleases("osrID")) = strosrID) then 
				Response.Write server.HTMLEncode(rsImageOSReleases("osrName"))			
		    else
				Response.Write "<a href=""javascript:imageOSReleaseLink_onClick('" & rsImageOSReleases("osrID") & "','" & TabName & "')"">" & server.HTMLEncode(rsImageOSReleases("osrName")) & "</a>"
			end if 
			bFirstRound = False
			rsImageOSReleases.MoveNext
		Loop
				
		response.write "</td>"
		response.write "</tr>"

		
		rsImageOSReleases.Close
		cn3.close
		set rsImageOSReleases = nothing 
		set cn3 = nothing

	end function
%>
<br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<input type="hidden" id="txtSAAdmin" name="txtSAAdmin" value="<%=trim(SAAdmin)%>" />
<input type="hidden" id="inpSEPMProductUser" value="<%=blnSEPMProducts%>" />
<input type="hidden" id="txtProductBrandID" value="<%=m_BrandID%>" />
<input type="hidden" id="inpFusionRequirements" value="<%=strFusionRequirements%>" />
<input type="hidden" id="bNoteExists" value="0" /> 
<input type="hidden" id="TargetNotes" value="" />
</body>
</html>
<script type="text/javascript">
    var bIsPulsarProduct = '<%=bIsPulsarProduct%>';
    globalVariable.save(bIsPulsarProduct, 'product_type');

    $(window).load(function () {
        ValidatePagePermission("PMView", "Product");
    });

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

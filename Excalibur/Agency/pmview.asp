<%@  language="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%
Response.Clear 
Response.Buffer = False
	Dim regEx
	Set regEx = New RegExp
	regEx.Global = True

	regEx.Pattern = "[^0-9]"
	Dim PVID : PVID = regEx.Replace(Request.QueryString("ID"), "")
	'if PVID = "" Then PVID = 100
	
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
	Dim LinkAnchor : LinkAnchor = regEx.Replace(Request("Anchor"),"")

Dim rs, dw, cn, cmd, strSql
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
Dim isSysAdmin : isSysAdmin = false
Dim currentUser : currentUser = lcase(Session("LoggedInUser"))
Dim currentDomain
Dim currentUserPartner
Dim currentUserName
Dim currentUserID
Dim currentUserSysAdmin
Dim currentWorkgroupId
Dim faves
Dim faveCount
Dim productName
Dim displayedProductName
Dim devCenter
Dim SEPMID
Dim PMID
Dim productVersion
Dim displayedList
Dim currentUserDefaultTab
Dim productType

Dim AppRoot
AppRoot = Session("ApplicationRoot")

If instr(currentUser,"\") > 0 Then
	currentDomain = left(currentUser, instr(currentUser,"\") - 1)
	currentUser = mid(currentUser,instr(currentUser,"\") + 1)
End If

Dim isEditModeOn

'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim Security
	Set Security = New ExcaliburSecurity
	
	isSysAdmin = Security.IsSysAdmin()
	
	If isSysAdmin Then
		isEditModeOn = true
	End If
	
	Set Security = Nothing
'##############################################################################	


'
' Setup the data connections
'
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "spGetUserInfo")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, currentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, currentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then
	currentUserName = rs("Name") & ""
	currentUserID = rs("ID") & ""
	currentUserSysAdmin = rs("SystemAdmin")
	currentWorkgroupId = rs("WorkgroupID") & ""
	currentUserPartner = trim(rs("PartnerID") & "")
	currentUserDefaultTab = rs("DefaultProductTab") & ""

	faves = trim(rs("Favorites") & "")
	faveCount = trim(rs("FavCount") & "")
End If
rs.Close

on error resume next
	if trim(sList) <> "" then
		displayedList = sList
	else
		displayedList = "General"
	end if
		
	Response.Cookies("LastProductDisplayed") = PVID
on error goto 0

dim ShowItem
If currentUserPartner = "1" Then
	ShowItem = ""
Else
	ShowItem = "none"
End If

productType = "1"

Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, PVID
Set rs = dw.ExecuteCommandReturnRS(cmd)
Dim strError, strWarning

If (rs.EOF And rs.BOF) And PVID <> "-1" Then
	strError =  "Unable to find the selected program.<br /><font size=1>ID=" & PVID & "</font>"
	strError = strError & "<br /><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & PVID & ")""><font face=verdana size=1>Remove From Favorites</font></a>"
	strError = strError & "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
	strError = strError & "<span id=EditLink style=""Display:none""></span><span id=StatusLink style=""Display:none""></span><span id=menubar style=""Display:none""></span><span ID=Wait style=""Display:none""></span>"
	strError = strError & "<INPUT type=""hidden"" id=txtError name=txtError value=""1"">"
Else
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""0"">"
	productName = rs("Name") & " " & rs("Version") 
	displayedProductName = rs("Name") & " " & rs("Version")
	productVersion = rs("Version") & ""
	devCenter = trim(rs("DevCenter") & "")
	SEPMID = rs("sepmid")
	PMID = rs("PMID")
	if rs("SMID") & "" <> "" then
		PMID = PMID & "_" & rs("SMID")
	end if
	PMID = "_" & PMID & "_"
	productType = rs("TypeID") & ""
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
'or currentUserID = 391 or currentUserID = 397 
If currentUserSysAdmin or SEPMID = currentUserID or instr(trim(PMID),"_" & trim(currentUserID) & "_") > 0  Then
	isSysAdmin = true
End If

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />

	<meta http-equiv="cache-control" content="no-cache" />
	<meta http-equiv="expires" content="0" />
	<meta http-equiv="X-UA-Compatible" content="IE=9" />
	<title></title>
	<style type="text/css">
		body
		{
		}
		#loading
		{
			text-align: center;
			font: bold medium verdana;
		}
		#body
		{
			display: none;
		}
		#status
		{
			font-size: small;
			font-weight: bold;
		}
		#status dl dd
		{
			margin-left: 10px;
		}
		#status dl dt
		{
			padding-left: 3px;
		}
		#status dt
		{
			border: 1px solid grey;
		}
		#status dl
		{
			border: 1px solid black;
		}
		.StatusTable
		{
			width: 100%;
			border-collapse: collapse;
		}
		.StatusTable tbody td, .StatusTable tbody th, .ui-state-default, .ui-widget-content
		{
			font: x-small verdana;
			padding: 2px;
		}
		.red
		{
			background-color: #de2e43;
		}
		.yellow
		{
			background-color: #fdc643;
		}
		.green
		{
			background-color: #9cdb2c;
		}
		.highlight
		{
			background-color: lightsteelblue;
			background-image: none;
		}
		#status div
		{
			width: 99.7%;
			border: solid 1px black;
			padding-left: 3px;
		}
		
		.StatusTable thead tr
		{
			height: 15px;
		}
	</style>
	<link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
	<link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet"
		type="text/css" />
	<link href="<%= AppRoot %>/uploadify/uploadify.css" type="text/css" rel="stylesheet" />
	<script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
	<script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
	<script src="<%= AppRoot %>/uploadify/swfobject.js" type="text/javascript"></script>
	<script src="<%= AppRoot %>/uploadify/jquery.uploadify.v2.1.4.min.js" type="text/javascript"></script>
	<script src="<%= AppRoot %>/Agency/AgencyDetailsDialog.js" type="text/javascript"></script>
	<script src="<%= AppRoot %>/Agency/AgencyInitWorkflowDialog.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/spin.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/jquery.spin.js" type="text/javascript"></script>
    <script src="/pulsar2/js/userfavorite.js" type="text/javascript"></script>

	<script type="text/javascript">
		//
		// Code for the PMView Header
		//

		var _userFullName = '<%= currentUserName %>';

		String.prototype.trim = function () {
			return this.replace(/^\s+|\s+$/g, "");
		}

		String.prototype.ltrim = function () {
			return this.replace(/^\s+/, "");
		}

		String.prototype.rtrim = function () {
			return this.replace(/\s+$/, "");
		}


		//
		// BEGIN HEADER SUPPORT
		//

		function window_onload() {
			var anchor = document.getElementById("hidAnchor").value
			var strFavorites = "," + document.getElementById("txtFavs").value;
			var strID = document.getElementById("txtID").value
			var found = strFavorites.indexOf(",P" + strID.trim() + ",");

			if (txtClass.value == "") {
				EditLink.style.display = "none";
				RFLink.style.display = "none";
				AFLink.style.display = "none";
				StatusLink.style.display = "";
			}
			else if (found == -1) {
				EditLink.style.display = "";
				RFLink.style.display = "none";
				AFLink.style.display = "";
			}
			else {
				EditLink.style.display = "";
				RFLink.style.display = "";
				AFLink.style.display = "none";
			}

			EditLink.style.display = "";

			var lblAvCount = document.getElementById("lblAvCount");
			if (lblAvCount != null) {
				lblAvCount.innerHTML = "( " + hidAvCount.value + " AVs Displayed )";
			}

			document.location.hash = "#" + anchor;
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

		function ShowProperties(DisplayedID) {
			var strID
			strID = window.parent.showModalDialog("<%=AppRoot %>/mobilese/today/programs.asp?Commodity=0&ID=" + DisplayedID, "", "dialogWidth:675px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
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

		function SelectTab(strStep, blnLoad) {
			var i;
			var expireDate = new Date();

			expireDate.setMonth(expireDate.getMonth() + 12);
			document.cookie = "PMTab=" + strStep + ";expires=" + expireDate.toGMTString() + ";path=/";

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

		var oldColor;
		function HeaderMouseOver() {
			window.event.srcElement.style.cursor = "hand";
			oldColor = window.event.srcElement.style.color;
			window.event.srcElement.style.color = "red";
		}

		function HeaderMouseOut() {
			window.event.srcElement.style.color = oldColor;
		}
	</script>
	<script type="text/javascript">
		(function ($) {
			$.fn.isNullOrEmpty = function () {
				return (this === null || this === undefined || this === '');
			};
		})(jQuery);

		$(function () {

			//Setup Table Format
			$(".StatusTable th").addClass("ui-state-default");
			$(".StatusTable td").addClass("ui-widget-content");

			/*
			$(".StatusTable thead th").hover(function () {
			$(this).addClass("hover");
			}, function () {
			$(this).removeClass("hover");
			});
			*/
			$(".StatusTable tr").hover(function () {
				$(this).children("td").addClass("ui-state-hover2");
				$(this).addClass("hover");
			}, function () {
				$(this).children("td").removeClass("ui-state-hover2");
				$(this).removeClass("hover");
			});

			$(".StatusTable tbody tr").click(function () {
				var row = $(this)
				var deliverableVersionId = $(".deliverableVersionId", row).val();
				if (!$(deliverableVersionId).isNullOrEmpty()) {
					$("#pmViewCertStatusTabs").tabs("destroy");
					$("#pmViewCertStatusTab1").attr("href", "<%=AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=1&ProductVersionId=<%=PVID %>&DeliverableVersionId=" + deliverableVersionId);
					$("#pmViewCertStatusTab2").attr("href", "<%=AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=2&ProductVersionId=<%=PVID %>&DeliverableVersionId=" + deliverableVersionId);
					$("#pmViewCertStatusTab3").attr("href", "<%=AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=3&ProductVersionId=<%=PVID %>&DeliverableVersionId=" + deliverableVersionId);
					$("#pmViewCertStatusTab4").attr("href", "<%=AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=4&ProductVersionId=<%=PVID %>&DeliverableVersionId=" + deliverableVersionId);
					$("#pmViewCertStatusTabs").tabs({
						ajaxOptions: {
							error: function (xhr, status, index, anchor) {
								$(anchor.hash).html("Couldn't load this tab. We'll try to fix this as soon as possible. ");
							}
						}
					});
					$("#pmViewCertStatus").dialog("open");
				}
			});

			if ($("#errorText").text() == '') {
				$("#error").hide();
			}

			if ($("#warningText").text() == '') {
				$("#warning").hide();
			}

			$(".Link").hover(function () {
				$(this).addClass("hover");
			}, function () {
				$(this).removeClass("hover");
			});

			$("#initWorkflow").click(function () {
				InitDialog.Show();
			});

			$("#showFrame").click(function () {
				ShowIframeDialog();
			});

			$("#iframeDialog").dialog({
				modal: true,
				autoOpen: false,
				width: 800,
				height: 800
			});

			$("#pmViewCertStatus").dialog({
				modal: true,
				autoOpen: false,
				width: 650,
				height: 400,
				buttons: {
					"Close": function () {
						$(this).dialog("close");
					}
				}
			});

			$("#pmViewCertStatusTabs").tabs();
			//InitDialog.Setup();

			//Hide Loading panel and show the rest of the body.
			$("#ToolMenu").hide();
			$("#loading").hide();
			$("#body").fadeIn('slow');
		});

		function ShowIframeDialog() {
			$("#iframeDialog iframe").attr("width", "95%");
			$("#iframeDialog iframe").attr("height", "95%");
			$("#iframeDialog iframe").attr("src", "Agency.asp?ID=5000");
			$("#iframeDialog").dialog("open");
		}

	   
	</script>
</head>
<body>
	<div id="PMViewHeader">
		<!-- PMView Header -->
		<span id="productNameTitle" style="font: bold medium Verdana;">
			<%= productName%>
			Information</span><br />
		<br />
		<%    if (clng(PVID) = 344 or clng(PVID) = 347 or clng(PVID) = 1107) and (not isSysAdmin) then %>
		<td nowrap id="EditLink" style="display: none">
		</td>
		<%else%>
		<td nowrap id="EditLink" style="display: none">
			<font size="1">
				<!--Contact the PM to edit product properties-->
			</font>
		</td>
		<%end if%>
        <span id="loadingProgress"></span>
		<span id="RFLink" style="display: none"><a href="javascript:RemoveFavorites(<%=PVID%>)">
			<font face="verdana" size="1">Remove From Favorites</font></a> | </span><span id="AFLink"
				style="display: none"><a href="javascript:AddFavorites(<%=PVID%>)"><font face="verdana"
					size="1">Add To Favorites</font></a> | </span><span id="StatusLink" style="display: none">
						<a href="<%=AppRoot %>/Productstatus.asp?Product=<%=displayedProductName%>&ID=<%=PVID%>">
							<font face="verdana" size="1">Real-Time Status Report</font></a> |</span>
		<%if displayedList <> currentUserDefaultTab and trim(productType) <> "2" then%>
		<span id="DefaultTabLink"><a href="javascript:SetDefaultDisplay('<%=displayedList%>',<%=currentUserId%>)">
			<font face="verdana" size="1">Set Default List</font></a></span>
		<%end if%>
		<br />
		<br />
		<div>
			<table class="PageTabs">
				<tr bgcolor="<%=strTitleColor%>">
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('DCR',1)">Change&nbsp;Requests</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Action',1)">Action&nbsp;Items</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('OTS',1)">Observations</a>
					</td>
					<td width="10" class="Selected">
						&nbsp;Certification
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('PMR',1)">SMR</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Service',1)">Service</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('General',1)">General</a>
					</td>
				</tr>
				<tr bgcolor="<%=strTitleColor%>">
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Requirements',1)">Requirements</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Country',1)">Localization</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Local',1)">Images</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Deliverables',1)">Deliverables</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Schedule',1)">Schedule</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('SCM',1)">Supply&nbsp;Chain</a>
					</td>
					<td width="10">
						&nbsp;<a href="javascript:SelectTab('Documents',1)">Documents</a>
					</td>
				</tr>
			</table>
		</div>
		<div class="ui-helper-hidden">
			<table class="DisplayBar">
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
									RSL | <a href="#" onclick="ServiceReport('spb');">SPB</a> | <a href="#" onclick="ServiceReport('ecr');">
										ECR</a> | <a href="#" onclick="ServiceReport('calls');">Call Data</a>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
		<input type="hidden" id="txtClass" name="txtClass" value="<%=sClass%>" />
		<input type="hidden" id="txtID" name="txtID" value="<%=PVID%>" />
		<input type="hidden" id="txtFavs" name="txtFavs" value="<%=faves%>" />
		<input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=faveCount%>" />
		<input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>" />
		<input type="hidden" id="hidLastPublishDt" value="<%=m_LastPublishDt %>" />
		<input type="hidden" id="hidBusinessID" value="<%=m_BusinessID %>" />
	</div>
	<!-- PMView Header -->
	<div id="ToolMenu" class="us-helper-hidden">
		<span class="Title">Tools:&nbsp;</span> <span id="initWorkflow" class="Link">Init Workflow</span>
		<span id="showFrame" class="Link Hidden">Show Frame</span>
	</div>
	<div id="loading">
		<img src="<%= AppRoot %>/images/loading24.gif" alt="Loading" />
		Loading...</div>
	<div id="body">
		<div class="ui-widget ui-helper-clearfix">
			<div id="error" class="ui-state-error ui-corner-all" style="padding: 0 .7em;">
				<p>
					<span class="ui-icon ui-icon-alert" style="float: left; margin-right: .3em;"></span>
					<strong>Alert:&nbsp;</strong><span id="errorText"><%= strError & ""%></span></p>
			</div>
			<div id="warning" class="ui-state-highlight ui-corner-all" style="padding: 0 .7em;">
				<p>
					<span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
					<strong>Notice:&nbsp;</strong><span id="warningText"><%= strWarning & ""%></span></p>
			</div>
		</div>
		<div id="newPmView" style="padding-top: 10px;">
			<table class="StatusTable">
				<thead>
					<tr>
						<th rowspan="2">
							Version ID
						</th>
						<th rowspan="2">
							Name
						</th>
						<th rowspan="2">
							Version
						</th>
						<th rowspan="2">
							Targeted
						</th>
						<th colspan="2">
							Doucments
						</th>
						<th colspan="2">
							Countries
						</th>
					</tr>
					<tr>
						<th>
							Completed
						</th>
						<th>
							Pending
						</th>
						<th>
							Completed
						</th>
						<th>
							Blocked
						</th>
					</tr>
				</thead>
				<tbody>
					<%
			Set cmd = dw.CreateCommandSP(cn, "usp_AgencyStatusSelectProductOverview") 
			dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, PVID  
			Set rs = dw.ExecuteCommandReturnRS(cmd)

			If rs.Eof And rs.Bof Then
				Response.Write "<tr><td colspan=""8"">No Certifications To Report</td></tr>"
			End If

			Do Until rs.Eof
				Response.Write "<tr><td>"
				Response.Write "<input type=""hidden"" class=""deliverableVersionId"" value=""" & rs("DeliverableVersionId") & """>"
				Response.Write rs("DeliverableVersionId") & ""
				Response.Write "</td><td>"
				Response.Write rs("DeliverableName") & ""
				Response.Write "</td><td>"
				Response.Write rs("Version") & ""
				Response.Write "</td><td>"
				Response.Write rs("Targeted") & ""
				Response.Write "</td><td>"
				Response.Write rs("CompletedDocumentCount") & ""
				Response.Write "</td><td>"
				Response.Write rs("PendingDocumentCount") & ""
				Response.Write "</td><td>"
				Response.Write rs("CompleteCountryCount") & ""
				Response.Write "</td><td>"
				Response.Write rs("BlockedCountryCount") & ""
				Response.Write "</td></tr>"

				rs.MoveNext
			Loop
		
			rs.Close

					%>
				</tbody>
			</table>
		</div>
		<div id="dialogContainer" style="display: none;">
			<div id="pmViewCertStatus" title="Certification Status">
				<div id="pmViewCertStatusTabs">
					<ul>
						<li><a id="pmViewCertStatusTab1" href="<%= AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=1&ProductVersionId=100">
							Completed Documents</a></li>
						<li><a id="pmViewCertStatusTab2" href="<%= AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=2&ProductVersionId=100">
							Pending Documents</a></li>
						<li><a id="pmViewCertStatusTab3" href="<%= AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=3&ProductVersionId=100">
							Completed Countries</a></li>
						<li><a id="pmViewCertStatusTab4" href="<%= AppRoot %>/Agency/AgencyPmViewStatus.aspx?ReportTypeId=4&ProductVersionId=100">
							Blocked Countries</a></li>
					</ul>
				</div>
			</div>
			<div id="iframeDialog" title="Coolbeans">
				<iframe frameborder="0"></iframe>
			</div>
			<input type="hidden" id="currentUser" value="<%=currentUserName %>" />
			<input type="hidden" id="productVersionId" value="<%=PVID %>" />
		</div>
	</div>
	<!-- DivBody -->
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

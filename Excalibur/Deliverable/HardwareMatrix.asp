<%@ Language="VBScript" %> 

<% Option Explicit
   Server.ScriptTimeout = 1800 %>
<%
	if request("FileType")= 1  or request("FileType")= 2  then
		Response.ContentType = "application/vnd.ms-excel"
	else
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
	end if

%>
<!DOCTYPE HTML>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script src="../includes/client/jquery-1.10.2.js"></script>
<script src="../includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.js"></script>
<% if request("FileType")<> 1  and request("FileType")<> 2  then%>
	<link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
<% end if%>

<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

var oPopup = window.createPopup();
var SelectedRow;

function Export(strID){
	if (txtCurrentFilter.value == "")
		window.open (window.location.href + "?FileType=" + strID);
	else	
		window.open (window.location.pathname + "?" + txtCurrentFilter.value + "&FileType=" + strID);
}

function button1_onclick(NewColor) {
	//MyTable.borderColor = NewColor;
	Row1.bgColor=NewColor;
	Row2.bgColor=NewColor;
}

function button2_onclick() {
	COL1.style.width=0;
	//COL1.width=10;
}

function Commodity_onclick() {
	var RowElement;
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
		if(RowElement.style.backgroundColor=="cornflowerblue")//lightgoldenrodyellow
			RowElement.style.backgroundColor="";
		else
			RowElement.style.backgroundColor="cornflowerblue";//lightgoldenrodyellow

		if 	(typeof(SelectedRow) != "undefined")
			{
			if (SelectedRow!=RowElement)
				SelectedRow.style.backgroundColor="";
			}
		SelectedRow=RowElement;

}
function Commodity_onmouseover() {
	var RowElement;
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
	
	if (RowElement.style.backgroundColor == "")
		{
		RowElement.style.backgroundColor="#99ccff";
		RowElement.style.cursor="hand";
		}
}

function Commodity_onmouseout() {
	var RowElement;
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
	
	if (RowElement.style.backgroundColor == "#99ccff")
		{
		RowElement.style.backgroundColor="";
		}
}


function DisplayVersion(VersionID){
	var strResult;
	
	strResult = window.showModalDialog("../WizardFrames.asp?Type=1&ID=" + VersionID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strResult) != "undefined")
		{
			window.location.reload(true);
		}
	if 	(typeof(SelectedRow) != "undefined")
		{
			SelectedRow.style.backgroundColor="";
		}

}

function SwitchFilterView(strType){
	if (strType == 1)
		{
		QuickLinks.style.display="none";
		FilterBox.style.display="";
		}
	else if (strType == 2)
		{
		QuickLinks.style.display="";
		FilterBox.style.display="none";
		}
}


function window_onload() {
	lblLoad.style.display = "none";
	if (txtScrollToRow.value != "")
		document.all("Row" + txtScrollToRow.value).scrollIntoView();
	window.name = "HardwareMatrix";
	//self.focus();
}


function Test(ID){
	
    DisplayArea.innerHTML = divQuickReports.innerHTML;
}

function ShowChanges(ID) {    

    $("#releaseNotes").dialog({
        autoopen: false, bigframe: true, modal: true, height: 350, width: 400, closeonescape: true, resizable: false, draggable: true, position:{my: "center", at: "center", of: window},
        buttons: { Close: function () { $(this).dialog("close"); } },
        open: function (ev, ui) {$("#rnFrame").attr('src', "../ReleaseNotes.asp?ID=" + ID); $(".ui-dialog-titlebar-close", ui.dialog).hide();},
        close: function () { $("#rnFrame").attr('src', ""); }
    });

    if ($("#releaseNotes").dialog("isOpen") == true) {
        $("#releaseNotes").dialog("close");
    }

    $("#releaseNotes").dialog("open");
    $('#releaseNotes').parent().css({ 'z-index': '100' });
    $(".ui-widget-overlay").css("opacity", "0"); //remove overlay
}

function cbo_onchange(cboId) {
    var url = $(cboId).val()
    window.location.href = url;
}

//-->
</script>
<% if request("FileType")<> 1  and request("FileType")<> 2  then%>
	<LINK href="../style/Excalibur.css" type="text/css" rel="stylesheet" >
<% end if%>
</head>
<style type= "text/css">
.CatRow TD{
	FONT-SIZE:xx-small;
	FONT-FAMILY:Verdana;
	Background-COLOR=DarkSeaGreen;
}
.RootRow TD{
	FONT-SIZE:xx-small;
	FONT-FAMILY:Verdana;
	//Background-COLOR=LightSteelBlue;
	Background-COLOR=#cccc99;
}
.HeaderRow TH{
	FONT-SIZE:xx-small;
	FONT-FAMILY:Verdana;
//	Background-COLOR=Lavender;
	Background-COLOR=Beige;
}

TD{
	FONT-SIZE:xx-small;
	FONT-FAMILY:Verdana;
}
TH{
	FONT-SIZE:xx-small;
	TEXT-ALIGN: left;
	FONT-FAMILY:Verdana;
}

A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
A
{
    COLOR: blue
}
</style>
<body onload="return window_onload()">
<%if request("FileType")<> 1  and request("FileType")<> 2 then %>
	<DIV ID=lblLoad>Loading. Please Wait...<BR /><BR /></DIV>
<%else%>
	<font ID=lblLoad></font>
<%end if%>


<%
'Response.Flush

	dim cn
	dim rs
	dim strSQL
	dim LastCategory
	dim LastRoot
	dim strDCR
	dim strBaseSub
	dim counter
	dim ProductArray
	dim ProductIDArray
	dim ProductList
	dim ExtraColumnCount
	dim ProductIDList
    dim ProductReleaseIDList
    dim ProductReleaseIDArray
	dim ProductBuffer
	dim strFilter
	dim ColumnCount
	dim strProduct
	dim strSpin
	dim strTestStatus
	dim EOLBGColor
	dim strEOLDate
	dim strDCRTitle
	dim rowCounter
	dim strSQLSelect
	dim strSQLSelectProdList
    dim strSQLSelectProdListForDisplay
	dim strSQLSelectFamilyList
    dim strSQLSelectPulsarProductList
	dim strSQLSelectCatList
	dim strSQLSelectDelList
	dim strSQLSelectPhaseList
	dim LastVersion
	dim ProductLoop
	dim ProductBuckets
	dim strBucket
	dim BucketParts
	dim ProductSubArray
	dim i
	dim rs2
	dim ShowFilters
	dim ShowQuickLinks
	dim strListHeaderName
	dim CurrentUserID
	dim cm
	dim CurrentUser
	dim p
	dim FilterArray
	dim blnAllowQuickFilters
	dim strFullFilter
	dim FilterValueArray
	dim strCategory 
	dim strRoot
	dim strFamily
	dim	strProducts 
	dim blnVersionChange
	dim TestDateBGColor
	dim ScrollToRow
	dim ProductCount
	dim DCRCellCount
	dim strQueryString
	dim strQueryString2
	dim strTestColor
	dim strGreenSpecColor
	dim strPilotStatus
	dim strAccessoryStatus
	dim strPilotColor
	dim strAccessoryColor
	dim StatusCellCount
	dim strHistoryFilter 
	dim LastSub
	dim blnFiltersSelected
	dim LastSubassemblyBase
	dim LastRootID
	dim LastNativeBase
    dim blnShowTestStep	
    dim blnShowReleaseStep	
    dim blnShowCompleteStep	
    dim strStepFilter
    dim strReportDateRange
    dim strProductsPulsar
    dim strProductReleaseIDs
    dim strSQLGetPulsarProducts
    dim responseBuffer
    dim firstCache   

    strProductsPulsar = ""
    strReportDateRange = ""
	lastSub = ""

    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.ConnectionTimeout = 300
	cn.IsolationLevel=256
	cn.commandtimeout = 300
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    strProducts = request("lstProducts")                        'get list of products
    strProductReleaseIDs = request("lstProductsPulsar")            'get list of Pulsar products
    
    if strProductReleaseIDs <> "" then
        'get the pulsar products from the productreleaseid
        strSQLGetPulsarProducts = "SELECT DISTINCT ProductVersionID FROM ProductVersion_Release WITH (NOLOCK) WHERE ID IN (" & strProductReleaseIDs & ")"
	    rs.Open strSQLGetPulsarProducts,cn,adOpenForwardOnly
	    do while not rs.EOF
		    strProductsPulsar =  strProductsPulsar & "," & rs("ProductVersionID")			
		    rs.MoveNext
	    loop
	    rs.Close        
    end if
    
    if strProducts <> "" and strProductsPulsar <> "" then       'the next if statements combine the list of legacy and pulsar products into one list
        strProducts = strProductsPulsar + "," + strProducts
    end if

    if strProducts = "" and strProductsPulsar <> "" then
        strProducts =  strProductsPulsar
    end if

    if strProducts <> "" and strProductsPulsar = "" then
        strProducts =  strProducts
    end if

	if left(strProducts,1) = "," then   
		strProducts = mid(strProducts,2)
	end if
       	
	if request("chkTest") <> "on" and request("chkRelease") <> "on" and request("chkComplete") <> "on" then
	    blnShowTestStep = false
	    blnShowReleaseStep = false
	    blnShowCompleteStep = true
	else
        if request("chkTest") = "on" then
            blnShowTestStep = true
        else
            blnShowTestStep = false
        end if
        if request("chkRelease") = "on" then
            blnShowReleaseStep = true
        else
            blnShowReleaseStep = false
        end if
        if request("chkComplete") = "on" then
            blnShowCompleteStep = true
        else
            blnShowCompleteStep = false
        end if
	end if
	
	if instr(lcase(trim(Request.Form)),"cboprofile")>0 then 
		strQueryString = trim(FormatInputStrings())
	else
		strQueryString = trim(Request.QueryString)
	end if

	if strQueryString = "" or strQueryString = "ReportFormat=1" or strQueryString = "ReportFormat=2" or strQueryString = "ReportFormat=3" or strQueryString = "ReportFormat=4" or strQueryString = "ReportFormat=5" or strQueryString = "ReportFormat=6" then
		blnFiltersSelected = false
	else
		blnFiltersSelected = true
	end if
		
	FilterArray = split(strQueryString,"&")
	blnAllowQuickFilters = true
	for each strFullFilter in FilterArray
		FilterValueArray = split(strFullFilter,"=")
		if ubound(FilterValueArray)<> 1 then
			blnAllowQuickFilters = false
			exit for
		elseif lcase(trim(FilterValueArray(0))) <> "lstproducts" and lcase(trim(FilterValueArray(0))) <> "lstroot" and lcase(trim(FilterValueArray(0))) <> "lstsubassembly" and lcase(trim(FilterValueArray(0))) <> "lstcategory" and lcase(trim(FilterValueArray(0))) <> "lstphase" and lcase(trim(FilterValueArray(0))) <> "lstfamily" and lcase(trim(FilterValueArray(0))) <> "reportformat" and lcase(trim(FilterValueArray(0))) <> "lstproductspulsar"then
			blnAllowQuickFilters = false
			exit for
		elseif (not isnumeric(FilterValueArray(1))) or instr(FilterValueArray(1),",") > 0  or instr(FilterValueArray(1),";") > 0  or instr(FilterValueArray(1),":") > 0  or instr(FilterValueArray(1),".") > 0 then
			blnAllowQuickFilters = false
			exit for
		end if
	next
		
	'Get User
	dim CurrentDomain
	dim CurrentUserPartnerID
	CurrentUser = lcase(Session("LoggedInUser"))

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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserPartnerID = rs("PartnerID") & ""
	else
		CurrentUserID = 0
		CurrentUserPartnerID = 0
	end if
	rs.Close

	if blnAllowQuickFilters then
		ShowFilters = ""
		ShowQuickLinks = "none"
	else
		ShowFilters = "none"
		ShowQuickLinks = ""
	end if
	
	'Defined Selects
    strSQLSelectProdListForDisplay = "Select distinct pv.ID, pv.DotsName "
	strSQLSelectProdList = "Select distinct pv.ID, ProductName = pv.Dotsname, Dotsname = pv.Dotsname + case isnull(pv.FusionRequirements,0) when 1 then ' - ' + pvrelease.Name else '' end , ReleaseID = isnull(pvr.ReleaseID,0),pvrelease.ReleaseYear, pvrelease.ReleaseMonth "
	strSQLSelectFamilyList = "Select distinct f.ID, f.Name, productversionID=0 "
    strSQLSelectPulsarProductList = "select distinct f.ID, Name= pv.productname, productversionID=pv.id "
	strSQLSelectCatList = "Select distinct c.ID, c.Name "
	strSQLSelectPhaseList = "Select distinct ps.id, ps.name as Phase "
	if request("ReportFormat")="2" then
		strSQLSelectDelList = "Select distinct r.ID, r.Name, case when prr.ID is null then isnull(pr.base,'') else isnull(prr.base,'') end as Subassembly "
	elseif request("ReportFormat")="5" then
		strSQLSelectDelList = "Select distinct r.ID, r.Name, case when prr.ID is null then coalesce(pr.servicebase,pr.base) else coalesce(prr.servicebase,prr.base) end as Subassembly "
	else
		strSQLSelectDelList = "Select distinct r.ID, r.Name, '' as Subassembly "
	end if
	strFilter=""
	
	if CurrentUserPartnerID <> 1 then
		strFilter = strFilter & " and (pv.Partnerid in (" & clng(CurrentUserPartnerID) & ") or (pv.Partnerid in (SELECT ProductPartnerId FROM PartnerODMProductWhitelist WHERE UserPartnerId = " & clng(CurrentUserPartnerID) & ")) ) " 
	end if
	
	if request("lstVendor") <> "" then
        dim allVendorStr
		dim selectedVendor
		selectedVendor = scrubsql(request("lstVendor"))
		allVendorStr = "SELECT ID FROM VENDOR WITH (NOLOCK) WHERE ID IN (SELECT ID FROM VENDOR WITH (NOLOCK) WHERE NAME IN (SELECT DISTINCT NAME FROM VENDOR WITH (NOLOCK) WHERE ID IN(" & selectedVendor & ")))"
		rs.Open allVendorStr,cn,adOpenForwardOnly
		do while not rs.EOF
			if (Instr(selectedVendor, rs("ID"))) < 1 then
				selectedVendor = selectedVendor & ", " & rs("ID")
			end if
			rs.MoveNext
		loop
		rs.Close
         strFilter = strFilter & " and v.vendorid in (" & selectedVendor & ") "
		'strFilter = strFilter & " and v.vendorid in (" & scrubsql(request("lstVendor")) & ") " 
	end if
	
	if trim(request("cboRohs")) = "1" then
		strFilter = strFilter & " and v.rohsid=1 " ' & clng(request("cboRohs")) & " " 
    elseif trim(request("cboRohs")) = "2" then
		strFilter = strFilter & " and v.greenspecid=1 " '& clng(request("cboRohs")) & " " 
	end if

	if request("cboEOL") <> "" then
		strFilter = strFilter & " and v.active= " & clng(request("cboEOL")) & " " 
	end if
	
	if request("lstPartner") <> "" then
		strFilter = strFilter & " and pv.Partnerid in (" & scrubsql(request("lstPartner")) & ") " 
	end if
	
	if request("lstQualStatus") <> "" then
	    if instr(request("lstQualStatus"),"-1") > 0 then
	    	'Modified from RiskRelease=1 to pd.RiskRelease=1 for Defect : 135436 -> Task 135437 by PMV Pandian
    		strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and (pd.TestStatusid in (" & scrubsql(request("lstQualStatus")) & ") or (pd.TestStatusid=5 and pd.RiskRelease=1))) or (isnull(pv.FusionRequirements, 0) = 1 and (pdr.TestStatusid in (" & scrubsql(request("lstQualStatus")) & ") or (pdr.TestStatusid=5 and pdr.RiskRelease=1))))" 
        else
    		strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.TestStatusid in (" & scrubsql(request("lstQualStatus")) & ")) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.TestStatusid in (" & scrubsql(request("lstQualStatus")) & "))) " 
	    end if
	end if

	if request("lstDevInput") <> "" then
		strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.DeveloperNotificationStatus in (" & scrubsql(request("lstDevInput")) & ")) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.DeveloperNotificationStatus in (" & scrubsql(request("lstDevInput")) & "))) " 
	end if

	if request("chkSCRestricted") <> "" then
		strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.SupplyChainRestriction=1) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.SupplyChainRestriction=1)) " 
	end if

	if request("lstDevManager") <> "" then
		strFilter = strFilter & " and r.DevManagerid in (" & scrubsql(request("lstDevManager")) & ") " 
	end if

	if request("lstCommodityPM") <> "" and request("lstCommodityPM") <> "0" then
		strFilter = strFilter & " and pv.PDEID in (" & scrubsql(request("lstCommodityPM")) & ") " 
	end if

	if request("lstPhase") <> "" and request("lstPhase") <> "0" then
		strFilter = strFilter & " and ps.ID in (" & scrubsql(request("lstPhase")) & ") " 
	end if

	if request("lstCategory") <> "" then
		strFilter = strFilter & " and r.Categoryid in (" & scrubsql(request("lstCategory")) & ") " 		
	end if

	if request("lstRoot") <> "" then
		strFilter = strFilter & " and r.id in (" & scrubsql(request("lstRoot")) & ") " 
	end if

	if request("lstSubassembly") <> "" and request("ReportFormat")="2" then
		strFilter = strFilter & " and ((prr.ID is null and pr.base in ('" & scrubsql(request("lstSubassembly")) & "')) or (prr.ID is not null and prr.base in ('" & scrubsql(request("lstSubassembly")) & "'))) " 
	elseif request("lstSubassembly") <> "" and request("ReportFormat")="5"  then
		strFilter = strFilter & " and (( prr.ID is null and (pr.servicebase is null and pr.base in ('" & scrubsql(request("lstSubassembly")) & "')) or pr.servicebase in ('" & scrubsql(request("lstSubassembly")) & "') ) or ( prr.ID is not null and (prr.servicebase is null and prr.base in ('" & scrubsql(request("lstSubassembly")) & "')) or prr.servicebase in ('" & scrubsql(request("lstSubassembly")) & "') )) " 
	end if

	if request("lstFamily") <> "" then
		strFilter = strFilter & " and pv.productfamilyid in (" & scrubsql(request("lstFamily")) & ") " 
	end if



'*******Process Product Groups 
	if request("lstProductGroups") <> "" then
		dim ProductGroupsArray
		dim ProductGroupArray
		dim strProductGroup
		dim lastProductGroup
		dim strProductGroupFilter
		dim strCycleList
		dim strGroupSQL
		ProductGroupsArray = split(request("lstProductGroups"),",")
		lastProductGroup = 0
		strProductGroupFilter = ""
		strCycleList = ""
        strGroupSQL = ""
		for each strProductGroup in ProductGroupsArray
			if instr(strProductGroup,":")>0 then
				ProductGroupArray = split(strProductGroup,":")
				if trim(lastproductgroup) <> "0" and trim(ProductGroupArray(0)) <> "2" and lastproductgroup <> trim(ProductGroupArray(0)) then
					strProductGroupFilter = strProductGroupFilter & " ) and  "
				end if
				if trim(lastproductgroup) <> trim(ProductGroupArray(0)) then
					if trim(ProductGroupArray(0)) = "1" then
						strProductGroupFilter = strProductGroupFilter & " ( partnerid = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					elseif trim(ProductGroupArray(0)) = "2" then
						strCycleList = strCycleList & "," & clng(ProductGroupArray(1))
					elseif trim(ProductGroupArray(0)) = "3" then
						strProductGroupFilter = strProductGroupFilter & " ( devcenter = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					elseif trim(ProductGroupArray(0)) = "4" then
						strProductGroupFilter = strProductGroupFilter & " ( productstatusid = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					end if
				else
					if trim(ProductGroupArray(0)) = "1" then
						strProductGroupFilter = strProductGroupFilter & " or partnerid = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					elseif trim(ProductGroupArray(0)) = "2" then
						strCycleList = strCycleList & "," & clng(ProductGroupArray(1))
					elseif trim(ProductGroupArray(0)) = "3" then
						strProductGroupFilter = strProductGroupFilter & " or devcenter = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					elseif trim(ProductGroupArray(0)) = "4" then
						strProductGroupFilter = strProductGroupFilter & " or productstatusid = " & trim(ProductGroupArray(1))
						lastproductgroup = trim(ProductGroupArray(0))
					end if
				end if
			end if
		next
		    if strProductGroupFilter <> "" then
			    strGroupSQl = strGroupSQL & " and ( " & ScrubSQL(strProductGroupFilter) &  ") ) "
		    end if
		    if strCycleList <> "" then
			    strGroupSQl = strGroupSQL & " and id in (Select ProductVersionid from product_program with (NOLOCK) where programid in (" & mid(strCycleList,2) &  ")) "
		    end if
		    if strGroupSQl <> "" then
		        strGroupSQl = mid(strGroupSQL,5)
		        rs.open "Select ID from productversion with (NOLOCK) where " & strgroupSQL,cn
		        do while not rs.eof
	                strProducts = strProducts & ", " & rs("ID") 
		            rs.movenext
		        loop
		        rs.close    
		    end if
		    if strProducts = ""  then 
	            strProducts = "0"
	        elseif left(strproducts,2) = ", " then
	            strproducts = mid(strproducts,3) 
	        end if
	end if

'*******End Product Groups
	if strProducts <> "" then
		strFilter = strFilter & " and pv.id in (" & scrubsql(strProducts) & ",0) " 
	else
'		strFilter = strFilter & " and pv.oncommoditymatrix=1 " 'Only show products specified as On the Matrix if no products are specifically selected
		strFilter = strFilter & " and pv.id <> 100 and pv.oncommoditymatrix=1 and pv.productstatusid<5 " 'Only show products specified as On the Matrix if no products are specifically selected
	end if
	
	if request("lstTeamID") <> "" then
		strFilter = strFilter & " and c.teamid in (" & scrubsql(request("lstTeamID")) & ") " 
	end if
	
	if trim(request("txtAdvanced")) <> "" then
		strFilter = strFilter & " and ( " & scrubsql(request("txtAdvanced")) & " ) " 
	end if
	
	if request("txtCompleteDateStart") <> "" and request("txtCompleteDateEnd") <> "" then
		if isdate(request("txtCompleteDateStart")) and isdate(request("txtCompleteDateEnd")) then
			strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.teststatusid = 3 and pd.TestDate between '" & request("txtCompleteDateStart") & "' and '" & request("txtCompleteDateEnd") & "') or (isnull(pv.FusionRequirements, 0) = 1 and pdr.teststatusid = 3 and pdr.TestDate between '" & request("txtCompleteDateStart") & "' and '" & request("txtCompleteDateEnd") & "')) "
		end if
	end if
	
	if request("ReportSplit") = "1" then
		strFilter = strFilter & " and upper(pv.dotsname) >= 'A'  and upper(pv.dotsname) < 'N' " 
	elseif request("ReportSplit") = "2" then
		strFilter = strFilter & " and upper(pv.dotsname) >= 'N' " 
	end if

	if request("chkChangeType") <> "" then
		strHistoryFilter = " and pd.id in (" & BuildHistoryFilter() & ") "
	else 
		strHistoryFilter = ""
	end if
    if blnShowCompleteStep or  blnShowReleaseStep or blnShowTestStep then
    	strFilter = strFilter & " and v.status <> 5 and v.location not like 'Development%' and ((isnull(pv.FusionRequirements, 0) = 0 and (pd.teststatusid <> 1 or ( pd.teststatusid = 1 and pd.DeveloperNotificationStatus=1))) or (isnull(pv.FusionRequirements, 0) = 1 and (pdr.teststatusid <> 1 or ( pdr.teststatusid = 1 and pdr.DeveloperNotificationStatus=1)))) "
        strStepFilter = ""
        if blnShowCompleteStep then
            strStepFilter = " location like '%Workflow Complete%' "
        end if
        if blnShowReleaseStep and strStepFilter = "" then
            strStepFilter = " location like '%Core Team%' "
        elseif blnShowReleaseStep then
            strStepFilter = strStepFilter & " or location like '%Core Team%' "
        end if
        if blnShowTestStep and strStepFilter = "" then
            strStepFilter = " (location like '%Engineering%' or location like '%Eng. Dev%') "
        elseif blnShowReleaseStep then
            strStepFilter = strStepFilter & " or (location like '%Engineering%' or location like '%Eng. Dev%') "
        end if
        if strStepFilter <> "" then
            strFilter = strFilter & " and ( " & strStepFilter & " ) "
        end if
	else
        strFilter = strFilter & " and ((isnull(pv.FusionRequirements, 0) = 0 and (pd.teststatusid <> 1 or ( pd.teststatusid = 1 and pd.DeveloperNotificationStatus=1 and v.location like '%Workflow Complete%'))) or (isnull(pv.FusionRequirements, 0) = 1 and (pdr.teststatusid <> 1 or ( pdr.teststatusid = 1 and pdr.DeveloperNotificationStatus=1 and v.location like '%Workflow Complete%')))) "	    
    end if
'	strFilter = strFilter & " and (v.location like '%Workflow Complete%') "
	
	if request("HighlightRow") <> "" then
		if strFilter = "" then
			strFilter = strFilter & " and v.id in ( " & scrubsql(request("HighlightRow")) & " ) "
		else
			strFilter = " and (v.id in ( " & scrubsql(request("HighlightRow")) & " ) or ( " 	& mid(strFilter,5) & " )) "
		end if
	end if

    'filter the releases selected for the pulsar products
    if strProductReleaseIDs <> "" then
        strFilter = strFilter & " and ((isnull(pv.FusionRequirements,0) = 0) or (isnull(pv.FusionRequirements,0) = 1 and pvr.ID in ( " & scrubsql(strProductReleaseIDs) & ")))" 
    end if

	'Build Base SQL
	if Request("ReportFormat") = "4"  then 'Accessory Report
		if Request("ReportFormat") = "6" then
	    	ColumnCount = 15
		else
	    	ColumnCount = 14
		end if
		strSQLSelect = "Select  v.location, v.sampledate, case when pdr.ID is null then pd.TargetNotes else pdr.TargetNotes end as targetnotes, case when pdr.ID is null then pd.riskrelease else pdr.riskrelease end as riskrelease, gs.MatrixBGColor as GreenSpecBGColor, " & _
                       "case when pdr.ID is null then a.MatrixBGColor else aRelease.MatrixBGColor end as AccessoryBGColor, c.commodity, " & _
                       "case when pdr.ID is null then t.MatrixBGColor else tsRelease.MatrixBGColor end as MatrixBGColor, v.id as DeliverableVersionID, v.Serviceactive, " & _
                       "v.active, v.serviceeoadate, v.endoflifedate as EOLDate, r.id as RootID, c.ID as CategoryID, v.assemblycode, cv.suppliercode as suppliercode, v.leadfree, v.rohsid, v.greenspecid, " & _
                       "c.name as category, r.name as DeliverableName, case when pdr.ID is null then pd.testconfidence else pdr.testconfidence end as testconfidence, pd.productversionid, '' as DCRSummary, '' as SubAssembly, '' as subassemblySpin, '' as subassemblyBase, " & _
                       "case when pdr.ID is null then a.Name else aRelease.Name end as AccessoryStatus, case when pdr.ID is null then pd.PilotDate else pdr.PilotDate end as AccessoryDate,  " & _ 
                       "case when pdr.ID is null then t.status else tsRelease.Status end as TestStatus, case when pdr.ID is null then pd.TestDate else pdr.TestDate end as TestDate , vd.name as Vendor, v.version, v.revision,v.pass,v.partnumber,v.ModelNumber, pd.DCRID, v.endoflifedate, r.Name as VersionDeliverableName, gs.name as greenSpec, rh.name as RoHS, Feature.FeatureName as FeatureName,FusionRequirements = isnull(pv.FusionRequirements, 0),ReleaseID = isnull(pvr.ReleaseID,0) "
		strSQL =            "FROM dbo.DeliverableRoot AS r WITH (NOLOCK) " & _
                             " INNER JOIN      dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID  " & _
                             " INNER JOIN      dbo.Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID" & _
                             " INNER JOIN      dbo.ProductVersion AS pv WITH (NOLOCK) ON pv.ID = pd.ProductVersionID  " & _
                             " INNER JOIN      dbo.ProductStatus AS ps WITH (NOLOCK) ON pv.ProductStatusID = ps.ID " & _
                             " INNER JOIN      dbo.AccessoryStatus AS a WITH (NOLOCK) ON pd.AccessoryStatusID = a.ID  " & _
                             " INNER JOIN      dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID  " & _
                             " INNER JOIN      dbo.ProductFamily AS f WITH (NOLOCK) ON f.ID = pv.ProductFamilyID " & _
                             " INNER JOIN      dbo.GreenSpec AS gs WITH (NOLOCK)  ON gs.ID = v.GreenSpecID" & _
                             " INNER JOIN      dbo.RoHS AS rh WITH (NOLOCK) ON v.RoHSID = rh.ID  " & _
                             " INNER JOIN dbo.Vendor AS vd WITH (NOLOCK)  on vd.ID = v.VendorID " & _
                             " LEFT OUTER JOIN dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID " & _
                             " LEFT OUTER JOIN dbo.DeliverableCategory_Vendor AS cv WITH (NOLOCK)  ON c.ID = cv.DeliverableCategoryID AND  vd.ID = cv.VendorID" & _
                             " LEFT OUTER JOIN dbo.Feature_Root AS FR WITH (NOLOCK) ON r.id = FR.ComponentRootID " & _
                             " LEFT OUTER JOIN dbo.Feature WITH (NOLOCK) ON Feature.FeatureID = FR.FeatureID " & _
                             " LEFT OUTER JOIN dbo.Product_Deliverable_Release pdr WITH(NOLOCK) ON pd.ID = pdr.ProductDeliverableID " & _
                             " LEFT OUTER JOIN dbo.ProductVersion_Release pvr WITH(NOLOCK) ON pvr.ReleaseID = pdr.ReleaseID and pvr.ProductVersionID = pd.ProductVersionID " & _
                             " LEFT OUTER JOIN dbo.ProductVersionRelease pvrelease WITH(NOLOCK) ON pvrelease.ID = pdr.ReleaseID " & _
                             " LEFT OUTER JOIN dbo.AccessoryStatus aRelease WITH(NOLOCK) ON pdr.AccessoryStatusID = aRelease.ID " & _
                             " LEFT OUTER JOIN dbo.TestStatus tsRelease WITH(NOLOCK) ON pdr.TestStatusID = tsRelease.ID " & _
                             " WHERE pv.typeid in (1,3)  " & _
                                " AND c.accessory=1  " & _
                                " AND r.kitnumber<> ''  " & _
                                " AND r.kitnumber is not null " & _
                                " AND r.rootfilename <> 'HFCN' "

		if request("chkChangeType") = "" then
				strSQl = strSQl & " and pd.accessorystatusid <> 0 " 
		end if				
	elseif Request("ReportFormat") <> "2" and request("ReportFormat")<> "5"  then 'Pilot or Qual Matrix
		if Request("ReportFormat") = "6" then
    		ColumnCount = 15
	    else
    		ColumnCount = 14
	    end if
		strSQLSelect = "Select  v.location,v.sampledate, case when pdr.ID is null then pd.TargetNotes else pdr.TargetNotes end as targetnotes, case when pdr.ID is null then pd.riskrelease else pdr.riskrelease end as riskrelease, gs.MatrixBGColor as GreenSpecBGColor, ui.FullName as ComponentPM, " & _
                       "case when pdr.ID is null then p.MatrixBGColor else pRelease.MatrixBGColor end as PilotBGColor, case when pdr.ID is null then t.MatrixBGColor else tsRelease.MatrixBGColor end as MatrixBGColor, v.id as DeliverableVersionID " & _
                       ", v.Serviceactive, v.active, v.serviceeoadate, v.endoflifedate as EOLDate, r.id as RootID, c.ID as CategoryID, v.assemblycode, cv.suppliercode as suppliercode, v.leadfree, " & _
                       "v.greenspecid, v.rohsid, c.name as category, r.name as DeliverableName, case when pdr.ID is null then pd.testconfidence else pdr.testconfidence end as testconfidence, pd.productversionid, '' as DCRSummary, '' as SubAssembly, '' as subassemblySpin,'' as subassemblyBase,  " & _
                       "case when pdr.ID is null then p.Name else pRelease.Name end as PilotStatus, case when pdr.ID is null then pd.PilotDate else pdr.PilotDate end as PilotDate, " & _
                       "case when pdr.ID is null then t.status else tsRelease.Status end as TestStatus, case when pdr.ID is null then pd.TestDate else pdr.TestDate end as TestDate, vd.name as Vendor, v.version, v.revision,v.pass,v.partnumber,v.ModelNumber, case when pdr.ID is null then pd.DCRID else pdr.DCRID end as DCRID, v.endoflifedate, r.Name as VersionDeliverableName, gs.name as greenSpec, rh.name as RoHS, Feature.FeatureName as FeatureName ,FusionRequirements = isnull(pv.FusionRequirements, 0),ReleaseID = isnull(pvr.ReleaseID,0) "
		strSQL = " FROM dbo.DeliverableRoot AS r WITH (NOLOCK) " & _
                      " INNER JOIN      dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID  " & _
	                  " INNER JOIN      dbo.Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID " & _
	                  " INNER JOIN      dbo.ProductVersion AS pv WITH (NOLOCK) ON pv.ID = pd.ProductVersionID " & _  
                      " INNER JOIN      dbo.Userinfo as ui WITH (NOLOCK) ON r.DevManagerID = ui.UserId  " & _
	                  " INNER JOIN      dbo.ProductStatus AS ps WITH (NOLOCK) ON pv.ProductStatusID = ps.ID " & _
		              " INNER JOIN      dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID   " & _
	                  " INNER JOIN      dbo.ProductFamily AS f WITH (NOLOCK) ON f.ID = pv.ProductFamilyID  " & _
		              " INNER JOIN      dbo.Vendor AS vd WITH (NOLOCK)  on vd.ID = v.VendorID  " & _
		              " INNER JOIN      dbo.GreenSpec AS gs WITH (NOLOCK)  ON gs.ID = v.GreenSpecID " & _
                      " INNER JOIN      dbo.RoHS AS rh WITH (NOLOCK) ON v.RoHSID = rh.ID   " & _
		              " LEFT OUTER JOIN dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID  " & _
                      " LEFT OUTER JOIN dbo.pilotstatus p with (NOLOCK) ON  pd.pilotstatusid =p.id  " & _
	                  " LEFT OUTER JOIN dbo.DeliverableCategory_Vendor AS cv WITH (NOLOCK)  ON c.ID = cv.DeliverableCategoryID AND  vd.ID = cv.VendorID " & _
                      " LEFT OUTER JOIN dbo.Feature_Root AS FR WITH (NOLOCK) ON r.id = FR.ComponentRootID " & _
                      " LEFT OUTER JOIN dbo.Feature WITH (NOLOCK) ON Feature.FeatureID = FR.FeatureID " & _
                      " LEFT OUTER JOIN dbo.Product_Deliverable_Release pdr WITH(NOLOCK) ON pd.ID = pdr.ProductDeliverableID " & _
                      " LEFT OUTER JOIN dbo.ProductVersion_Release pvr WITH(NOLOCK) ON pvr.ReleaseID = pdr.ReleaseID and pvr.ProductVersionID = pd.ProductVersionID " & _
                      " LEFT OUTER JOIN dbo.ProductVersionRelease pvrelease WITH(NOLOCK) ON pvrelease.ID = pdr.ReleaseID " & _
                      " LEFT OUTER JOIN dbo.TestStatus tsRelease WITH(NOLOCK) ON pdr.TestStatusID = tsRelease.ID " & _
                      " LEFT OUTER JOIN dbo.PilotStatus pRelease WITH(NOLOCK) ON pdr.PilotStatusID = pRelease.ID " & _
                        " WHERE pv.typeid in (1,3)   " & _
                            " AND r.rootfilename <> 'HFCN'  "             		

		if request("FullReport") = "1" then
			strSQl = strSQL & " and c.commodity=1 and pv.typeid in (1) and pv.productstatusid<4 " 
		else
			strSQl = strSQL & " and r.typeid=1 " 
		end if
				
		if request("chkChangeType") = "" then
				strSQl = strSQl & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.teststatusid <> 0) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.TestStatusid <> 0)) " 
		end if				
	elseif request("ReportFormat") = "5"  then 'Service
   		ColumnCount = 14
		strSQLSelect = "Select v.location, v.sampledate, case when pdr.ID is null then pd.TargetNotes else pdr.TargetNotes end as targetnotes, case when pdr.ID is null then pd.riskrelease else pdr.riskrelease end as riskrelease, gs.MatrixBGColor as GreenSpecBGColor, p.MatrixBGColor as PilotBGColor,   " & _
                       "case when pdr.ID is null then t.MatrixBGColor else tsRelease.MatrixBGColor end as MatrixBGColor, v.id as DeliverableVersionID,v.serviceactive, v.active, v.serviceeoadate, v.endoflifedate as EOLDate, r.id as RootID, c.ID as CategoryID,   " & _
                       "v.assemblycode, cv.suppliercode as suppliercode, v.leadfree, v.rohsid, v.greenspecid, c.name as category, r.name as DeliverableName,   " & _
                       "case when pdr.ID is null then pd.testconfidence else pdr.testconfidence end as testconfidence, pd.productversionid, '' as DCRSummary, case when prr.ID is null then coalesce(pr.servicesubassembly,pr.subassembly) else coalesce(prr.servicesubassembly,prr.subassembly) end as Subassembly,   " & _
                       "coalesce(pr.servicespin,pr.spin) as subassemblySpin, case when prr.ID is null then coalesce(pr.servicebase,pr.Base) else coalesce(prr.servicebase,prr.Base) end as subassemblyBase, p.Name as PilotStatus, case when pdr.ID is null then pd.PilotDate else pdr.PilotDate end as PilotDate, " & _
                       "case when pdr.ID is null then t.status else tsRelease.Status end as TestStatus, case when pdr.ID is null then pd.TestDate else pdr.TestDate end as TestDate, " & _
                       "vd.name as Vendor, v.version, v.revision,v.pass,v.partnumber,v.ModelNumber, case when pdr.ID is null then pd.DCRID else pdr.DCRID end as DCRID, v.endoflifedate, r2.Name as VersionDeliverableName, r2.ID as NativeSubassemblyRootID, " & _
                       "gs.name as greenSpec, rh.name as Rohs, Feature.FeatureName as FeatureName,FusionRequirements = isnull(pv.FusionRequirements, 0),ReleaseID = isnull(pvr.ReleaseID,0) "
		strSQL = " FROM	dbo.DeliverableCategory AS c WITH (NOLOCK)  " & _
			        " INNER JOIN dbo.DeliverableRoot AS r WITH (NOLOCK) ON c.ID = r.CategoryID  " & _
			        " INNER JOIN dbo.ProdDel_DelRoot AS pddr WITH (NOLOCK) ON r.ID = pddr.DeliverableRootID   " & _
			        " INNER JOIN dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.ID = pddr.ProductDeliverableID  " & _
			        " INNER JOIN dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID  " & _
                    " INNER JOIN dbo.Userinfo as ui WITH (NOLOCK) ON r.DevManagerID = ui.UserId  " & _
			        " INNER JOIN dbo.ProductFamily AS f WITH (NOLOCK) ON pv.ProductFamilyID = f.ID  " & _
			        " INNER JOIN dbo.ProductStatus AS ps WITH (NOLOCK) ON pv.ProductStatusID = ps.ID  " & _
			        " INNER JOIN dbo.Product_DelRoot AS pr WITH (NOLOCK) ON r.ID = pr.DeliverableRootID AND pv.ID = pr.ProductVersionID " & _
			        " INNER JOIN dbo.DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID   " & _
			        " INNER JOIN dbo.DeliverableRoot AS r2 WITH (NOLOCK) ON v.DeliverableRootID = r2.ID  " & _
			        " INNER JOIN dbo.GreenSpec AS gs WITH (NOLOCK) ON v.GreenSpecID = gs.ID " & _
			        " INNER JOIN dbo.RoHS AS rh WITH (NOLOCK) ON v.RoHSID = rh.ID   " & _
			        " INNER JOIN dbo.Vendor AS vd WITH (NOLOCK) on vd.ID = v.VendorID  " & _
			        " LEFT OUTER JOIN dbo.PilotStatus AS p WITH (NOLOCK) ON p.ID = pd.PilotStatusID  " & _
			        " LEFT OUTER JOIN dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pd.TestStatusID " & _
			        " LEFT OUTER JOIN dbo.DeliverableCategory_Vendor AS cv WITH (NOLOCK) ON c.ID = cv.DeliverableCategoryID AND  vd.ID = cv.VendorID  " & _
                    " LEFT OUTER JOIN dbo.Feature_Root AS FR WITH (NOLOCK) ON r.id = FR.ComponentRootID " & _
                    " LEFT OUTER JOIN dbo.Feature WITH (NOLOCK) ON Feature.FeatureID = FR.FeatureID " & _
                    " LEFT OUTER JOIN dbo.Product_DelRoot_Release prr WITH(NOLOCK) ON pr.ID = prr.ProductDelRootID " & _
                    " LEFT OUTER JOIN dbo.Product_Deliverable_Release pdr WITH(NOLOCK) ON pd.ID = pdr.ProductDeliverableID " & _
                    " LEFT OUTER JOIN dbo.ProductVersion_Release pvr WITH(NOLOCK) ON pvr.ReleaseID = pdr.ReleaseID and pvr.ProductVersionID = pd.ProductVersionID " & _
                    " LEFT OUTER JOIN dbo.ProductVersionRelease pvrelease WITH(NOLOCK) ON pvrelease.ID = pdr.ReleaseID " & _
                    " LEFT OUTER JOIN dbo.TestStatus tsRelease WITH(NOLOCK) ON pdr.TestStatusID = tsRelease.ID " & _
                    " LEFT OUTER JOIN dbo.PilotStatus pRelease WITH(NOLOCK) ON pdr.PilotStatusID = pRelease.ID " & _
                " WHERE	( r.RootFilename <> 'HFCN' )  " & _
		            " AND (r.TypeID = 1) " & _
		            " AND ((prr.ID is null AND (pr.ServiceSubassembly <> '' or (pr.ServiceSubassembly is null and pr.Subassembly <> '') ) AND (pr.ServiceSubassembly is not null or (pr.ServiceSubassembly is null and pr.Subassembly is not null))) " & _
		            " OR (prr.ID is not null AND (prr.ServiceSubassembly <> '' or (prr.ServiceSubassembly is null and prr.Subassembly <> '') ) AND (prr.ServiceSubassembly is not null or (prr.ServiceSubassembly is null and prr.Subassembly is not null))))" 

		if request("FullReport") = "1" then
			strSQl = strSQL & " and c.commodity=1 and pv.typeid in (1) and pv.productstatusid<4 " 
		else
			strSQl = strSQL & " and r.typeid=1 " 
		end if
		if request("chkChangeType") = "" then
				strSQl = strSQl & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.teststatusid <> 0) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.TestStatusid <> 0)) " 
		end if				
	else 'Subassembly Report
		if Request("ReportFormat") = "6" then
    		ColumnCount = 15
	    else
    		ColumnCount = 14
	    end if
		strSQLSelect = "Select v.location, v.sampledate, case when pdr.ID is null then pd.TargetNotes else pdr.TargetNotes end as targetnotes, case when pdr.ID is null then pd.riskrelease else pdr.riskrelease end as riskrelease, " & _
                        "gs.MatrixBGColor as GreenSpecBGColor, p.MatrixBGColor as PilotBGColor, " & _
                        "case when pdr.ID is null then t.MatrixBGColor else tsRelease.MatrixBGColor end as MatrixBGColor, v.id as DeliverableVersionID,v.serviceactive, v.active, " & _
                        "v.serviceeoadate, v.endoflifedate as EOLDate, r.id as RootID, c.ID as CategoryID," & _
                        " v.assemblycode, cv.suppliercode as suppliercode, v.leadfree,v.rohsid," & _
                        "v.greenspecid, c.name as category, r.name as DeliverableName, case when pdr.ID is null then pd.testconfidence else pdr.testconfidence end as testconfidence," & _
                        " pd.productversionid, '' as DCRSummary, case when prr.ID is null then pr.subassembly else prr.subassembly end as subassembly, case when prr.ID is null then pr.spin else prr.spin end as subassemblySpin," & _
                        " case when prr.ID is null then pr.Base else prr.Base end as subassemblyBase, case when pdr.ID is null then p.Name else p2.Name end as PilotStatus, case when pdr.ID is null then pd.PilotDate else pdr.PilotDate end as PilotDate," & _
                        " case when pdr.ID is null then t.status else tsRelease.Status end as TestStatus, case when pdr.ID is null then pd.TestDate else pdr.TestDate end as TestDate, vd.name as Vendor, v.version, v.revision," & _
                        " v.pass,v.partnumber,v.ModelNumber, case when pdr.ID is null then pd.DCRID else pdr.DCRID end as DCRID, v.endoflifedate," & _
                        " r2.Name as VersionDeliverableName, r2.ID as NativeSubassemblyRootID," & _
                        " gs.name as greenSpec, rh.name as Rohs,FusionRequirements = isnull(pv.FusionRequirements, 0),ReleaseID = isnull(pvr.ReleaseID,0)," & _
                        "  FeatureName= isnull((SELECT STUFF((select '; ' + isnull(avdetail.AvNo, '') + ' / ' +  avdetail.GPGDescription + ' / ' + feature.FeatureName " & _
									 " from Feature_Root  FR WITH (NOLOCK) " & _
									"		inner join feature WITH (NOLOCK) on Feature.FeatureID = FR.FeatureID " & _
									"		inner join avdetail WITH (NOLOCK) on avdetail.featureID=feature.featureID " & _
									"		inner join AvDetail_ProductBrand apb WITH (NOLOCK) on avdetail.avdetailID= apb.AvDetailID " & _
									"		inner join Product_Brand with (nolock) on apb.ProductBrandID=Product_Brand.ID " & _
									"  WHERE  FR.ComponentRootID = r.id    " & _
									"		and Product_Brand.ProductVersionID =pv.ID" & _
									"		for xml path('') ), 1, 2, '')), 'Not linked') " 
		
    
   
        strSQL = " FROM	dbo.DeliverableCategory AS c WITH (NOLOCK)  " & _
			        " INNER JOIN dbo.DeliverableRoot AS r WITH (NOLOCK) ON c.ID = r.CategoryID  " & _
			        " INNER JOIN dbo.ProdDel_DelRoot AS pddr WITH (NOLOCK) ON r.ID = pddr.DeliverableRootID   " & _
			        " INNER JOIN dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.ID = pddr.ProductDeliverableID  " & _
			        " INNER JOIN dbo.DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID  " & _ 
                    " INNER JOIN dbo.Userinfo as ui WITH (NOLOCK) ON r.DevManagerID = ui.UserId  " & _
			        " INNER JOIN dbo.Vendor AS vd WITH (NOLOCK) on vd.ID = v.VendorID  " & _
			        " INNER JOIN dbo.DeliverableRoot AS r2 WITH (NOLOCK) ON v.DeliverableRootID = r2.ID  " & _
			        " INNER JOIN dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _ 
			        " INNER JOIN dbo.ProductStatus AS ps WITH (NOLOCK) ON pv.ProductStatusID = ps.ID  " & _
			        " INNER JOIN dbo.ProductFamily AS f WITH (NOLOCK) ON pv.ProductFamilyID = f.ID  " & _
			        " INNER JOIN dbo.Product_DelRoot AS pr WITH (NOLOCK) ON r.ID = pr.DeliverableRootID AND pv.ID = pr.ProductVersionID " & _
			        " INNER JOIN dbo.GreenSpec AS gs WITH (NOLOCK) ON v.GreenSpecID = gs.ID " & _
			        " INNER JOIN dbo.RoHS AS rh WITH (NOLOCK) ON v.RoHSID = rh.ID   " & _
			        " LEFT OUTER JOIN dbo.PilotStatus AS p WITH (NOLOCK) ON p.ID = pd.PilotStatusID " & _
			        " LEFT OUTER JOIN dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pd.TestStatusID " & _
			        " LEFT OUTER JOIN dbo.DeliverableCategory_Vendor AS cv WITH (NOLOCK) ON c.ID = cv.DeliverableCategoryID AND  vd.ID = cv.VendorID  " & _
                    " LEFT OUTER JOIN dbo.Product_DelRoot_Release prr WITH(NOLOCK) ON pr.ID = prr.ProductDelRootID " & _
                    " LEFT OUTER JOIN dbo.Product_Deliverable_Release pdr WITH(NOLOCK) ON pd.ID = pdr.ProductDeliverableID " & _
                    " LEFT OUTER JOIN dbo.PilotStatus AS p2 WITH (NOLOCK) ON p2.ID = pdr.PilotStatusID " &_
                    " LEFT OUTER JOIN dbo.ProductVersion_Release pvr WITH(NOLOCK) ON pvr.ReleaseID = pdr.ReleaseID and pvr.ProductVersionID = pd.ProductVersionID " & _
                    " LEFT OUTER JOIN dbo.ProductVersionRelease pvrelease WITH(NOLOCK) ON pvrelease.ID = pdr.ReleaseID " & _
                    " LEFT OUTER JOIN dbo.TestStatus tsRelease WITH(NOLOCK) ON pdr.TestStatusID = tsRelease.ID " & _
                    "  WHERE r.RootFilename <> 'HFCN'  " & _
		               " AND ((prr.ID is null AND pr.Subassembly <> '' AND pr.Subassembly is not null) " & _
		               " OR (prr.ID is not null AND prr.Subassembly <> '' AND prr.Subassembly is not null))  "

		if request("FullReport") = "1" then
			strSQl = strSQL & " and c.commodity=1 and pv.typeid in (1) and pv.productstatusid<4 " 
		else
			strSQl = strSQL & " and r.typeid=1 " 
		end if
		if request("chkChangeType") = "" then
				strSQl = strSQl & " and ((isnull(pv.FusionRequirements, 0) = 0 and pd.teststatusid <> 0) or (isnull(pv.FusionRequirements, 0) = 1 and pdr.TestStatusid <> 0)) " 
		end if				
	end if
		
    strSQl = strSQL & strFilter & strHistoryFilter


    

	if not (strQueryString = "" or strQueryString = "ReportFormat=2" or strQueryString = "ReportFormat=3" or strQueryString = "ReportFormat=4" or strQueryString = "ReportFormat=1" or strQueryString = "ReportFormat=5") then 'No filters selected except possible reportformat
			'Get Product Buckets        
        rs.Open strSQLSelectProdList & " " & strSQL & " order by ProductName, pvrelease.ReleaseYear desc, pvrelease.ReleaseMonth desc",cn,adOpenForwardOnly
		ProductList = ""
		ProductIDList = ""
        ProductReleaseIDList = ""
		i=0
		do while not rs.EOF
			ProductList = ProductList & "," & rs("DOTSName")
			ProductIDList = ProductIDList & "," & rs("ID")
            ProductReleaseIDList = ProductReleaseIDList & "," & rs("ID") & ";" & rs("ReleaseID")
            i=i+1
			rs.MoveNext
		loop
		rs.Close
	end if
	
	dim strPVDate
	dim strSI1Date
	dim strSI2Date
	dim strDateOffset
	dim strMilestoneOffsetBGColor
	
    strSI1Date = ""
    strSI2Date = ""
    strPVDate = ""
    strMilestoneOffsetBGColor = ""
    
	if i = 1 then
	    rs.Open "spGetProductMilestoneDate " & clng(mid(ProductIDList,2)) & ", 18"
	    if not (rs.EOF and rs.BOF) then
	        strSI1Date = rs("MilestoneDate")
	    end if
	    rs.Close
	    rs.Open "spGetProductMilestoneDate " & clng(mid(ProductIDList,2)) & ", 25"
	    if not (rs.EOF and rs.BOF) then
	        strSI2Date = rs("MilestoneDate")
	    end if
	    rs.Close
	    rs.Open "spGetProductMilestoneDate " & clng(mid(ProductIDList,2)) & ", 32"
	    if not (rs.EOF and rs.BOF) then
	        strPVDate = rs("MilestoneDate")
	    end if
	    rs.Close
	end if	
%>

	<table width=100% border=0>
		<tr><TD valign=top colspan=6>
			<font size=3 face=verdana><b>
			<%
			if instr(mid(ProductList,2),",")=0 then
				Response.Write mid(ProductList,2) 
			end if
			if request("txtTitle") = "" then%>
				<%if Request("ReportFormat")="6" and i = 1 then%>
					Samples Availability Report
				<%elseif Request("ReportFormat")="5" then%>
					Service Report
				<%elseif Request("ReportFormat")="4" then%>
					Accessory Report
				<%elseif Request("ReportFormat")="3" then%>
					Pilot Report
				<%elseif Request("ReportFormat")="2" then%>
					Hardware SubAssembly Report
				<%else%>
					Hardware Qualification Report
				<%end if%>
			<%else%>
				<%=request("txtTitle")%>
			<%end if%></b>
            <%=strReportDateRange%>
			</font>		
		<% 	    			

        if (trim(strSI1Date) <> "" or trim(strSI2Date) <> "" or trim(strPVDate) <> "") and trim(request("ReportFormat"))="6" then
            response.Write "<BR><BR><Table><tr>"
            if trim(strSI1Date) <> "" then
                response.Write "<TD><b>SI1 Start Date:</b></TD><td>" & strSI1Date & "&nbsp;&nbsp;&nbsp;</td>"
            end if
            if trim(strSI2Date) <> "" then
                response.Write "<TD><b>SI2 Start Date:</b></TD><td>" & strSI2Date & "&nbsp;&nbsp;&nbsp;</td>"
            end if
            if trim(strPVDate) <> "" then
                response.Write "<TD><b>PV Start Date:</b></TD><td>" & strPVDate & "</td>"
            end if
            response.Write "</tr></table>"
        end if
        %>
		<td valign=top style="display:<%=ShowQuickLinks%>" ID=QuickLinks align=right>
			<%if request("FileType") = "" then%>
				<table><tr><td><font size=1 face=verdana>	Export: <a href="javascript: Export(1);">Excel</a><%=strExportLink%></td></tr>
				<%''***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
                    'if currentuserid = 30 then
                %>
					<!--//<tr><td><font size=1 face=verdana>Display:<u LANGUAGE="javascript" onclick="Test(1)">Quick&nbsp;Reports</u><%=strExportLink%></td></tr>//-->
				<%'end if%>
				<%if blnAllowQuickFilters then%>
						<td><font size=1 face=verdana>	Report: <a href="javascript: SwitchFilterView(1);">Filters</a></td></tr>
				<%else%>
						<td><font size=1 face=verdana>	Report: <a href="HardwareMatrix.asp">Filters</a></td></tr>
				<% end if%>
				</table>
			<%end if%>
		</td></tr></table>
		
		
	<%

	if strQueryString = "" or strQueryString = "ReportFormat=2" or strQueryString = "ReportFormat=5" or strQueryString = "ReportFormat=3" or strQueryString = "ReportFormat=4" or strQueryString = "ReportFormat=1" then 'No filters selected except possible reportformat
		if blnAllowQuickFilters then
			DrawFilterBox CurrentUserID,ShowFilters
		end if
		Response.Write "<font size=2 face=verdana><b>Please select a filter to continue.</b></font>"
		Response.Write "<BR><BR><BR>"
		
	else
		dim strExportLink

		if blnAllowQuickFilters then
			DrawFilterBox CurrentUserID,ShowFilters
		end if	
		ProductList = mid(ProductList,2)
		ProductIDList = mid(ProductIDList,2)
		ProductArray = split(ProductList,",")
		ProductIDArray = split(ProductIDList,",")
		ProductBuckets = split(ProductIDList,",")
		ProductSubArray = split(ProductIDList,",")
        ProductReleaseIDList = mid(ProductReleaseIDList,2)
        ProductReleaseIDArray = split(ProductReleaseIDList, ",")
        
        dim strProductReleaseID
		
		ProductCount = ubound(ProductArray) + 1
		if ProductCount = 1 then
			DCRCellCount = 1
			if request("ReportFormat") = "3" or request("ReportFormat") = "4" then
				StatusCellCount = 1
			else
				StatusCellCount = 0
			end if
		else
			DCRCellCount = 0
			StatusCellCount = 0
		end if
		
		
		
		ColumnCount = DCRCellCount + ColumnCount + ubound(ProductArray) + StatusCellCount
		
		 if ubound(ProductIDArray) = 0 then
			columncount=columncount +1 
		 end if
       	
        if Request("ReportFormat") = "2" then
			rs.Open	strSQLSelect & " " & strSQL & " order by c.name, case when prr.ID is null then pr.Base else prr.Base end, r.id, vd.name, v.id;", cn,adOpenForwardOnly
        elseif Request("ReportFormat") = "5" then
			rs.Open	strSQLSelect & " " & strSQL & " order by c.name, case when prr.ID is null then pr.Base else prr.Base end, vd.name, v.id;", cn,adOpenForwardOnly
		else
			rs.Open	strSQLSelect & " " & strSQL & " order by c.name, r.name, vd.name, v.id;", cn,adOpenForwardOnly
		end if
		LastCategory = ""
		LastRoot=""
		LastVersion = ""
		Response.Write "<TABLE bgcolor=Ivory width=100% ID=MyTable border=1 cellspacing=0 cellpadding=2 bordercolor=gainsboro>"
		counter = 1
		rowCounter = 0
		ProductLoop = 0     
		if ubound(ProductIDArray) = 0 then
			Response.Write "<TR class=HeaderRow>" 
			Response.Write "<TH>ID</TH>"
			if Request("ReportFormat")="2" or Request("ReportFormat") = "5" then
				Response.Write "<TH>Bridged</TH>"
			end if
			Response.Write "<TH>Supplier</TH>"
			if Request("ReportFormat")<>"2" and Request("ReportFormat")<>"4" and Request("ReportFormat")<>"5" then
				Response.Write "<TH>Component&nbsp;PM</TH>"
				Response.Write "<TH>Release&nbsp;Notes</TH>"
            end if
			Response.Write "<TH>Model/Vendor&nbsp;Part&nbsp;No.</TH>"
			Response.Write "<TH>HW</TH>"
			Response.Write "<TH>FW</TH>"
			Response.Write "<TH>Rev</TH>"
			Response.Write "<TH>RoHS/Green&nbsp;Spec</TH>"
			Response.Write "<TH>Samples&nbsp;</TH>"
			if Request("ReportFormat") = "6" and  productcount=1 then
                Response.Write "<TH>&nbsp;SI1&nbsp;Available</TH>"
                Response.Write "<TH>&nbsp;SI2&nbsp;Available</TH>"
    			Response.Write "<TH>&nbsp;PV&nbsp;Available</TH>"
            end if
            if Request("ReportFormat") = "5" then
				Response.Write "<TH>Service&nbsp;EOA</TH>"
			else
				Response.Write "<TH>Factory&nbsp;EOA</TH>"
			end if
			if not (Request("ReportFormat") = "6" and productCount=1) then
			    Response.Write "<TH>DCR/HFCN</TH>"
			    Response.Write "<TH>A.&nbsp;Code</TH>"
			end if
			Response.Write "<TH>HP&nbsp;Part&nbsp;No.</TH>"
            if request("ReportFormat") = "3"  then
				Response.Write "<TH nowrap>Qual&nbsp;Status</TH><TH nowrap>Pilot&nbsp;Status</TH>"
			elseif request("ReportFormat") = "4" then
				Response.Write "<TH nowrap>Qual&nbsp;Status</TH><TH nowrap>Accessory&nbsp;Status</TH>"
			else
				Response.Write "<TH nowrap>Qual&nbsp;Status</TH>"
			end if
			Response.Write "<TH>Comments</TH>"
			Response.Write "</TR>"
		end if

		

		do while not rs.EOF
			rowCounter = rowCounter + 1
			blnVersionChange = false        
			if LastVersion <> rs("DeliverableVersionID") or ((LastRoot <> rs("RootID") and (request("ReportFormat") <> "2" and request("ReportFormat") <> "5" )) or (trim(LastSub) <> trim(rs("subassemblyBase") & "")  and (request("ReportFormat") = "2" or request("ReportFormat") = "5"))) then 'LastRoot <> rs("RootID")
				ProductLoop = 0
				blnVersionChange = true
				if lastversion <> "" then                    
        			for i = lbound(ProductBuckets) to ubound(ProductBuckets)                                               
                        if ProductBuckets(i) = "&nbsp;" then
							Response.Write "<TD>" & ProductBuckets(i) & "</TD>"
						else                       
							Response.Write ProductBuckets(i)
						end if	
						ProductBuckets(i) = "&nbsp;"
					next
					Response.Write "</TR>" 			
				else
					for i = lbound(ProductBuckets) to ubound(ProductBuckets)
						ProductBuckets(i) = "&nbsp;"
					next
				end if
			end if
		    Response.Flush	

					   
	        if LastCategory <> rs("CategoryID") then
				Response.Write "<TR bgcolor=SeaGreen class=CatRow><TD colspan=" & ColumnCount & ">" & rs("Category") & "</TD></tr>"
				LastCategory= rs("CategoryID")
			end if
        
			if (LastRoot <> rs("RootID") and (request("ReportFormat") <> "2" and request("ReportFormat") <> "5" )) or (trim(LastSub) <> trim(rs("subassemblyBase") & "") and (request("ReportFormat") = "2" or request("ReportFormat") = "5") ) then
                strBaseSub = rs("subassemblyBase") & ""
        
                if strBaseSub <> "" then
					strBaseSub = strBaseSub & "-XXX"
				end if

		        if request("ReportFormat") = "2" or request("ReportFormat") = "5" then
                    if not (rs("CategoryID") = 227 and request("ReportFormat") = "2" ) then ''' not power cords in Subassembly Report
					    for i = lbound (ProductSubArray) to ubound (ProductSubArray)
						    ProductSubArray(i) = "&nbsp;&nbsp;"
					    next
					    set	rs2 = server.CreateObject("ADODB.recordset")
					
					    if request("ReportFormat") = "2" then
					        rs2.open "spListSubassembliesForBase '" & rs("subassemblyBase") & "'",cn,adOpenStatic
					    else
					        rs2.open "spListSubassembliesForBaseService '" & rs("subassemblyBase") & "'",cn,adOpenStatic
					    end if
					    do while not rs2.EOF
						    for i = lbound(ProductReleaseIDArray) to ubound(ProductReleaseIDArray)
                                strProductReleaseID = Split(ProductReleaseIDArray(i),";")
							    if (trim(strProductReleaseID(0)) = trim(rs2("ProductVersionID") & "")) and (trim(strProductReleaseID(1)) = trim(rs2("ReleaseID"))) then                                    
								    if trim(rs2("SubassemblySpin") & "") = "" then
									    ProductSubArray(i) =  "&nbsp;TBD&nbsp;"
								    else
									    ProductSubArray(i) =  "&nbsp;" & rs2("SubassemblySpin") & "&nbsp;"
								    end if
								    exit for
							    end if
						    next
						    rs2.MoveNext
					    loop	
					    set rs2=nothing
					    if ubound (ProductSubArray) = 0 then 
						    Response.Write "<TR bgcolor=Burlywood class=RootRow><TD colspan=" & 11 + DCRCellCount + StatusCellCount & ">" & rs("subassemblyBase")  & " [" & rs("DeliverableName") & "]" & " (" & rs("FeatureName") & ")" & " </TD>"
						    Response.Write "<TD colspan=3 nowrap align=left>" & trim(left(strBaseSub,instr(strBaseSub,"-"))) & trim(mid(ProductSubArray(0),7)) & "</TD>"
					    else
						    if strBaseSub = "" then
							    strBaseSub= "TBD"
						    end if 
						    Response.Write "<TR bgcolor=Burlywood class=RootRow><TD colspan=" & 11 + DCRCellCount + StatusCellCount & ">" & rs("subassemblyBase")  & " [" & rs("DeliverableName") & "]" & " (" & rs("FeatureName") & ")" & "</TD><TD align=center>" & strBaseSub & "</TD>"
						    for i = lbound (ProductSubArray) to ubound (ProductSubArray)
							    Response.Write "<TD align=center>" & ProductSubArray(i) & "</TD>"
						    next
					    end if
					    Response.Write "</tr>" 
                    end if
				else
					if request("ReportFormat") = "6" and productcount=1 then
					    extracolumncount=1
                    else
					    extracolumncount=0
					end if
        
                    dim strFeatureName
                    
                    if  (strProducts) = "" then
                        
                         if   request("lstFamily") <> "" and rs("FusionRequirements") ="False"  then
                                 strFeatureName  = ""
                         else
                                if isnull(rs("FeatureName")) then
                                    strFeatureName = "&nbsp;(Not Linked)&nbsp;" 
				                else
					                strFeatureName = "&nbsp;(&nbsp;" & rs("FeatureName") & "&nbsp;)&nbsp;" 
                                end if
                         end if
                       
                    else
                        if   rs("FusionRequirements") ="True" then
                             if isnull(rs("FeatureName")) then
                                strFeatureName = "&nbsp;(Not Linked)&nbsp;" 
				             else
					            strFeatureName = "&nbsp;(&nbsp;" & rs("FeatureName") & "&nbsp;)&nbsp;" 
                             end if

                         else
                            strFeatureName  = ""
                         end if
                    end if
                    
                    dim colspan
          		    if (request("lstFamily") <> "") then 
                        colspan = 15 
                    else 
                        colspan = 14 
                    end if
					Response.Write "<TR bgcolor=Burlywood class=RootRow><TD colspan=" & colspan + extracolumncount + StatusCellCount+ DCRCellCount + Ubound(ProductArray) & ">" & rs("RootID") & "&nbsp;-&nbsp;" & rs("DeliverableName") & strFeaturename & "</TD></tr>" 'LightSteelBlue
                  
				end if

				if ubound(ProductIDArray) > 0 and (not (rs("CategoryID") = 227 and request("ReportFormat") = "2" )) then
					Response.Write "<TR bgcolor=LightGoldenrodYellow class=HeaderRow>" 
					Response.Write "<TH>ID</TH>"
					if Request("ReportFormat")="2"  or Request("ReportFormat")="5" then
						Response.Write "<TH>Bridged</TH>"
					end if
					Response.Write "<TH>Supplier</TH>"
					if Request("ReportFormat")<>"2" and Request("ReportFormat")<>"4" and Request("ReportFormat")<>"5" then
					    Response.Write "<TH>Component&nbsp;PM</TH>"
					    Response.Write "<TH>Release&nbsp;Notes</TH>"
                    end if
			        Response.Write "<TH>Model/Vendor&nbsp;Part&nbsp;No.</TH>"
					Response.Write "<TH>HW</TH>"
					Response.Write "<TH>FW</TH>"
					Response.Write "<TH>Rev</TH>"
					Response.Write "<TH>RoHS/Green&nbsp;Spec</TH>"
					Response.Write "<TH>Samples&nbsp;</TH>"
        			if Request("ReportFormat") = "6" and productcount=1 then
					    Response.Write "<TH>&nbsp;SI1&nbsp;Available</TH>"
	    				Response.Write "<TH>&nbsp;SI2&nbsp;Available</TH>"
    					Response.Write "<TH>&nbsp;PV&nbsp;Available</TH>"
					end if
					if Request("ReportFormat")="5" then
						Response.Write "<TH>Service&nbsp;EOA</TH>"
					else
						Response.Write "<TH>Factory&nbsp;EOA</TH>"
					end if
					'Response.Write "<TH>DCR/HFCN</TH>"
        			if not(Request("ReportFormat") = "6" and productcount=1) then
		    			Response.Write "<TH>A.&nbsp;Code</TH>"
			        end if
			    	Response.Write "<TH>HP&nbsp;Part&nbsp;No.</TH>"
					for each strProduct in ProductArray					    	
	               		Response.Write "<TH nowrap>" & strProduct & "</TH>"
					next
					Response.Write "</TR>"
				end if
		        counter=counter+1
		     end if

            '''if "Power Cords" and "Subassambly Report"
            if LastRoot <> rs("RootID") and request("ReportFormat") = "2" and rs("CategoryID") = 227 then
				for i = lbound (ProductSubArray) to ubound (ProductSubArray)
					ProductSubArray(i) = "&nbsp;&nbsp;"
				next
				set	rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spListSubassembliesForRoot '" & rs("RootID") & "'",cn,adOpenStatic
				
				do while not rs2.EOF
					for i = lbound(ProductIDArray) to ubound(ProductIDArray)
						if trim(ProductIDArray(i)) = trim(rs2("ProductVersionID") & "") then
							if trim(rs2("SubassemblySpin") & "") = "" then
								ProductSubArray(i) =  "&nbsp;TBD&nbsp;"
							else
								ProductSubArray(i) =  "&nbsp;" & rs2("SubassemblySpin") & "&nbsp;"
							end if
							exit for
						end if
					next
					rs2.MoveNext
				loop	
				set rs2=nothing
				if ubound (ProductSubArray) = 0 then 
					Response.Write "<TR bgcolor=Burlywood class=RootRow><TD colspan=" & 11 + DCRCellCount + StatusCellCount & ">" & rs("subassemblyBase")  & " [" & rs("DeliverableName") & "]" & " (" & rs("FeatureName") & ")" & " </TD>"
					Response.Write "<TD colspan=3 nowrap align=left>" & rs("subassembly") & "</TD>"
				else
					if strBaseSub = "" then
						strBaseSub= "TBD"
					end if 
					Response.Write "<TR bgcolor=Burlywood class=RootRow><TD colspan=" & 11 + DCRCellCount + StatusCellCount & ">" & rs("subassemblyBase")  & " [" & rs("DeliverableName") & "]" & " (" & rs("FeatureName") & ")" & "</TD><TD align=center>" & strBaseSub & "</TD>"
					for i = lbound (ProductSubArray) to ubound (ProductSubArray)
						Response.Write "<TD align=center>" & ProductSubArray(i) & "</TD>"
					next
				end if
				Response.Write "</tr>" 

				if ubound(ProductIDArray) > 0 then
					Response.Write "<TR bgcolor=LightGoldenrodYellow class=HeaderRow>" 
					Response.Write "<TH>ID</TH>"
					if Request("ReportFormat")="2"  or Request("ReportFormat")="5" then
						Response.Write "<TH>Bridged</TH>"
					end if
					Response.Write "<TH>Supplier</TH>"
					if Request("ReportFormat")<>"2" and Request("ReportFormat")<>"4" and Request("ReportFormat")<>"5" then
					    Response.Write "<TH>Component&nbsp;PM</TH>"
					    Response.Write "<TH>Release&nbsp;Notes</TH>"
                    end if
			        Response.Write "<TH>Model/Vendor&nbsp;Part&nbsp;No.</TH>"
					Response.Write "<TH>HW</TH>"
					Response.Write "<TH>FW</TH>"
					Response.Write "<TH>Rev</TH>"
					Response.Write "<TH>RoHS/Green&nbsp;Spec</TH>"
					Response.Write "<TH>Samples&nbsp;</TH>"
        			if Request("ReportFormat") = "6" and productcount=1 then
					    Response.Write "<TH>&nbsp;SI1&nbsp;Available</TH>"
	    				Response.Write "<TH>&nbsp;SI2&nbsp;Available</TH>"
    					Response.Write "<TH>&nbsp;PV&nbsp;Available</TH>"
					end if
					if Request("ReportFormat")="5" then
						Response.Write "<TH>Service&nbsp;EOA</TH>"
					else
						Response.Write "<TH>Factory&nbsp;EOA</TH>"
					end if
					'Response.Write "<TH>DCR/HFCN</TH>"
        			if not(Request("ReportFormat") = "6" and productcount=1) then
		    			Response.Write "<TH>A.&nbsp;Code</TH>"
			        end if
			    	Response.Write "<TH>HP&nbsp;Part&nbsp;No.</TH>"
					for each strProduct in ProductArray			    	
	               		Response.Write "<TH nowrap>" & strProduct & "</TH>"
					next
					Response.Write "</TR>"
				end if

                counter=counter+1
            end if

			LastRoot = rs("RootID")
			LastSub = rs("subassemblyBase") & ""

		     if LastVersion <> rs("DeliverableVersionID") or blnVersionChange then
				LastVersion = rs("DeliverableVersionID")
				
				if rs("DCRID") > 2 then
					strDCR = "DCR:&nbsp;" & rs("DCRID")
				elseif rs("DCRID")=2 then
					strDCR = "HFCN"
				elseif rs("DCRID")=1 then
					strDCR = "POR"
				else
					strDCR = "&nbsp;"
				end if
				
				if Request("ReportFormat")="5" then
					if isnull(rs("ServiceEOADate")) and rs("ServiceActive") then
						EOLBGColor = ""
					elseif datediff("d",rs("ServiceEOADate"),Now)>=-120 and datediff("d",rs("ServiceEOADate"),Now)<0 then
						EOLBGColor="#ffff99"
					elseif DateDiff("d",rs("ServiceEOADate"),Now()) > 0 or (not rs("ServiceActive")) then
						EOLBGColor = "salmon"
					else
						EOLBGColor = ""
					end if
					
					if not rs("ServiceActive") then
						strEOLDate = "Unavailable"
					else
						strEOLDate = rs("ServiceEOADate")
					end if
				else
					if isnull(rs("EOLDate")) and rs("Active") then
					elseif datediff("d",rs("EndOfLifeDate"),Now)>=-120 and datediff("d",rs("EndOfLifeDate"),Now)<0 then
						EOLBGColor="#ffff99"
					elseif DateDiff("d",rs("EOLDate"),Now()) > 0 or (not rs("Active")) then
						EOLBGColor = "salmon"
					else
						EOLBGColor = ""
					end if
					if (not rs("Active")) and rs("ServiceActive") then
						strEOLDate = "Service&nbsp;Only"
					elseif (not rs("Active")) and (not rs("ServiceActive")) then
						strEOLDate = "Unavailable"
					else
						strEOLDate = rs("EndOfLifeDate")
					end if
				end if				
				'if request("FileType")="" then
				'	strDCRTitle = rs("DCRSummary") & ""
				'else
					strDCRTitle = ""
				'end if
				if  instr("," & replace(request("HighlightRow")," ","") & ",","," & trim(rs("Deliverableversionid")) & ",") > 0 then
					response.write "<TR bgcolor=Plum  ID=Row" & rowcounter & " class=Row LANGUAGE=javascript onmouseover=""return Commodity_onmouseover()"" onmouseout=""return Commodity_onmouseout()"" onclick=""return Commodity_onclick()"">"
					ScrollToRow = rowcounter
				else
					response.write "<TR ID=Row" & rowcounter & " class=Row LANGUAGE=javascript onmouseover=""return Commodity_onmouseover()"" onmouseout=""return Commodity_onmouseout()"" onclick=""return Commodity_onclick()"">"
				end if
'				response.write "<TD><a href=""javascript: DisplayVersion(" & rs("DeliverableVersionID") & ");"">" & rs("DeliverableVersionID") & "</a></TD>"
				if request("FileType")="" then
					response.write "<TD><a target=_blank href=""../Query/DeliverableVersionDetails.asp?ID=" & rs("DeliverableVersionID") & """>" & rs("DeliverableVersionID") & "</a></TD>"
				else
					response.write "<TD>" &  rs("DeliverableVersionID") &  "</TD>"
				end if
				strGreenSpecColor = trim(lcase(rs("GreenSpecBGColor") & ""))
				if Request("ReportFormat")="2" or Request("ReportFormat")="5" then
					
					if trim(rs("RootID")) = trim(rs("NativeSubassemblyRootID")) then
						response.write "<TD nowrap>&nbsp;</TD>"
					else
						
						if LastSubassemblyBase = rs("subassemblyBase") & "" and LastRootID = rs("NativeSubassemblyRootID") then
							response.write "<TD title=""" & rs("VersionDeliverableName") & """>"& LastNativeBase & "</TD>"
						else						
						
							set	rs2 = server.CreateObject("ADODB.recordset")
							rs2.open "spGetNativeSubAssemblyBase '" & rs("subassemblyBase") & "'," & rs("NativeSubassemblyRootID") & ", " & rs("ProductVersionID") & ", " & rs("ReleaseID"),cn,adOpenStatic
					
							if rs2.EOF and rs2.BOF then
								response.write "<TD title=""" & rs("VersionDeliverableName") & """>Unknown</TD>"				
							else
								response.write "<TD title=""" & rs("VersionDeliverableName") & """>"& rs2("NativeSubassembly") & "</TD>"
								LastSubassemblyBase = rs("subassemblyBase") & ""
								LastRootID = rs("NativeSubassemblyRootID")
								LastNativeBase = rs2("NativeSubassembly") & ""
							end if
							rs2.close
							set rs2=nothing
						end if												
					end if
				end if
				if trim(rs("SupplierCode") & "") = "" or trim(rs("SupplierCode") & "") = "TBD"  then
					response.write "<TD nowrap>" & rs("vendor") & "</TD>"
				else
					response.write "<TD nowrap>" & rs("vendor") & " (" & rs("supplierCode") & ")</TD>"
				end if
                if ((Request("ReportFormat")<>"2" and Request("ReportFormat")<>"4" and Request("ReportFormat")<>"5")) then      'Show following columns only if it is a qual or pilot run --added the Accessory report
                    response.write "<TD>" & rs("ComponentPM") & "</TD>"                         'Component PM column
				    response.write "<TD><a href=""javascript: ShowChanges(" & rs("DeliverableVersionID") & ");"">Notes</a></TD>"             'Notes Column
                end if

				response.write "<TD>" & server.htmlencode(rs("ModelNumber")& "") & "</TD>"
				response.write "<TD>" & rs("Version") & "</TD>"
				response.write "<TD>" & rs("Revision") & "&nbsp;</TD>"
				response.write "<TD align=center>&nbsp;" & rs("Pass") & "&nbsp;</TD>"
				if trim(rs("Rohs") & "") = "" and trim(rs("Greenspec") & "") = "" then
				    response.write "<TD bgcolor=""Salmon"" align=left nowrap>&nbsp;</td>"
				elseif trim(rs("Rohs") & "") <> "" and trim(rs("Greenspec") & "") <> "" then
				    response.write "<TD align=left nowrap>" & rs("Rohs") & "_" & rs("GreenSpec") & "</td>"
                elseif trim(rs("Rohs") & "") <> "" then
				    response.write "<TD bgcolor=""#ffff99"" align=left nowrap>" & rs("Rohs") & "</td>"
				else
				    response.write "<TD bgcolor=""#ffff99"" align=left nowrap>" & rs("GreenSpec") & "</td>"
				end if
				if trim(rs("SampleDate") & "") = "" then
				    response.write "<TD align=left nowrap>&nbsp;</td>"
				else
				    response.write "<TD align=left nowrap>" & rs("SampleDate") & "</td>"
				end if
    			if Request("ReportFormat") = "6" and productcount=1 then
    			    if trim(strSI1Date) = "" or (not isdate(strSI1Date)) then
		    		    response.write "<TD align=middle>N/A</TD>"
    			    elseif trim(rs("SampleDate") & "") = "" or (not isdate(rs("SampleDate"))) then
		    		    response.write "<TD align=middle bgcolor=salmon>Unknown</TD>"
		    		else
		    		    strDateOffset = datediff("d",strSI1Date,rs("SampleDate"))
		    		    if strDateOffset <= 0 then
		    		        if abs(strDateOffset) < 8 then
		    		            strMilestoneOffsetBGColor = "#ffff99"
		    		        else
		    		            strMilestoneOffsetBGColor = ""
		    		        end if
        				    response.write "<TD align=middle bgcolor=" & strMilestoneOffsetBGColor & ">" & abs(strDateOffset) & "&nbsp;days&nbsp;early&nbsp;</TD>"
        				else
        				    response.write "<TD align=middle bgcolor=salmon>" & abs(strDateOffset) & "&nbsp;days&nbsp;late&nbsp;</TD>"
        				end if
		    		end if
    			    if trim(strSI2Date) = "" or (not isdate(strSI2Date)) then
		    		    response.write "<TD align=middle>N/A</TD>"
    			    elseif trim(rs("SampleDate") & "") = "" or (not isdate(rs("SampleDate"))) then
		    		    response.write "<TD align=middle bgcolor=salmon>Unknown</TD>"
		    		else
		    		    strDateOffset = datediff("d",strSI2Date,rs("SampleDate"))
		    		    if strDateOffset <= 0 then
		    		        if abs(strDateOffset) < 8 then
		    		            strMilestoneOffsetBGColor = "#ffff99"
		    		        else
		    		            strMilestoneOffsetBGColor = ""
		    		        end if
        				    response.write "<TD align=middle bgcolor=" & strMilestoneOffsetBGColor & ">" & abs(strDateOffset) & "&nbsp;days&nbsp;early&nbsp;</TD>"
        				else
        				    response.write "<TD align=middle bgcolor=salmon>" & abs(strDateOffset) & "&nbsp;days&nbsp;late&nbsp;</TD>"
        				end if
		    		end if
    			    if trim(strPVDate) = "" or (not isdate(strPVDate)) then
		    		    response.write "<TD align=middle>N/A</TD>"
    			    elseif trim(rs("SampleDate") & "") = "" or (not isdate(rs("SampleDate"))) then
		    		    response.write "<TD align=middle bgcolor=salmon>Unknown</TD>"
		    		else
		    		    strDateOffset = datediff("d",strPVDate,rs("SampleDate"))
		    		    if strDateOffset <= 0 then
		    		        if abs(strDateOffset) < 8 then
		    		            strMilestoneOffsetBGColor = "#ffff99"
		    		        else
		    		            strMilestoneOffsetBGColor = ""
		    		        end if
        				    response.write "<TD align=middle bgcolor=" & strMilestoneOffsetBGColor & ">" & abs(strDateOffset) & "&nbsp;days&nbsp;early&nbsp;</TD>"
        				else
        				    response.write "<TD align=middle bgcolor=salmon>" & abs(strDateOffset) & "&nbsp;days&nbsp;late&nbsp;</TD>"
        				end if
		    		end if
	            end if			
				response.write "<TD align=center bgcolor=" & EOLBGColor & ">" & strEOLDate & "&nbsp;</TD>"
    			if not(Request("ReportFormat") = "6" and productcount=1) then
				    if DCRCellCount = 1 then
    					response.write "<TD title=""" & server.htmlencode(strDCRTitle) & """>" & strDCR & "</TD>"
				    end if
				    response.write "<TD>" & rs("AssemblyCode") & "&nbsp;</TD>"
                end if
				response.write "<TD nowrap>" & rs("PartNumber") & "&nbsp;</TD>"
			end if
		    
		    strTestStatus = ""
			strPilotStatus = ""
			strTestColor = ""
			strGreenSpecColor = ""
			strPilotColor = ""
	
			if request("ReportFormat") = "3" then 'Pilot Report
				if trim(rs("PilotStatus") & "") = "P_Scheduled" then
					strPilotStatus = rs("PilotDate") & ""
				else
					strPilotStatus = rs("PilotStatus")
				end if
				strPilotColor = lcase(trim(rs("PilotBGColor") & ""))
				if request("FileType") <> "" and strpilotcolor = "darkseagreen" then
					strPilotColor = "SeaGreen"
				elseif request("FileType") <> "" and strpilotcolor = "lightsteelblue" then
					strPilotColor = "LightSkyBlue"
				end if				
			elseif request("ReportFormat") = "4" then 'Accessory Report
				if trim(rs("AccessoryStatus") & "") = "Scheduled" then
					strAccessoryStatus = rs("AccessoryDate") & ""
				else
					strAccessoryStatus = rs("AccessoryStatus")
				end if
				strAccessoryColor = lcase(trim(rs("AccessoryBGColor") & ""))
				if request("FileType") <> "" and strAccessorycolor = "darkseagreen" then
					strAccessoryColor = "SeaGreen"
				elseif request("FileType") <> "" and strAccessorycolor = "lightsteelblue" then
					strAccessoryColor = "LightSkyBlue"
				end if				
			end if
			if ((request("ReportFormat") = "3" or request("ReportFormat") = "4") and productCount = 1) or (request("ReportFormat") <> "3" and request("ReportFormat") <> "4") then
				if rs("TestStatus") = "Date" then
					strTestStatus = rs("TestDate")
				elseif rs("TestStatus") = "QComplete" and rs("RiskRelease") then
					strTestStatus = "Risk&nbsp;Release"	
				elseif  rs("TestStatus") = "Investigating" and (blnShowTestStep or blnShowReleaseStep) then
					if lcase(left(rs("location") & "",9)) = "core team" then
					    strTestStatus = "Core&nbsp;Team"
					elseif lcase(left(rs("location") & "",11)) = "engineering" or lcase(left(rs("location") & "",8)) = "eng. dev" then
					    strTestStatus = "Engineering"
					else
					    strTestStatus = "Investigating"
					end if
				elseif (request("ReportFormat") = "1" or request("ReportFormat") = "" or request("ReportFormat") = "2" or request("ReportFormat") = "3") and lcase(trim(rs("TestStatus"))) = "service only" then
					strTestStatus = "Dropped"
				else
					strTestStatus = rs("TestStatus")
				end if
				
				if rs("TestStatus") = "Date" or rs("TestStatus") = "OOC" or rs("TestStatus") = "FCS" then
					if rs("TestConfidence") = 3 then
						strTestColor = "salmon"
					elseif rs("TestConfidence") = 2 then
						strTestColor = "#ffff99"
					else
						strTestColor = ""
					end if
				end if
					
				if	strTestColor = "" then
					strTestColor = trim(lcase(rs("MatrixBGColor") & ""))
				end if			

				if request("FileType") <> "" and strTestcolor = "darkseagreen" then
					strTestColor = "SeaGreen"
				elseif request("FileType") <> "" and strTestcolor = "lightsteelblue" then
					strTestColor = "LightSkyBlue"
				end if

				if Request("ReportFormat") = "4" then
					if not rs("Commodity") then
						strTestStatus = "N/A"
						strTestColor=""
					end if
				end if
			end if
			
			for i = 0 to ubound(ProductReleaseIDArray) 'ProductLoop
                strProductReleaseID = Split(ProductReleaseIDArray(i),";")                
				if (trim(strProductReleaseID(0)) = trim(rs("productversionid"))) and (trim(strProductReleaseID(1)) = trim(rs("ReleaseID"))) then
                    if (request("ReportFormat") <> "3" and request("ReportFormat") <> "4") then
						ProductBuckets(i) =  "<TD bgcolor=""" & strTestColor & """ align=center>" & strTestStatus & "</TD>"
					end if
					if request("ReportFormat") = "3" then
						if productCount = 1  then
							ProductBuckets(i) = "<TD bgcolor=""" & strTestColor & """ align=center>" & strTestStatus & "</TD>" & "<TD bgcolor=""" & strPilotColor & """ align=center>" & strPilotStatus & "</TD>"
						else
							ProductBuckets(i) = "<TD bgcolor=""" & strPilotColor & """ align=center>" & strPilotStatus & "</TD>"
						end if
					end if
					if request("ReportFormat") = "4" then
						if productCount = 1 then
							ProductBuckets(i) = "<TD bgcolor=""" & strTestColor & """ align=center>" & strTestStatus & "</TD>" & "<TD bgcolor=""" & strAccessoryColor & """ align=center>" & strAccessoryStatus & "</TD>"
						else
							ProductBuckets(i) = "<TD bgcolor=""" & strAccessoryColor & """ align=center>" & strAccessoryStatus & "</TD>"
						end if
					end if
				end if
			next
			if ubound(ProductIDArray) = 0 then
				ProductBuckets(0) = ProductBuckets(0) & "<TD nowrap>" & rs("TargetNotes") & "&nbsp;</TD>"
			end if
			rs.MoveNext			
		loop
		rs.Close
		
		if lastversion <> "" then
					for i = lbound(ProductBuckets) to ubound(ProductBuckets)                        
						if ProductBuckets(i) = "&nbsp;" then
							Response.Write "<TD>" & ProductBuckets(i) & "</TD>"
						else
							Response.Write ProductBuckets(i) 
						end if
						ProductBuckets(i) = "&nbsp;"
					next
					
					Response.Write "</TR>" 		
		end if
		
		Response.Write "</TABLE><BR><BR>"
		

		
	end if '****************
	
	set rs = nothing
	cn.close
	set cn = nothing

%>

<% if request("FileType")<> 1  and request("FileType")<> 2  then%>
	<%
		if Request.QueryString = "" then
			strQueryString2 = ""
		else
			strQueryString2 = "&" & Request.QueryString
		end if
	%>

	<DIV ID=divQuickReports style=display:>
	<TABLE BGCOLOR=lavender WIDTH="100%" BORDER=1 CELLSPACING=0 CELLPADDING=2>
		<TR>
			<TD>
			<TABLE border=0 width="100%"><TR>
			<TD width=100 nowrap valign=top><b><font size=2 color=navy>Quick&nbsp;Reports:</font>&nbsp;&nbsp;&nbsp;</b></TD>
			<TD width=170 nowrap valign=top>
				<b>Workflow</b><BR>
				<A href="QuickReports.asp?Report=1<%=strQueryString2%>" target=_blank>Development</a><BR>
				<A href="QuickReports.asp?Report=2<%=strQueryString2%>" target=_blank>Engineering Development</a>&nbsp;&nbsp;&nbsp;<BR>
				<A href="QuickReports.asp?Report=3<%=strQueryString2%>" target=_blank>Core Team</a>
			</TD>
			<TD width=110 valign=top><B>Qual Status</B><BR>
				<A href="QuickReports.asp?Report=4<%=strQueryString2%>" target=_blank>Investigating</a>&nbsp;&nbsp;&nbsp;<BR>
				<A href="QuickReports.asp?Report=10<%=strQueryString2%>" target=_blank>Scheduled</a><BR>
				<A href="QuickReports.asp?Report=5<%=strQueryString2%>" target=_blank>Failed</a><BR>
				<A href="QuickReports.asp?Report=6<%=strQueryString2%>" target=_blank>Qual Hold</a><br>
				<A href="QuickReports.asp?Report=19<%=strQueryString2%>" target=_blank>Risk Release</a><br>
				<A href="QuickReports.asp?Report=12<%=strQueryString2%>" target=_blank>Restricted</a>
			</TD>
			<TD width=110 valign=top><B>Pilot Status</B><BR>
				<A href="QuickReports.asp?Report=7<%=strQueryString2%>" target=_blank>Failed</a><BR>
				<A href="QuickReports.asp?Report=11<%=strQueryString2%>" target=_blank>Scheduled</a><BR>
				<A href="QuickReports.asp?Report=8<%=strQueryString2%>" target=_blank>Pilot Hold</a><BR>
				<A href="QuickReports.asp?Report=9<%=strQueryString2%>" target=_blank>Factory Hold</a><BR>
				<A href="QuickReports.asp?Report=13<%=strQueryString2%>&DateStart=<%=formatdatetime(Now()-7,vbshortdate)%>" target=_blank>Pilot Complete</a>
			</TD>
			<TD width=110 valign=top><B>Accessory Status</B><BR>
				<A href="QuickReports.asp?Report=14<%=strQueryString2%>" target=_blank>Failed</a><BR>
				<A href="QuickReports.asp?Report=15<%=strQueryString2%>" target=_blank>Scheduled</a><BR>
				<A href="QuickReports.asp?Report=16<%=strQueryString2%>" target=_blank>On Hold</a><BR>
				<A href="QuickReports.asp?Report=17<%=strQueryString2%>&DateStart=<%=formatdatetime(Now()-7,vbshortdate)%>" target=_blank>Complete</a>
			</TD>
			<TD valign=top><B>Availability</B><BR>
				<A href="QuickReportsAvailability.asp?Report=1&ProductStatus=3<%=strQueryString2%>" target=_blank>Production Audit</a><BR>
				<A href="QuickReportsAvailability.asp?Report=1&ProductStatus=4<%=strQueryString2%>" target=_blank>Post-Production Audit</a><BR>
			</TD>
			
			</TR></TABLE>
			<font face=verdana size=1 color=red>Note: Quick Reports contain information about current deliverable statuses only.  They currently do not contain any historical information.</font>
			</TD>
		</TR>
	</TABLE><BR>
	</div>
<%end if%>


<INPUT type="hidden" id=txtCurrentFilter name=txtCurrentFilter value="<%=replace(strQueryString,"+","%2B")%>" />
<INPUT type="hidden" id=txtScrollToRow name=txtScrollToRow value="<%=ScrollToRow%>" />
<%if rowCounter= 0 and not (strQueryString = "" or strQueryString = "ReportFormat=1" or strQueryString = "ReportFormat=2" or strQueryString = "ReportFormat=5" or strQueryString = "ReportFormat=3" or strQueryString = "ReportFormat=4") then%>
<font size=3 face=verdana><BR><BR>No commodities found matching your filter criteria.</font>
<%end if%>

<%
		'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
        if request("FileType")<> 1  and request("FileType")<> 2 then
'			Response.Write "<BR><BR>Dsiplay the <a target=""_blank"" href=""Commodity/QualMatrix.asp"">old report format</a> if you wish."	
			Response.Write "<BR><BR>Display the <a target=""_blank"" href=""HardwareMatrix.asp?FullReport=1"">weekly qualification report</a> (For Development and Production products only).<BR>"	
			Response.Write "Display the <a target=""_blank"" href=""HardwareMatrix.asp?FullReport=1&ReportFormat=2"">weekly subassembly report</a> (For Development and Production products only).<BR>"	
		end if

'	if currentuserid = 31 then
'		Response.Write Request.Form
'	end if

%>

<div id="dialog" style="display:none;" title="Dialog Title"><iframe id= iFrameID frameborder="0" scrolling="no" width="700" height="700" src=""></iframe></div>




<div id="releaseNotes" title="Release Notes"> 
    <div class="holder" id=""><iframe id ="rnFrame" frameborder="0" scrolling="no" src=""></iframe></div>
</div>


</body>
</html>


<%

function BuildHistoryFilter()
	dim SpecificStatusArray
	dim strSpecificPilotStatus
	dim strSpecificQualStatus
	dim strTempSQL
    dim strDateRange

			if request("txtSpecificPilotStatus") <> "" and instr(request("txtSpecificPilotStatus"),":") > 0 and instr(request("chkChangeType"),"22") > 0  then
				SpecificStatusArray = split(request("txtSpecificPilotStatus"),":")
				strSpecificPilotStatus = ""
				if trim(SpecificStatusArray(0)) <> "" then
					strSpecificPilotStatus = strSpecificPilotStatus & " and l.FromID in (" & SpecificStatusArray(0) & ") "
				end if
				if trim(SpecificStatusArray(1)) <> "" then
					strSpecificPilotStatus =  strSpecificPilotStatus & " and l.ToID in (" & SpecificStatusArray(1) & ") "
				end if
			else
				strSpecificPilotStatus = ""
			end if
			if request("txtSpecificQualStatus") <> "" and instr(request("txtSpecificQualStatus"),":") > 0 and instr(request("chkChangeType"),"21") > 0 then
				SpecificStatusArray = split(request("txtSpecificQualStatus"),":")
				strSpecificQualStatus = ""
				if trim(SpecificStatusArray(0)) <> "" then
					strSpecificQualStatus = strSpecificQualStatus & " and l.FromID in (" & SpecificStatusArray(0) & ") "
				end if
				if trim(SpecificStatusArray(1)) <> "" then
					strSpecificQualStatus =  strSpecificQualStatus & " and l.ToID in (" & SpecificStatusArray(1) & ") "
				end if
			else
				strSpecificQualStatus = ""
			end if
            dim strStartDate
            dim strEndDate
            dim tmpDate
            if true  then
 	            if request("cboHistoryRange") = "Range" then
		            if request("txtStartDate") = "" then
			            strStartDate = formatdatetime("1/1/1970",vbshortdate)
		            else
			            strStartDate =  scrubsql(request("txtStartDate"))
		            end if
                    if request("txtEndDate") = "" then
                        strEndDate = formatdatetime(now,vbshortdate)
                    else
                        strEndDate =  scrubsql(request("txtEndDate"))
                    end if

                    if datediff("d",strStartDate,strEndDate)< 0 then
                        'Switch them around if the end is before the start
			            tmpDate = strStartDate
			            strStartDate = strEndDate
			            strEndDate = tmpDate
		            end if
	            else
		            tmpDate = formatdatetime(dateadd("d",clng("-" & request("txtHistoryDays")),now),vbshortdate)
		            if request("cboHistoryRange") = "=" then
			            strStartDate = tmpDate
			            strEndDate = tmpDate
		            elseif request("cboHistoryRange") = ">=" then
			            strStartDate = "1/1/1970"
			            strEndDate = tmpDate
		            else
			            strStartDate = tmpDate
			            strEndDate = formatdatetime(Now,vbshortdate)
		            end if
	            end if	
            else
                strStartDate = request("txtStartDate")
                strEndDate = request("txtEndDate")
            end if
    	    if strStartDate <> "1/1/1970" then
	            strReportDateRange = "<br><br><font size=1 face=verdana>" & strStartDate & " - " & strEndDate & "<BR><BR></font>"
            else
	            strReportDateRange = "<br><br><font size=1 face=verdana>Before " & strEndDate & "<BR><BR></font>"
	        end if

            strDateRange = " l.Updated between '" & cdate(strStartDate) & "' and '" & Dateadd("d",1,cdate(strEndDate)) & "' "

		    strTempSQL = ""
			if instr(request("chkChangeType"),"21") > 0 then
				strTempSQL = strTempSQL & " Union Select pd.id from Actions a with (NOLOCK), ActionLog l with (NOLOCK), ProductVersion p with (NOLOCK), vendor vd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), TestStatus t1 with (NOLOCK), TestStatus t2 with (NOLOCK), product_deliverable pd with (NOLOCK) where pd.productversionid = p.id and pd.deliverableversionid = v.id and " & strDateRange & " " & strSpecificQualStatus & " and t1.id = l.FromID and t2.id = l.ToID and r.id = v.deliverablerootid and a.actionid = l.actionid and v.id = l.deliverableversionid and vd.id = v.vendorid and l.productversionid = p.id and l.actionid in(21) "
			end if
			if instr(request("chkChangeType"),"22") > 0 then
				strTempSQL = strTempSQL & " Union Select pd.id from Actions a with (NOLOCK), ActionLog l with (NOLOCK), ProductVersion p with (NOLOCK), vendor vd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), PilotStatus t1 with (NOLOCK), PilotStatus t2 with (NOLOCK), product_deliverable pd with (NOLOCK) where pd.productversionid = p.id and pd.deliverableversionid = v.id and " & strDateRange & " " & strSpecificPilotStatus & " and t1.id = l.FromID and t2.id = l.ToID and r.id = v.deliverablerootid and a.actionid = l.actionid and v.id = l.deliverableversionid and vd.id = v.vendorid and l.productversionid = p.id and l.actionid in (22) "
			end if


			if strTempSQL <> "" then
				strTempSQL = mid(strTempSQL,8)
			end if			
		BuildHistoryFilter = strTempSQL	
end function

sub DrawFilterBox (CurrentUserID, ShowFilters)
	dim LoopCount
	dim strNewQueryString
	dim strFilterPath
	dim FilterArray
	dim strFilter
	dim strPathSoFar
	dim strFilterName
	dim FilterValueArray
	dim i
	dim strNewString
	
	FilterArray = split(strQueryString,"&")
	strFilterPath = ""
	i=0
	for each strFilter in FilterArray
		FilterValueArray=split(strFilter,"=")
		if ubound(FilterValueArray) = 1 then
			if strPathSoFar = "" then
				strPathSoFar = strPathSoFar & strFilter
			else
				strPathSoFar = strPathSoFar & "&" & strFilter
			end if
			if lcase(FilterValueArray(0)) = "lstproducts" and isnumeric(FilterValueArray(1)) then
				rs.Open "Select DotsName from productversion with (NOLOCK) where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFilterName = "Invalid Product Specified"
				else
					strFilterName = rs("DotsName")
				end if
				rs.Close	
            elseif lcase(FilterValueArray(0)) = "lstproductspulsar" and isnumeric(FilterValueArray(1)) then
				rs.Open "Select Distinct DotsName from productversion with (NOLOCK) inner join ProductVersion_Release with (NOLOCK) on ProductVersion_Release.ProductVersionID = ProductVersion.ID where ProductVersion_Release.ID = " & FilterValueArray(1),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFilterName = "Invalid Product Specified"
				else
					strFilterName = rs("DotsName")
				end if
				rs.Close	
			elseif lcase(FilterValueArray(0)) = "lstfamily" and isnumeric(FilterValueArray(1)) then
				rs.Open "Select Name from productfamily with (NOLOCK) where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFilterName = strFilter
				else
					strFilterName = rs("Name")
				end if
				rs.Close	
			elseif lcase(FilterValueArray(0)) = "lstcategory" and isnumeric(FilterValueArray(1)) then
				rs.Open "Select Name from deliverablecategory with (NOLOCK) where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFilterName = strFilter
				else
					strFilterName = rs("Name")
				end if
				rs.Close	
			elseif lcase(FilterValueArray(0)) = "lstroot" and isnumeric(FilterValueArray(1)) then
				'if request("ReportFormat")="2" then
				'	rs.Open "Select Subassembly as Name from deliverableroot where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				'else
					rs.Open "Select Name from deliverableroot with (NOLOCK) where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				'end if
				if rs.EOF and rs.BOF then
					strFilterName = strFilter
				else
					strFilterName = rs("Name")
				end if
				rs.Close	
			elseif lcase(FilterValueArray(0)) = "lstphase" and isnumeric(FilterValueArray(1)) then
				rs.Open "Select Name from productstatus with (NOLOCK) where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				if rs.EOF and rs.BOF then
					strFilterName = strFilter
				else
					strFilterName = rs("Name")
				end if
				rs.Close	
			elseif lcase(FilterValueArray(0)) = "lstsubassembly" and isnumeric(FilterValueArray(1)) then
				'rs.Open "Select Subassembly as Name from deliverableroot where id = " & FilterValueArray(1),cn,adOpenForwardOnly
				'if rs.EOF and rs.BOF then
			'		strFilterName = strFilter
			'	else
					strFilterName = FilterValueArray(1)
			'	end if
			'	rs.Close	
			elseif lcase(FilterValueArray(0)) = "reportformat" then
				if ubound(FilterValueArray) > 0 then
					if trim(lcase(FilterValueArray(1))) = "5" then
						strFilterName = "Service"
					elseif trim(lcase(FilterValueArray(1))) = "4" then
						strFilterName = "Accessory"
					elseif trim(lcase(FilterValueArray(1))) = "3" then
						strFilterName = "Pilot Run"
					elseif trim(lcase(FilterValueArray(1))) = "2" then
						strFilterName = "Subassembly"
					else
						strFilterName = "Qualification"
					end if
				else
					strFilterName = "Unknown Filter"
				end if
			end if
			
			if i = ubound(FilterArray) then
				strFilterPath = strFilterPath & " > " & "<b>" & strFilterName & "</b>"
			else	
				strFilterPath = strFilterPath & " > " & "<a href=""HardwareMatrix.asp?" & strPathSoFar & """>" & strFilterName & "</a>"
			end if
			i=i+1
		end if
	next
	Response.Write "<div style=""display:" & showFilters & """ ID=FilterBox>"
	
	if strFilterPath <> "" then
		Response.Write "<font size=1 face=verdana><BR><a href=""HardwareMatrix.asp"">Home</a> " & strFilterPath & "<BR></font>"
	end if
	Response.Write "<BR><table class=""DisplayBar"" cellspacing=0 cellpadding=2><TR><TD valign=top><table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>"
	Response.Write "<td width=""100%"">"
	Response.Write "<table cellpadding=2 cellspacing=0  width=""100%"">"
	if isnumeric(request("lstFamily")) and request("lstFamily") <> "" then
		'Product Versions
		Response.Write "<TR>"
		Response.Write "<TD valign=top><font size=1 face=verdana><b>Product:&nbsp;</b></font></TD><TD valign=top width=100% ><font size=1 face=verdana>"
		rs.Open strSQLSelectProdListForDisplay & " " & strSQL & " order by pv.DotsName;",cn,adOpenForwardOnly
		LoopCount = 0
		if rs.EOF and rs.BOF then
			response.write "<font size=1 face=verdana>none</font>"
		else
			do while not rs.EOF
				LoopCount =LoopCount+1
				if LoopCount>1 then
				Response.Write " , "
				end if
			
				if trim(strProducts) = trim(rs("ID")) then
					Response.Write replace(rs("DotsName")," ","&nbsp;")
					strListHeaderName = rs("DotsName") 
				else
					strNewQueryString = strQueryString & "&lstProducts=" & rs("ID")	
					if left(strNewQueryString,1) = "&" then
						strNewQueryString = mid(strNewQueryString,2)
				end if	
					Response.Write "<a href=""HardwareMatrix.asp?" & strNewQueryString & """>" & replace(rs("DotsName")," ","&nbsp;") & "</a>"
				end if
				rs.MoveNext
			loop
		end if
		rs.Close
		Response.Write "</font></td>"
	else

		'Product Family
        strSQLSelectFamilyList = strSQLSelectFamilyList & " " & strSQL & " and isnull(pv.FusionRequirements, 0) =0 " 
        strSQLSelectFamilyList = strSQLSelectFamilyList  & " order by name;"	
    
        Response.Write "<TR>"
		Response.Write "<TD valign=top><font size=1 face=verdana><b>Product&nbsp;Family&nbsp;(Legacy):</b></font></TD><TD width=""100%"">"
        Response.Write "<select id=""cboProductFamily"" name=""cboProductFamily"" style=""width: 160px;"" language=""javascript"">"

		rs.Open strSQLSelectFamilyList,cn,adOpenForwardOnly    
       
		do while not rs.EOF			
            strNewQueryString = strQueryString & "&lstFamily=" & rs("ID")	                   					
			if left(strNewQueryString,1) = "&" then
				strNewQueryString = mid(strNewQueryString,2)
			end if    

		    Response.Write "<option value=" & "HardwareMatrix.asp?" & strNewQueryString & ">"
            Response.Write rs("Name")
            Response.Write "</option>"

			rs.MoveNext
		loop
	    rs.Close     
        Response.Write "</select>"
        Response.Write " <input type=""button"" value=""Go"" style=""height: 22px"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cbo_onchange(cboProductFamily)"">"

        Response.Write "</td>"

        'Close Filter Box
	    if strQueryString <> "" and strQueryString <> "ReportFormat=1"and strQueryString <> "ReportFormat=2" and strQueryString <> "ReportFormat=5" and strQueryString <> "ReportFormat=3"and strQueryString <> "ReportFormat=4" then
		    Response.Write "<TD rowspan=2 aligh=right valign=top><a href=""javascript: SwitchFilterView(2);""><IMG border=0 SRC=""../images/X2.gif""></a></TD>"
	    end if
	    Response.Write "</TR>"

        'Pulsar Product
        strSQLSelectPulsarProductList = strSQLSelectPulsarProductList & " " & strSQL  & " and isnull(pv.FusionRequirements, 0) =1 "
        strSQLSelectPulsarProductList = strSQLSelectPulsarProductList  & " order by name;"	

        Response.Write "<TR>"    
        Response.Write "<TD valign=top><font size=1 face=verdana><b>Product&nbsp;(Pulsar):</b></font></TD><TD width=""100%"">"
        Response.Write "<select id=""cboPulsarProduct"" name=""cboPulsarProduct"" style=""width: 160px;"" language=""javascript"">"
		    
		rs.Open strSQLSelectPulsarProductList,cn,adOpenForwardOnly  

	    do while not rs.EOF       		
            strNewQueryString = strQueryString & "&lstFamily=" & rs("ID") & "&lstProducts=" & rs("productversionID")                    		
            if left(strNewQueryString,1) = "&" then		
                strNewQueryString = mid(strNewQueryString,2)		
            end if        

		    Response.Write "<option value=" & "HardwareMatrix.asp?" & strNewQueryString & ">"
            Response.Write rs("Name")
            Response.Write "</option>"
    				
		    rs.MoveNext
	    loop
	    rs.Close     
        Response.Write "</select>"
        Response.Write " <input type=""button"" value=""Go"" style=""height: 22px"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cbo_onchange(cboPulsarProduct)"">"

        Response.Write "</td>"
	end if
    
	'Product Phase
	Response.Write "<TR>"
	Response.Write "<TD valign=top><font size=1 face=verdana><b>Product&nbsp;Phase:"
	Response.write "&nbsp;</b></font></TD><TD><font size=1 face=verdana>"
'		rs.Open strSQLSelectPhaseList & " " &  strSQl & " order by ps.id;",cn,adOpenForwardOnly
	rs.Open "Select ID,Name as Phase from productstatus with (NOLOCK) where id <> 5 order by id;"
	loopcount=0
	do while not rs.EOF
		if loopcount > 0 then
			response.write " | " 
		end if

		if trim(request("lstPhase")) = trim(rs("ID")) then
			strListHeaderName = rs("Phase") 
			Response.Write replace(replace(rs("Phase")&""," ","&nbsp;"),"-","&#8209;")
		else
			strNewQueryString = StripParameter(strQueryString,"lstPhase") & "&lstPhase=" & rs("ID")	
			if left(strNewQueryString,1) = "&" then
				strNewQueryString = mid(strNewQueryString,2)
			end if
			Response.Write "<a href=""HardwareMatrix.asp?" & strNewQueryString & """>"
			Response.Write  replace(replace(rs("Phase")&""," ","&nbsp;"),"-","&#8209;")
			Response.Write "</a>"
		end if


		loopcount=loopcount+1
		rs.MoveNext
	loop
	rs.Close
		
	Response.write "</td></tr>"


	if isnumeric(request("lstCategory")) and request("lstCategory") <> "" then

		'Component for selected Category
        dim component
        responseBuffer = ""
		Response.Write "<TR>"
		Response.Write "<TD valign=top><font size=1 face=verdana><b>"
		if request("ReportFormat")="2" or request("ReportFormat")="5" then
			Response.write "Subassembly:"
            component = "Subassembly"
		else
			Response.write "Root&nbsp;Deliverable:"
            component = "Name"
		end if
		Response.write "&nbsp;</b></font></TD><TD><font size=1 face=verdana>"
    			 
		rs.Open strSQLSelectDelList & " " &  strSQl & " order by r.name;",cn,adOpenForwardOnly

		if rs.EOF and rs.BOF then
			Response.Write "<font size=1 face=verdana>none</font>"
		else
            responseBuffer = "<select id=""cboComponent"" name=""cboComponent"" style=""width: 160px;"" language=""javascript"">"
            firstCache = rs(component)

            LoopCount = 0
            do while not rs.EOF
		        LoopCount = LoopCount+1

    			if request("ReportFormat")="2" or request("ReportFormat")="5" then
					strNewQueryString = strQueryString & "&lstSubassembly=" & rs("Subassembly")	
				else
					strNewQueryString = strQueryString & "&lstRoot=" & rs("ID")	
				end if

				if left(strNewQueryString,1) = "&" then
					strNewQueryString = mid(strNewQueryString,2)
				end if
				
		        responseBuffer = responseBuffer & "<option value=" & "HardwareMatrix.asp?" & strNewQueryString & ">" & rs(component) & "</option>"
    
				rs.MoveNext
			loop

            responseBuffer = responseBuffer & "</select>"
            responseBuffer = responseBuffer & " <input type=""button"" value=""Go"" style=""height: 22px"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cbo_onchange(cboComponent)"">"

            if LoopCount = 1 then
                responseBuffer = replace(firstCache," ","&nbsp;")
            end if

            Response.Write responseBuffer
		end if
		rs.Close
        Response.Write "</td>"

	else
	
		'Category
        responseBuffer = ""
        Response.Write "<TR>"    
        Response.Write "<TD valign=top><font size=1 face=verdana><b>Category:&nbsp;</b></font></TD><TD width=""100%"">"

		if not blnFiltersSelected then
			rs.Open "Select ID, Name from deliverablecategory with (NOLOCK) where deliverabletypeid=1 order by name;",cn,adOpenForwardOnly
		else
			rs.Open strSQLSelectCatList & " " & strSQL & " order by c.name;",cn,adOpenForwardOnly
		end if

		if rs.EOF and rs.BOF then
			Response.Write "<font size=1 face=verdana>none</font>"
		else
			responseBuffer = "<select id=""cboCategory"" name=""cboCategory"" style=""width: 160px;"" language=""javascript"">"
            firstCache = rs("Name")        

            LoopCount = 0
            do while not rs.EOF
		        LoopCount = LoopCount+1
                    
			    strNewQueryString = strQueryString & "&lstCategory=" & rs("ID")
			    if left(strNewQueryString,1) = "&" then
				    strNewQueryString = mid(strNewQueryString,2)
			    end if

		        responseBuffer = responseBuffer & "<option value=" & "HardwareMatrix.asp?" & strNewQueryString & ">" & rs("Name") & "</option>"
    
			    rs.MoveNext
		    loop
    
            responseBuffer = responseBuffer &  "</select>"
            responseBuffer = responseBuffer &  " <input type=""button"" value=""Go"" style=""height: 22px"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cbo_onchange(cboCategory)"">"

            if LoopCount = 1 then
                responseBuffer = replace(firstCache," ","&nbsp;")
            end if

            Response.Write responseBuffer
        end if
	    rs.Close     
        Response.Write "</td>"

	end if	
		
	Response.Write "<TR height=10><TD colspan=2 height=1><HR height=1></td></tr>"
	Response.Write "<tr>"
	Response.Write "<TD valign=top><font size=1 face=verdana><b>Custom&nbsp;Filter:&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana>"
	rs.Open "Select ProfileName, ID, ReportFilters from ReportProfiles (nolock) where value8=1 and profiletype=3 and EmployeeID=" & CurrentUserID & " order by profilename",cn,adOpenForwardOnly
	LoopCount = 0
	
	Response.Write "<a target=""_blank"" href=""../Query/Deliverables.asp?HardwareMatrix=1"">Create New Filter</a></td></tr>"
	
	if not (rs.EOF and rs.BOF) then
		Response.Write "<tr><TD><b>Saved Filters:</b></td><td>"
	end if
	do while not rs.EOF
		LoopCount =LoopCount+1
		if loopcount <> 1 then
			Response.Write " , "
		end if
	
		if request("Profile") = rs("ID") then
			Response.Write replace(rs("ProfileName")," ","&nbsp;")
			strListHeaderName = rs("ProfileName") 
		else
			'Response.Write "<a href=""HardwareMatrix.asp?Profile=" & rs("ID") & """>" & replace(rs("ProfileName")," ","&nbsp;") & "</a>"
			Response.Write "<a href=""HardwareMatrix.asp?" & rs("ReportFilters") & """>" & replace(rs("ProfileName")," ","&nbsp;") & "</a>"
		end if
		rs.MoveNext
	loop
	rs.Close
	Response.Write "</font></td>"
	Response.Write "</TR>"
	Response.Write "<TR><TD colspan=2><HR></td></tr><tr>"
	Response.Write "<TD valign=top><font size=1 face=verdana><b>Report:&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana>"
	strNewString = StripParameter(strQueryString,"ReportFormat")
	if strNewString <> "" then
		strNewString = "&" & strNewString
	end if
	if request("ReportFormat") = "5" then
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=1" & strNewString & """>Qualification</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=2" & strNewString  & """>Subassembly</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=3" & strNewString  & """>Pilot Run</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=4" & strNewString  & """>Accessory</a> | "
		Response.Write "Service"
	elseif request("ReportFormat") = "4" then
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=1" & strNewString & """>Qualification</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=2" & strNewString  & """>Subassembly</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=3" & strNewString  & """>Pilot Run</a> | "
		Response.Write "Accessory | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=5" & strNewString  & """>Service</a>"
	elseif request("ReportFormat") = "3" then
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=1" & strNewString & """>Qualification</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=2" & strNewString  & """>Subassembly</a> | "
		Response.Write "Pilot Run | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=4" & strNewString  & """>Accessory</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=5" & strNewString  & """>Service</a>"
	elseif request("ReportFormat") = "2" then
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=1" & strNewString  & """>Qualification</a> | "
		Response.write "Subassembly | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=3" & strNewString  & """>Pilot Run</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=4" & strNewString  & """>Accessory</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=5" & strNewString  & """>Service</a>"
	else
		Response.Write "Qualification | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=2" & strNewString  & """>Subassembly</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=3" & strNewString  & """>Pilot Run</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=4" & strNewString  & """>Accessory</a> | "
		Response.Write "<a href=""HardwareMatrix.asp?ReportFormat=5" & strNewString  & """>Service</a>"
	end if
	Response.Write "</font></td>"
	Response.Write "</TR>"
    
    'add filter for releases
    dim strSQLSelectReleases
    dim bFirstWrite : bFirstWrite = false

    if request("lstproducts") <> "" and isnumeric(request("lstproducts")) then
        'Pulsar Product
        strSQLSelectReleases = "select pv_r.ID as ProductVersionReleaseID,pvr.Name as ReleaseName from ProductVersion_Release pv_r with(nolock) inner join ProductVersionRelease pvr with(nolock) on pv_r.ReleaseID = pvr.ID where pv_r.ProductVersionID in (" & request("lstproducts") & ")"
        strSQLSelectReleases = strSQLSelectReleases  & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc;"	
    	 
	   rs.Open strSQLSelectReleases,cn,adOpenForwardOnly  
        if not (rs.EOF and rs.BOF) then
            Response.Write "<TR><TD colspan=2><HR></td></tr><tr>"
	        Response.Write "<TD valign=top><font size=1 face=verdana><b>Releases:&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana>"
	        strNewString = StripParameter(strQueryString,"lstProductsPulsar")
            if request("lstProductsPulsar") = "" then
                Response.Write "All"
            else
                Response.Write "<a href=""HardwareMatrix.asp?" & strNewString  & """>All</a>"
            end if
        
	        do while not rs.EOF  
                If Not bFirstWrite Then
				    Response.Write "&nbsp;|&nbsp;"
			    End If
                if cint(request("lstProductsPulsar")) = cint(rs("ProductVersionReleaseID")) then
                    Response.Write rs("ReleaseName")
                else
		            Response.Write "<a href=""HardwareMatrix.asp?lstProductsPulsar="& rs("ProductVersionReleaseID") & "&" & strNewString  & """>"
                    Response.Write rs("ReleaseName")
                    Response.Write "</a>"    	
                end if		
                bFirstWrite = False	
		        rs.MoveNext
	        loop	        
            Response.Write "</font></td>"
	        Response.Write "</TR>"
        end if
        rs.Close
    end if	

	Response.Write "</table>"
	Response.Write "</td></tr>"
	Response.Write "</table>"
	Response.Write "</TD>"
	Response.Write "</TR>"
	Response.Write "</table>"
	Response.Write "<BR></div>"

end sub


	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i

		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

	function FormatInputStrings()

		dim strBuffer
		dim BufferArray
		dim strItem
		dim strLastParam
		dim ItemArray
		dim strOutBuffer
	
		strOutBuffer = ""
	
		strBuffer = Request.Form
		if trim(strBuffer) <> "" then
			strBuffer = strBuffer & "&"
		end if
		strBuffer = strBuffer & strQueryString
		
		
		BufferArray = split(strBuffer,"&")
'		
'		strLastParam = ""
		for each strItem in BufferArray
			ItemArray = split(strItem,"=")
			if ubound(ItemArray) > 0 then
				if strLastParam <> trim(ItemArray(0)) then
					if trim(request(trim(ItemArray(0)))) <> "" then
						if lcase(trim(ItemArray(0))) <> "txtdivision" and lcase(trim(ItemArray(0))) <> "cboprofile" and lcase(trim(ItemArray(0))) <> "chknamesearch" and lcase(trim(ItemArray(0))) <> "cboformat" then
							strOutBuffer = strOutBuffer & "&" & trim(ItemArray(0)) & "=" & request(trim(ItemArray(0)))
						end if
					end if
				end if
			strLastParam = trim(ItemArray(0))
			end if
		next
		
		if len(strOutBuffer) > 0 then
			strOutbuffer = replace(mid(strOutBuffer,2),", ",",")
		end if
		FormatInputStrings =  strOutBuffer
	end function


function StripParameter(strQuery,strName)
	dim ParameterArray
	dim strParameter
	dim strOutput
	
	ParameterArray = split(strQuery,"&")
	strOutput = ""
	for each strParameter in ParameterArray
		if instr(lcase(strParameter),lcase(strName)) = 0 then
			strOutput = strOutput & "&" & strParameter
		end if
	next
	
	if strOutput <> "" then
		strOutput = mid(strOutput,2)
	end if
	
	StripParameter = strOutput
	
end function

%>


<%@ Language=VBScript %>

<%
	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
	end if
	%>

<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<title>Deliverable Query Results</title>
<meta name="GENER" content="Microsoft Visual Studio 6.0" />
<script  id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->
var oPopup = window.createPopup();

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;

	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}

	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function window_onload() {
	if (typeof(OTSSummaryTable) != "undefined")
		{
		OTSSummaryTableTop.innerHTML = OTSSummaryTable.innerHTML;
		OTSSummaryTable.innerHTML = "";
		}
	else if (typeof(OTSSummaryTableTop) != "undefined")
		OTSSummaryTableTop.innerHTML = "";

}

function row_onmouseover() {
	window.event.srcElement.parentElement.style.cursor = "hand"
	window.event.srcElement.parentElement.style.color = "red"
}
function row_onmouseout() {
	window.event.srcElement.parentElement.style.color = "black"
}

function row_onclick(ProdDelID,RootID,VersionID) {
	DeliverableMenu(ProdDelID,RootID,VersionID);
}

function DisplayVersion(RootID, VersionID) {
	var strID;
	strID = window.showModalDialog("../WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 

	if (typeof(strID) != "undefined")
		{
			window.location.reload(true);
		}
}

function DisplayDetails(RootID, VersionID) {
	window.open("DeliverableVersionDetails.aspx?Type=1&RootID=" + RootID + "&ID=" + VersionID, "_blank",
        "width=800,status=1,resizable=1,scrollbars=1") 
}

function GetVersion(VersionID){
	var strPath = trim(document.all("Path" + VersionID).innerText);   
	window.open ("file://" + strPath);
}

function DeliverableMenu(ProdDelID,RootID, VersionID)
{
    // The variables "lefter" and "topper" store the X and Y coordinates
    // to use as parameter values for the following show method. In this
    // way, the popup displays near the location the user clicks. 
    var lefter = event.clientX;
    var topper = event.clientY;
    var popupBody;
	var strPath = trim(document.all("Path" + ProdDelID).innerText);    

	if(typeof(oPopup) == "undefined") 
		{
		DisplayVersion(RootID,VersionID);
		return;
		}
  if (strPath != "")
	{
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";
	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";

	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:GetVersion(" + ProdDelID + ")'\" >&nbsp;&nbsp;&nbsp;Download</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDetails(" + RootID + "," + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Detail&nbsp;Info</FONT></SPAN></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + RootID + "," + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Properties</FONT></SPAN></DIV>";

	popupBody = popupBody + "</DIV>";

    oPopup.document.body.innerHTML = popupBody; 

	oPopup.show(lefter, topper, 160, 82, document.body);
	}
	else
		DisplayVersion(RootID,VersionID);
}

function ROW_onmouseover() {
	var srcElem = window.event.srcElement;
	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;
	if (srcElem.className =="Row")
		{
		srcElem.style.backgroundColor = "Thistle";
		srcElem.style.cursor="hand";
		}


}

function ROW_onmouseout() {
	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row

	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;

	if (srcElem.className =="Row")
		srcElem.style.backgroundColor = "Ivory";

}


function OTSROW_onclick(ID){

	var strResult;

	strResult = window.open("../search/ots/Report.asp?txtReportSections=1&txtObservationID=" + ID,"_blank","width=700, height=400,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes") 



}



function DelROW_onclick(ID, RootID){

	var strResult;

	strResult = window.showModalDialog("../WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + ID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 



}



function ShowDeliverableDetails(ID){

	var strResult;

	strResult = window.open("DeliverableVersionDetails.asp?ID=" + ID,"_blank","width=700, height=400,location=yes, menubar=yes, status=yes,toolbar=yes, scrollbars=yes, resizable=yes");



}





//-->

</script>
</head>
<style>
TABLE
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana, Tahoma, Arial
}
A:link
{
    COLOR: Blue;
}
A:visited
{
    COLOR: Blue;
}
A:hover
{
    COLOR: red;
}
TD.offsetcell
{
	BORDER-LEFT-STYLE: none;
	BORDER-RIGHT-STYLE: none;
	BORDER-TOP-STYLE: none;
	BORDER-Bottom-STYLE: none;
	BACKGROUND-COLOR: white;
	width=1;
}

TD.SectionHeader
{
	BACKGROUND-COLOR: gainsboro;
}

TD.RootHeader
{
	BACKGROUND-COLOR: gainsboro;
}

</STYLE>

<BODY LANGUAGE=javascript onload="return window_onload()">
<%if request("txtTitle") = "" then%>
	<H3><font face=verdana>Deliverable Report</font></H3>
<%elseif request("txtFunction") = "7" and (request("txtTitle") = "Deliverable Report" or request("txtTitle") = "") then%>
	<H3><font face=verdana>Deliverables Missing Translations</font></H3>
<%else%>
	<H3><font face=verdana><%= Server.HTMLEncode(request("txtTitle"))%></font></H3>
<%end if%>
<%
	function ScrubSQL(strWords) 
		dim badChars 
		dim newChars 

	'	strWords=replace(strWords,"'","''")

		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 

		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 

		ScrubSQL = newChars 

	end function 
	dim ColumnList
	dim ColumnArray
	dim TestStatusArray
	dim DeveloperStatusArray
	dim i
	dim strStartDate, strEndDate
	dim tmpDate
    if  trim(request("txtFunction")) = "1"  and request("lstColumns")="" then
	    ColumnList = "ID,Name,Release,Version,Vendor,Vendor Version,Part Number,IRS Part Number,Category,Developer,Dev Manager,Workflow"
    elseif request("lstColumns")="" then
	    ColumnList = "ID,Name,Version,Vendor,Vendor Version,Part Number,IRS Part Number,Category,Developer,Dev Manager,Workflow"
    else
		ColumnList = request("lstColumns")
	end if
    
	ColumnArray = split(ColumnList,",")
	TestStatusArray = split("TBD,Passed,Failed,Blocked,Watch,N/A",",")
	DeveloperStatusArray = split("TBD,Approved,Disapproved",",")
	
	
	if trim(request("txtFunction")) = "6" then
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
	    if strStartDate <> "1/1/1970" then
	        response.Write "<font size=1 face=verdana>" & strStartDate & " - " & strEndDate & "<BR><BR></font>"
	    end if
    end if	
	
	
	dim strSQL
	dim cn
	dim strProducts
	dim strOS
	dim strLanguages
	dim strStatus
	dim strType
	dim strRange1
	dim strRange2
	dim strVendor
	dim strCategory
	dim strDeveloper
	dim strmanager
	dim strRoot
	dim strBaseSQL
	dim strOTSIDList
	dim strSpecificQualStatus
	dim strSpecificPilotStatus
	dim SpecificStatusArray

	set cn = server.CreateObject("ADODB.Connection")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.IsolationLevel=256
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
		strSQL = strSQl &  "FROM ProductFamily AS f WITH (NOLOCK) " & _
				"INNER JOIN ProductVersion AS pv WITH (NOLOCK) ON f.ID = pv.ProductFamilyID " & _
				"INNER JOIN DeliverableRoot AS r WITH (NOLOCK) " & _
				"INNER JOIN DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID " & _
				"INNER JOIN Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID " & _
				"INNER JOIN DeliverableCoreTeam AS ct WITH (NOLOCK) ON r.CoreTeamID = ct.ID " & _
				"INNER JOIN DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID " & _
				"INNER JOIN Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID " & _
				"INNER JOIN Employee AS e2 WITH (NOLOCK) ON r.DevManagerID = e2.ID " & _
				"INNER JOIN Employee AS e WITH (NOLOCK) ON v.DeveloperID = e.ID ON pv.ID = pd.ProductVersionID " & _
				"LEFT OUTER JOIN PilotStatus AS ps WITH (NOLOCK) ON pd.PilotStatusID = ps.ID " & _
				"LEFT OUTER JOIN TestStatus AS ts WITH (NOLOCK) ON pd.TestStatusID = ts.ID " & _
                "LEFT OUTER JOIN Product_Deliverable_Release  AS PDR WITH (NOLOCK) ON pd.ID = PDR.ProductDeliverableID " & _
                "LEFT OUTER JOIN ProductVersion_Release  AS pv_r WITH (NOLOCK) ON pv.id = pv_r.ProductVersionID " & _
                "LEFT OUTER JOIN ProductVersionRelease  AS pvr WITH (NOLOCK) ON pvr.id = pv_r.ReleaseID " & _
				"WHERE 1=1 "

	strBaseSQL = strSQL
    dim strProductsPulsar
    dim productReleases
    dim productPulsarIds
    dim productReleaseIds
    dim productReleaseId
    
    productPulsarIds=""
    productReleaseIds=""
    productReleaseId=""

    strProductsPulsar = ""
	strProducts = request("lstProducts")
     ' wgomero: PBI 18749 add Pulsar products to the list if any
    strProductsPulsar = request("lstProductsPulsar")
    productReleases = split(request("lstProductsPulsar"),",")

    for each productRelease in productReleases
      if instr(productRelease,":")>0 then
		productReleaseId = split(productRelease,":")
        productPulsarIds = productPulsarIds + "," + productReleaseId(0)
        productReleaseIds = productReleaseIds + "," + productReleaseId(1)
      end if
    next
    if strProducts <> "" and productPulsarIds <> "" then
        strProducts = strProducts + "," + mid(productPulsarIds,2)
    end if

    if strProducts = "" and productPulsarIds <> "" then
        strProducts =  productPulsarIds
    end if

    if strProducts <> "" and productPulsarIds = "" then
        strProducts =  strProducts
    end if
   'end

	if left(strProducts,1) = "," then
		strProducts = mid(strProducts,2)
	end if

    if left(productReleaseIds,1) = "," then
		productReleaseIds = mid(productReleaseIds,2)
	end if


'*******Process Product Groups 
	if request("lstProductGroups") <> "" then
		dim ProductGroupsArray
		dim ProductGroupArray
		dim strProductGroup
		dim lastProductGroup
		dim strProductGroupFilter
		dim strCycleList
		ProductGroupsArray = split(request("lstProductGroups"),",")
		lastProductGroup = 0
		strProductGroupFilter = ""
		strCycleList = ""
 
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
		strSQl = strSQL & " and pd.ProductVersionID in ( " & ScrubSQL(strProducts) &  " ) "
	end if
    if productReleaseIds <> "" then
		strSQl = strSQL & _ 
        " and (" & _ 
        "     (pv_r.id in(" & ScrubSQL(productReleaseIds) & ") and pv.FusionRequirements is not NULL) " & _ 
        "  or (pv_r.id is NULL and pv.FusionRequirements is NULL)" & _ 
        "     )"
	end if
    strOS = request("lstOS")
	if left(strOS,1) = "," then
		strOS = mid(strOS,2)
	end if

	if strOS <> "" then
		strSQl = strSQL & " and v.ID in (Select DeliverableVersionID from os_delver with (NOLOCK) where osid in ( " & ScrubSQL(strOS) & ") union select 0) "
	end if

	strLanguages = request("lstLanguage")
	if left(strLanguages,1) = "," then
		strLanguages = mid(strLanguages,2)
	end if
	if strLanguages <> "" then
		strSQl = strSQL & " and v.ID in (Select DeliverableVersionID from language_delver with (NOLOCK) where languageid in ( " & ScrubSQL(strLanguages) & ")  union select 0 ) "
	end if

	strVendor = request("lstVendor")
	if left(strVendor,1) = "," then
		strVendor = mid(strVendor,2)
	end if
	if strVendor <> "" then
		strSQl = strSQL & " and v.VendorID in (" & ScrubSQL(strVendor) & ") "
	end if

	strCategory = request("lstCategory")
	if left(strCategory,1) = "," then
		strCategory = mid(strCategory,2)
	end if
	if strCategory <> "" then
		strSQl = strSQL & " and r.CategoryID in (" & ScrubSQL(strCategory) & ") "
	end if

	strCoreTeam = request("lstCoreTeam")
	if left(strCoreTeam,1) = "," then
		strCoreTeam = mid(strCoreTeam,2)
	end if
	if strCoreTeam <> "" then
		strSQl = strSQL & " and r.CoreTeamID in (" & ScrubSQL(strCoreTeam) & ") "
	end if



	strDeveloper = request("lstDeveloper")
	if left(strDeveloper,1) = "," then
		strDeveloper = mid(strDeveloper,2)
	end if

	if strDeveloper <> "" then
		strSQl = strSQL & " and v.DeveloperID in (" & ScrubSQL(strDeveloper) & ") "
	end if

	strManager = request("lstDevManager")
	if strManager <> "" then
		strSQl = strSQL & " and r.DevManagerID in (" & ScrubSQL(strManager) & ") "
	end if

	if request("Type") <> "" then
		if isnumeric(request("Type")) then
			strSQL = strSQL & " and r.typeid = " & clng(request("type")) & " "
		end if
	end if

	if request("chkSCRestricted") <> "" then
		strSQl = strSQL & " and ( (pd.supplychainrestriction = 1 and pv.FusionRequirements <> 1) or (pdr.supplychainrestriction = 1 and pv.FusionRequirements = 1) ) "
	end if

	strRoot = request("lstRoot")
	if strRoot <> "" then
		strSQl = strSQL & " and r.ID in (" & ScrubSQL(strRoot) & ") "
	end if

	if request("chkTarget") = "on" then
		strSQl = strSQL & " and ( (pd.Targeted = 1 and pv.FusionRequirements <> 1) or (pdr.Targeted = 1 and pv.FusionRequirements = 1) ) "
	end if

	if request("chkInImage") = "on" then
		strSQl = strSQL & " and pd.InImage = 1 "
	end if

	dim strStrings
	strStrings=""

	if request("txtSearch") <> "" then
		dim strSearch

		strSearch = replace(replace(replace(replace(request("txtSearch"),"""",""),"'",""),"%",""),"*","")
		strSearch=ScrubSQL(strSearch)

		if request("chkNameSearch") <> "" then
			strStrings = strStrings & " or  r.Name like '%" & strSearch & "%' or  v.DeliverableName like '%" & strSearch & "%' "
		end if
		if request("chkChangesSearch") <> "" then
			strStrings = strStrings & " or v.Changes like '%" & strSearch & "%' "
		end if
		if request("chkDescriptionSearch") <> "" then
			strStrings = strStrings & " or r.Description like '%" & strSearch & "%' "
		end if
		if request("chkCommentsSearch") <> "" then
			strStrings = strStrings & " or v.Comments like '%" & strSearch & "%' "
		end if

		if strStrings <> "" then
			strStrings = mid(strStrings,4)
			strSQL = strSQL & " and ( " & strstrings & " ) "
		end if
	end if
	strStrings=""
	if request("chkDevelopment") <> "" or request("chkTest") <> "" or request("chkRelease") <> "" or request("chkComplete") <> "" then
		if request("chkDevelopment") <> "" then
			strStrings = strStrings & " or  v.location like '%Development%' "
		end if
		if request("chkTest") <> "" then
			strStrings = strStrings & " or  v.location like '%Test%' "
		end if
		if request("chkRelease") <> "" then
			strStrings = strStrings & " or  v.location like '%Release%' "
		end if
		if request("chkComplete") <> "" then
			strStrings = strStrings & " or  v.location like '%Complete%' or  v.location like '%PM%' "
		end if

		if strStrings <> "" then
			strStrings = mid(strStrings,4)
			strSQL = strSQL & " and ( " & strstrings & " ) "
		end if
	end if

	if request("chkFailed") <> "" then
		strSQL = strSQL & " and  v.location like '%Failed%' "
	end if
	if trim(request("txtNumbers")) <> "" then
		strSQl = strSQl & " and v.id in ( " & ScrubSQL(request("txtNumbers")) & " ) "
	end if

    if trim(request("txtFunction")) = "6" and strStartDate <> "" and strEndDate <> "" then
		strSQl = strSQl & " and v.actualreleasedate between '" & cdate(strStartDate) & "' and '" & cdate(strEndDate)+1 & "' "
	end if
	
	if request("txtAdvanced") <> "" then
		strSQl = strSQl & " and ( " & ScrubSQL(request("txtAdvanced")) & " ) "
	end if


	if request("txtFunction") = "7" then
		strSQL = strSQL & " ORDER BY r.name;"
		strBaseSQL = strBaseSQL & " ORDER BY r.name;"
	elseif request("txtFunction") = "3" or request("txtFunction") = "4" then
		strSQL = strSQL & " ORDER BY r.name, v.id desc, pv.DOTSName;"
		strBaseSQL = strBaseSQL & " ORDER BY r.name, v.id desc, pv.DOTSName;"
	elseif request("txtFunction") = "5" then
		strSQL = strSQL & " ORDER BY r.name, v.id desc;"
		strBaseSQL = strBaseSQL & " ORDER BY r.name, v.id desc;"
	end if
	
	dim strActions
	dim strResolution
	dim strApprovals
	dim strActual
	dim strTarget
	dim strVersion
	dim strVendorVersion
	dim strTargeted
	dim strInImage
	dim strLanguage
	dim RowCount
	dim PathID
	dim ObservationRowCOunt

	RowCount=0
	if 	strBaseSQL = strSQL and request("txtFunction") <> "2" then
		Response.Write "<font size=2 face=verdana>No report criteria selected.  Please select the appropriate criteria and try again.</font>"
	else
		if request("txtFunction") = "7" then
				Response.Write "test"
		elseif request("txtFunction") = "1"  or request("txtFunction") = "5" then 'request("txtFunction") <> "3"  and request("txtFunction") <> "4"  then
			dim blnProductsColumns
			blnProductsColumns = false

			for i = 0 to Ubound(ColumnArray)						
				if lcase(trim(ColumnArray(i))) = "product" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "targeted" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "in image" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "mit signoff" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "odm signoff" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "wwan signoff" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "dev signoff" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "wwan samples" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "wwan notes" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "mit samples" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "mit notes" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "odm samples" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "odm notes" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "hw qual status" then
					blnProductsColumns = true		
					exit for
				elseif lcase(trim(ColumnArray(i))) = "pilot status" then
					blnProductsColumns = true		
					exit for
				end if
			next
            'make the whole inline query readable, not everything in one line
            dim strFieldNames    
            strFieldNames = "SELECT ct.name as CoreTeam," & _
                            "MITTestNotes = case when PDR.ProductDeliverableID is null then pd.IntegrationTestNotes" & _
			                "               else isnull(PDR.IntegrationTestNotes,'')" & _
			                "               end, " & _
                            "WWANTestNotes = case when PDR.ProductDeliverableID is null then pd.WWANTestNotes" & _
			                "                else isnull(PDR.WWANTestNotes,'')" & _
			                "                end, " & _
                            "ODMTestNotes = case when PDR.ProductDeliverableID is null then pd.ODMTestNotes" & _
			                "               else isnull(PDR.ODMTestNotes,'')" & _
			                "               end, " & _ 
                            "WWANSamples =  case when PDR.ProductDeliverableID is null then pd.WWANUnitsReceived" & _
			                "               else PDR.WWANUnitsReceived" & _
			                "               end, " & _ 
                            "ODMSamples = case when PDR.ProductDeliverableID is null then pd.ODMUnitsReceived" & _
			                "             else PDR.ODMUnitsReceived" & _
			                "             end, " & _ 
                            "MITSamples = case when PDR.ProductDeliverableID is null then pd.IntegrationUnitsReceived" & _
			                "             else PDR.IntegrationUnitsReceived" & _
			                "             end, " & _
                            "v.serviceeoaDate, v.Serviceactive," & _
                            "PilotStatus = case when PDR.ProductDeliverableID is null then ps.name" & _
			                "              else (select name from PilotStatus where id=PDR.PilotStatusID)" & _
			                "              end, " & _
                            "testdate = case when PDR.ProductDeliverableID is null then pd.testdate" & _
			                "           else pdr.testdate" & _
			                "           end, " & _
                            "riskrelease = case when PDR.ProductDeliverableID is null then pd.riskrelease" & _
			                "              else pdr.riskrelease" & _
			                "              end, " & _
                            "pilotdate = case when PDR.ProductDeliverableID is null then pd.pilotdate" & _
			                "            else pdr.pilotdate" & _
			                "            end, " & _
                            "QualStatus = case when PDR.ProductDeliverableID is null then ts.status" & _
			                "             else (select status from TestStatus where id=PDR.TestStatusID)" & _
			                "             end, " & _
                            "v.EndOfLifeDate, v.codename, v.active, " & _
                            "developerteststatus = case when PDR.ProductDeliverableID is null then pd.developerteststatus " & _ 
                            "            else Case When pdr.developerteststatus is null then 0 else pdr.developerteststatus end" & _
			                "            end, " & _
                            "wwanteststatus = case when PDR.ProductDeliverableID is null then pd.wwanteststatus" & _
			                "                 else isnull(PDR.wwanteststatus,0)" & _
			                "                 end, " & _ 
                            "odmteststatus = case when PDR.ProductDeliverableID is null then pd.odmteststatus" & _
			                "                else isnull(PDR.odmteststatus,0)" & _
			                "                end, " & _ 
                            "integrationteststatus = case when PDR.ProductDeliverableID is null then pd.integrationteststatus" & _
			                "                        else isnull(PDR.integrationteststatus,0)" & _
			                "                        end, " & _  
                            "vd.name as Vendor, " & _
                            "r.id as RootID, r.TypeID as DelTypeID, c.name as category, pd.SelectiveRestore, pd.ARCD, pd.DRDVD, pd.RACD_Americas, pd.RACD_APD, isnull(pdr.RACD_EMEA,pd.RACD_EMEA) as RACD_EMEA, pd.preinstall, pd.preload," & _
                            " pd.DropInBox, pd.patch, pd.web, pd.OSCD, pd.doccd, r.ID as RootID, pv.typeid, pv.productstatusid, pv.id as ProductID," & _
                            " Product = pv.DOTSName + case when PDR.ProductDeliverableID is null then ''" & _
			                "           else (select ' (' + name + ')' from productversionRelease where id=PDR.ReleaseID)" & _
			                "           end, " & _
                            "targetnotes = case when PDR.ProductDeliverableID is null then pd.targetnotes" & _
			                "              else pdr.targetnotes" & _
			                "              end, " & _
                            "pd.id as ProductDeliverableID, v.ID, r.Name, v.version, v.revision, v.pass, v.vendorversion," & _
                            "v.partnumber,v.irspartnumber, v.modelnumber, v.Location, v.ImagePath, v.tts, e2.name as DevManager, e.name as Developer, pd.imagesummary, " & _
                            "targeted =case when PDR.ProductDeliverableID is null then pd.targeted" & _
			                "          else PDR.targeted" & _
			                "          end, " & _
                            "pd.inimage, " & _
                            "v.languages, r.softpaq, v.softpaqnumbers, v.certificationstatus, r.certrequired, r.categoryid, " & _
                            " v.OEMReadyStatus,v.fccid,v.Anatel,v.ICASA,v.SecondaryRFKill, v.pnpdevices,pvr.Name as ReleaseName "
                
			if request("txtFunction") = "5" then
				rs.Open "SELECT ct.name as CoreTeam, v.changes, v.comments, v.imagepath, v.deliverablespec, v.Replicater, v.rompaq, v.preinstallrom, v.codename, v.binary, v.cab,v.cdimage, v.isoimage, v.ar, v.scriptpaq, v.floppydisk as Diskettepackage, v.preinstall as preinstallpackage, v.propertytabs, v.icondesktop, v.iconinfocenter, v.iconmenu, v.iconpanel, v.icontray, v.sampledate, v.introdate, v.introconfidence, v.samplesconfidence, v.InstallableUpdate, v.packageForWeb,v.levelid, v.filename, v.codename, v.HFCN, v.serviceeoaDate, v.Serviceactive," & _ 
                        "PilotStatus = case when PDR.ProductDeliverableID is null then ps.name" & _
			            "              else (select name from PilotStatus where id=PDR.PilotStatusID)" & _
			            "              end, " & _ 
                        "testdate = case when PDR.ProductDeliverableID is null then pd.testdate" & _
			            "           else pdr.testdate" & _
			            "           end, " & _ 
                        "riskrelease = case when PDR.ProductDeliverableID is null then pd.riskrelease" & _
			            "              else pdr.riskrelease" & _
			            "              end, " & _ 
                        "pilotdate = case when PDR.ProductDeliverableID is null then pd.pilotdate" & _
			            "            else pdr.pilotdate" & _
			            "            end, " & _ 
                        "QualStatus = case when PDR.ProductDeliverableID is null then ts.status" & _
			            "             else (select status from TestStatus where id=PDR.TestStatusID)" & _
			            "             end, " & _ 
                        "v.EndOfLifeDate, v.active," & _ 
                        "developerteststatus = case when PDR.ProductDeliverableID is null then pd.developerteststatus " & _ 
                            "            else Case When pdr.developerteststatus is null then 0 else pdr.developerteststatus end" & _
			                "            end, " & _
                        "wwanteststatus = case when PDR.ProductDeliverableID is null then pd.wwanteststatus" & _
			            "                 else isnull(PDR.wwanteststatus,0)" & _
			            "                 end, " & _ 
                        "odmteststatus = case when PDR.ProductDeliverableID is null then pd.odmteststatus" & _
			            "                else isnull(PDR.odmteststatus,0)" & _
			            "                end, " & _ 
                        "integrationteststatus = case when PDR.ProductDeliverableID is null then pd.integrationteststatus" & _
			            "                        else isnull(PDR.integrationteststatus,0)" & _
			            "                        end, " & _ 
                        "vd.name as Vendor, r.id as RootID, r.TypeID as DelTypeID, c.name as category, pd.SelectiveRestore, pd.ARCD, pd.DRDVD, pd.patch, pd.RACD_Americas, pd.RACD_APD, isnull(pdr.RACD_EMEA,pd.RACD_EMEA) as RACD_EMEA, pd.preinstall, pd.preload, pd.DropInBox, pd.web, pd.OSCD, pd.doccd, r.ID as RootID, pv.typeid, pv.productstatusid, pv.id as ProductID, pv.DOTSName as Product," & _ 
                        "targetnotes = case when PDR.ProductDeliverableID is null then pd.targetnotes" & _
			            "              else isnull(pdr.targetnotes,'')" & _
			            "              end, " & _ 
                        "pd.id as ProductDeliverableID, v.ID, r.Name, v.version, v.revision, v.pass, v.vendorversion, v.partnumber,v.irspartnumber, v.modelnumber, v.Location, v.ImagePath, v.tts, e2.name as DevManager, e.name as Developer, pd.imagesummary," & _ 
                        "targeted =case when PDR.ProductDeliverableID is null then pd.targeted" & _
			            "          else PDR.targeted" & _
			            "          end, " & _
                        "pd.inimage, " & _ 
                        "v.languages, r.softpaq, v.softpaqnumbers, v.certificationstatus, r.certrequired, r.categoryid, v.OEMReadyStatus, v.fccid,v.Anatel,v.ICASA,v.SecondaryRFKill, v.pnpdevices " & strSQl , cn,adOpenForwardOnly
			elseif blnProductsColumns then
			    rs.Open  strFieldNames & strSQl , cn,adOpenForwardOnly
            else
				rs.Open "SELECT distinct v.serviceeoaDate, v.Serviceactive, v.EndOfLifeDate, v.active, vd.name as Vendor, r.TypeID as DelTypeID, c.name as category, r.ID as RootID, v.ID, r.Name, v.version, v.revision, v.pass, v.codename, v.vendorversion, v.partnumber,v.irspartnumber, v.modelnumber, v.Location, v.ImagePath, v.tts, e2.name as DevManager, e.name as Developer, v.languages, r.softpaq, v.softpaqnumbers, v.certificationstatus, r.certrequired, ct.name as coreteam, v.OEMReadyStatus,v.fccid,v.Anatel,v.ICASA,v.SecondaryRFKill, v.pnpdevices,pvr.Name as ReleaseName " & strSQl , cn,adOpenForwardOnly
			end if
            
			if not (rs.EOF and rs.BOF) then
				Response.Write "<TABLE ID=tblResults width=""100%""cellspacing=0 cellpadding=2 bgcolor=ivory border=1>"
				if request("txtFunction") <> "5" then
					if blnProductsColumns then
						PathID = rs("ProductDeliverableID")
					else
						PathID = rs("ID")
					end if
					Response.Write "<THead bgcolor=beige>"
					for i = 0 to Ubound(ColumnArray)						
						if lcase(trim(ColumnArray(i))) = "id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",1,1);"">ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "name" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Name</a></font></TD>"
						end if
                        if trim(request("txtFunction")) = "1" then
                            if lcase(trim(ColumnArray(i))) = "release" then
							    Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Release</a></font></TD>"
						    end if 
                        end if 
						if lcase(trim(ColumnArray(i))) = "version" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Version</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "vendor version" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Vendor&nbsp;Version</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "developer" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Developer</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "dev manager" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Dev&nbsp;Manager</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "workflow" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Workflow</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "product" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Product</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "targeted" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Targeted</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "in image" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">In&nbsp;Image</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "softpaq" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Softpaq</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "softpaq numbers" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Softpaq&nbsp;Numbers</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "whql" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">WHQL</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "oem ready" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">OEM Ready</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "part number" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Part&nbsp;Number</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "irs part number" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">IRS&nbsp;Part&nbsp;Number</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "path" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Path</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "model" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Model</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "tts" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">TTS</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "device id string" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Device&nbsp;ID&nbsp;String</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "device id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Device&nbsp;ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "vendor id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Vendor&nbsp;ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "subsys ven id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Subsys&nbsp;Ven&nbsp;ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "subsys dev id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Subsys&nbsp;Dev&nbsp;ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "fcc id" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">FCC&nbsp;ID</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "anatel" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Anatel</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "icasa" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">ICASA</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "rf kill mechanism" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">RF&nbsp;Kill&nbsp;Mechanism</a></font></TD>"
						end if
                        
						if lcase(trim(ColumnArray(i))) = "vendor" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Vendor</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "mit signoff" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">MIT&nbsp;Signoff</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "code name" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Code&nbsp;Name</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "odm signoff" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">ODM&nbsp;Signoff</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "wwan signoff" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">WWAN&nbsp;Signoff</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "dev signoff" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Dev&nbsp;Signoff</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "wwan samples" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">WWAN&nbsp;Samples</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "wwan notes" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">WWAN&nbsp;Notes</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "odm samples" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">ODM&nbsp;Samples</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "odm notes" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">ODM&nbsp;Notes</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "mit samples" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">MIT&nbsp;Samples</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "mit notes" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">MIT&nbsp;Notes</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "factory eoa" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">Factory&nbsp;EOA</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "software eol" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">Software&nbsp;EOL</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "service eoa" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">Service&nbsp;EOA</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "hw qual status" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">HW&nbsp;Qual&nbsp;Status</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "pilot status" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",4,1);"">Pilot&nbsp;Status</a></font></TD>"
						end if

						if lcase(trim(ColumnArray(i))) = "hw version" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">HW&nbsp;Version</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "fw version" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">FW&nbsp;Version</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "hw rev" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">HW&nbsp;Rev</a></font></TD>"
						end if

						if lcase(trim(ColumnArray(i))) = "category" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Category</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "core team" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Core&nbsp;Team</a></font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "deliverable root name" then
							Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', " & i & ",0,1);"">Deliverable&nbsp;Root&nbsp;Name</a></font></TD>"
						end if
					next
                    if request("cboFormat") = "1" or request("cboFormat") = "2" then
                        response.write "</thead><tbody>"
                    else
					    Response.Write "<TD style=""Display:none"">Path</TD>"
					    Response.Write "</THead><TBODY LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()"">"
                    end if
				end if
			else
				Response.Write "<br><font size=2 face=verdana>No items match your query criteria</font>"
			end if
			RowCount = 0
			strlastroot=""
			do while not rs.EOF
			   ' response.Flush
    			if RowCount >=5000 and  request("txtFunction") = "1" then
					Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Summary Reports are limited to 5000 records.</b><BR><BR></font>"
					exit do
				elseif RowCount >=5000 and  request("txtFunction") <> "1" then
					Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Detailed Reports are limited to 5000 records.</b><BR><BR></font>"
					exit do
				end if

				RowCount = RowCount + 1
				
				strVersion = rs("Version") & ""
				if rs("revision") & "" <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if rs("pass") & "" <> "" then
					strVersion = strVersion & "," & rs("pass")
				end if

				if trim(rs("VendorVersion") & "") = "" then
					strVendorVersion = "&nbsp;"
				else
					strVendorVersion = rs("VendorVersion") & ""
				end if

				if rs("Softpaq") then
					strSoftpaq = "Yes"
				else
					strSoftpaq = "&nbsp;"
				end if
				if blnProductsColumns then
					if rs("Targeted") then
						strTargeted = "Yes"
					else
						strTargeted = "&nbsp;"
					end if
					if rs("InImage") then
						strInImage = "Yes"
					else	
						strInImage = "&nbsp;"
					end if
				end if 
				if rs("Languages") & "" = "" or rs("Languages") & "" = "<Language Independent>" then
					strlanguage = "N/A"
				else
					strlanguage =  rs("Languages") & ""
				end if
				
				if rs("Certrequired") then
					select case trim(rs("CertificationStatus") & "")
					case "0",""
						strCertificationStatus = "Required"
					case "1"
						strCertificationStatus = "Submitted"
					case "2"
						strCertificationStatus = "Approved"
					case "3"
						strCertificationStatus = "Failed"
					case "4"
						strCertificationStatus = "Waiver"
					case else
						strCertificationStatus = rs("CertificationStatus") & "&nbsp;"
					end select		
				else	
					strCertificationStatus = "&nbsp;"
				end if
				
				if rs("OEMReadyStatus") then
				    select case trim(rs("OEMReadyStatus") & "")
					case "0",""
						strOEMReadyStatus = "Required"
					case "1"
						strOEMReadyStatus = "Submitted"
					case "2"
						strOEMReadyStatus = "Approved"
					case "3"
						strOEMReadyStatus = "Failed"
					case else
						strOEMReadyStatus = rs("OEMReadyStatus") & "&nbsp;"
					end select		
				else	
					strOEMReadyStatus = "&nbsp;"
				end if

				strDistribution = ""
				if request("txtFunction") = "5" then
					if rs("Preinstall") then
						strDistribution = ", Preinstall"
					end if
					if rs("Preload") then
						strDistribution = strDistribution & ", Preload"
					end if
					if rs("DropInBox") then
						strDistribution = strDistribution & ", DIB"
					end if
					if rs("Web") then
						strDistribution = strDistribution & ", Web"
					end if
					if rs("SelectiveRestore") then
						strDistribution = strDistribution & ", Selective Restore"
					end if
					if rs("ARCD") then
						strDistribution = strDistribution & ", DRCD"
					end if
					if rs("DRDVD") then
						strDistribution = strDistribution & ", DRDVD"
					end if
					if rs("RACD_Americas") then
						strDistribution = strDistribution & ", RACD-Americas"
					end if
					if rs("RACD_APD") then
						strDistribution = strDistribution & ", RACD-APD"
					end if
					if rs("RACD_EMEA") then
						strDistribution = strDistribution & ", RACD-EMEA"
					end if
					if rs("OSCD") then
						strDistribution = strDistribution & ", OSCD"
					end if
					if rs("DocCD") then
						strDistribution = strDistribution & ", DocCD"
					end if
					if trim(rs("Patch")&"") <> "0" then
						strDistribution = strDistribution & ", Patch"
					end if
					
					if strDistribution <> "" then
						strDistribution = mid(strDistribution,3)
					else
						strDistribution = "&nbsp;"
					end if
				end if

				if trim(rs("DelTypeID")) = "1" then
					strDelname = rs("Vendor") & " " & rs("Name")
				else
					strDelName =  rs("Name")
				end if

				if request("txtFunction") = "1" then
					Response.Write  "<TR LANGUAGE=javascript onclick=""return row_onclick(" & PathID & "," & rs("RootID") & "," & rs("ID") & ")"">"
					for i = 0 to Ubound(ColumnArray)						
						if lcase(trim(ColumnArray(i))) = "id" then
							Response.Write   "<TD valign=top>" & rs("ID") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "name" then
							Response.Write   "<TD valign=top nowrap>" & strDelName & "</font></TD>"
						end if
                        if lcase(trim(ColumnArray(i))) = "release" then
							Response.Write   "<TD valign=top nowrap>" & rs("ReleaseName") & "</font></TD>"
						end if

						if lcase(trim(ColumnArray(i))) = "version" then
							Response.Write   "<TD valign=top nowrap>" & strVersion & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "vendor version" then
							Response.Write  "<TD valign=top >" & strVendorVersion & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "developer" then
							Response.Write  "<TD valign=top nowrap>" & rs("Developer") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "dev manager" then
							Response.Write  "<TD valign=top nowrap>" & rs("DevManager") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "workflow" then
							Response.Write "<TD valign=top nowrap>" & rs("Location") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "product" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("Product") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "targeted" then
							Response.Write "<TD valign=top nowrap align=center>" & strTargeted & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "in image" then
							Response.Write  "<TD valign=top nowrap align=center>" & strInImage & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "softpaq" then
							Response.Write "<TD valign=top nowrap align=center>" & strSoftpaq & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "softpaq numbers" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("SoftpaqNumbers") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "whql" then
							Response.Write "<TD valign=top nowrap align=center>" & strCertificationStatus & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "oem ready" then
							Response.Write "<TD valign=top nowrap align=center>" & strOEMReadyStatus & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "part number" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("PartNumber") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "irs part number" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("IRSPartNumber") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "path" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("imagepath") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "model" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("ModelNumber") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "fcc id" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("fccid") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "vendor id" then
                            if trim(rs("pnpdevices") & "") = "" or (instr(lcase(rs("pnpdevices")) & "","ven_") = 0 and instr(lcase(rs("pnpdevices")) & "","vid_") = 0) then
                                Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
                            else
                                Response.Write "<TD valign=top nowrap align=left>"

                                LineArray = split(lcase(rs("pnpdevices")),vbcrlf)

                                for each strLine in LineArray
                                    if instr(lcase(strLine) & "","ven_") > 0 then
                                        strTag = "ven_"
                                    else
                                        strTag = "vid_"
                                    end if

                                    if trim(strLine) <> "" then
                                        if instr(strLine,strTag) > 0 then
                                            response.write ucase(mid(strLine,instr(lcase(strLine),strTag)+4,4)) & "&nbsp;<BR>"
                                        else
                                            response.write "<BR>"
                                        end if
                                    end if
                                next

                                response.write "</font></TD>"
                            end if
						end if

						if lcase(trim(ColumnArray(i))) = "device id" then
                            if trim(rs("pnpdevices") & "") = "" or (instr(lcase(rs("pnpdevices")) & "","dev_") = 0 and instr(lcase(rs("pnpdevices")) & "","pid_") = 0) then
                                Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
                            else
                                Response.Write "<TD valign=top nowrap align=left>"

                                LineArray = split(lcase(rs("pnpdevices")),vbcrlf)
                                for each strLine in LineArray
                                    if trim(strLine) <> "" then
                                        if instr(lcase(strLine) & "","pid_") > 0 then
                                            strTag = "pid_"
                                        else
                                            strTag = "dev_"
                                        end if

                                        if instr(strLine,strTag) > 0 then
                                            response.write ucase(mid(strLine,instr(lcase(strLine),strTag)+4,4)) & "&nbsp;<BR>"
                                        else
                                            response.write "<BR>"
                                        end if
                                    end if
                                next

                                response.write "</font></TD>"
                            end if
						end if

						if lcase(trim(ColumnArray(i))) = "subsys ven id" then
                            if trim(rs("pnpdevices") & "") = "" or instr(lcase(rs("pnpdevices")) & "","subsys_") = 0 then
                                Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
                            else
                                Response.Write "<TD valign=top nowrap align=left>"

                                LineArray = split(lcase(rs("pnpdevices")),vbcrlf)
                                for each strLine in LineArray
                                    if trim(strLine) <> "" then
                                        if instr(strLine,"subsys_") > 0 then
                                            response.write ucase(mid(strLine,instr(lcase(strLine),"subsys_")+11,4)) & "&nbsp;<BR>"
                                        else
                                            response.write "<BR>"
                                        end if
                                    end if
                                next

                                response.write "</font></TD>"
                            end if
						end if


						if lcase(trim(ColumnArray(i))) = "subsys dev id" then
                            if trim(rs("pnpdevices") & "") = "" or instr(lcase(rs("pnpdevices")) & "","subsys_") = 0 then
                                Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
                            else
                                Response.Write "<TD valign=top nowrap align=left>"

                                LineArray = split(lcase(rs("pnpdevices")),vbcrlf)
                                for each strLine in LineArray
                                    if trim(strLine) <> "" then
                                        if instr(strLine,"subsys_") > 0 then
                                            response.write ucase(mid(strLine,instr(lcase(strLine),"subsys_")+7,4)) & "&nbsp;<BR>"
                                        else
                                            response.write "<BR>"
                                        end if
                                    end if
                                next

                                response.write "</font></TD>"
                            end if
						end if

						if lcase(trim(ColumnArray(i))) = "device id string" then
                            if trim(rs("pnpdevices") & "") = "" then
                                Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
                            else
                                Response.Write "<TD valign=top nowrap align=left>"

                                LineArray = split(rs("pnpdevices"),vbcrlf)
                                for each strLine in LineArray
                                    if trim(strLine) <> "" then
                                        response.write server.HTMLEncode(strLine) & "<BR>"
                                    end if
                                next

                                response.write "</font></TD>"
                            end if
						end if
						if lcase(trim(ColumnArray(i))) = "anatel" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("anatel") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "icasa" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("icasa") & "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "rf kill mechanism" then
							Response.Write "<TD valign=top nowrap align=left>"
                            if trim(rs("SecondaryRFKill") & "") = "" then
                                response.write "0 Hardware Pin"
                            elseif rs("SecondaryRFKill") then
                                response.write "Discrete Hardware Pins"
                            else
                                response.write "1 Hardware Pin"
                            end if
                            response.write "&nbsp;</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "tts" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & rs("TTS") & "&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "vendor" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("Vendor") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "hw qual status" then
							if rs("DelTypeID") = 1 then
								if rs("QualStatus") = "Date" then
									Response.Write "<TD valign=top nowrap align=left>" & rs("TestDate") & "</font></TD>"
								elseif rs("QualStatus") = "QComplete" and trim(rs("RiskRelease") & "") = "1" then
									Response.Write "<TD valign=top nowrap align=left>Risk Release</font></TD>"
								elseif trim(rs("QualStatus") & "") = "" then
									Response.Write "<TD valign=top nowrap align=left>Not Used</font></TD>"
								else
									Response.Write "<TD valign=top nowrap align=left>" & rs("QualStatus") & "</font></TD>"
								end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "pilot status" then
							if rs("DelTypeID") = 1 then
								if lcase(trim(rs("PilotStatus") & "")) = "p_scheduled" then
									Response.Write "<TD valign=top nowrap align=left>" & rs("PilotDate") & "</font></TD>"
								elseif trim(rs("QualStatus") & "") = "" and lcase(trim(rs("PilotStatus") & "")) = "p_planning" then
									Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
								else
									Response.Write "<TD valign=top nowrap align=left>" & rs("PilotStatus") & "</font></TD>"								
								end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "code name" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & rs("CodeName") & "&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "hw version" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & rs("Version") & "&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "fw version" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & rs("Revision") & "&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "hw rev" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & rs("Pass") & "&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "mit signoff" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & teststatusarray(rs("integrationteststatus")) & "</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "odm signoff" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>" & teststatusarray(rs("odmteststatus")) & "</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "wwan signoff" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>" & teststatusarray(rs("wwanteststatus")) & "</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "wwan notes" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							else
								Response.Write "<TD valign=top align=left>" & rs("wwantestnotes") & "&nbsp;</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "odm notes" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							else
								Response.Write "<TD valign=top align=left>" & rs("odmtestnotes") & "&nbsp;</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "mit notes" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							else
								Response.Write "<TD valign=top align=left>" & rs("mittestnotes") & "&nbsp;</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "dev signoff" then
							if rs("DelTypeID") = 1 then
                                if rs("developerteststatus") <> "" then
								    Response.Write "<TD valign=top nowrap align=left>" & Developerstatusarray(rs("developerteststatus")) & "</TD>"
                                else 
                                    Response.Write "<TD valign=top nowrap align=left>&nbsp;</TD>"
						        end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "mit samples" then
							if rs("DelTypeID") = 1 then
								if trim(rs("MITSamples") & "") = "" then
									Response.Write "<TD valign=top nowrap align=left>0&nbsp;</font></TD>"
								else	
									Response.Write "<TD valign=top nowrap align=left>" & rs("mitsamples") & "&nbsp;</font></TD>"
								end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "odm samples" then
							if rs("DelTypeID") = 1 then
								if trim(rs("ODMSamples") & "") = "" then
									Response.Write "<TD valign=top nowrap align=left>0&nbsp;</font></TD>"
								else	
									Response.Write "<TD valign=top nowrap align=left>" & rs("odmsamples") & "&nbsp;</font></TD>"
								end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "wwan samples" then
							if rs("DelTypeID") = 1 then
								if trim(rs("WWANSamples") & "") = "" then
									Response.Write "<TD valign=top nowrap align=left>0&nbsp;</font></TD>"
								else	
									Response.Write "<TD valign=top nowrap align=left>" & rs("wwansamples") & "&nbsp;</font></TD>"
								end if
							else
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "factory eoa" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							elseif (not rs("active")) and not(rs("Serviceactive"))  then
								Response.Write "<TD valign=top nowrap align=left>Unavailable</font></TD>"
							elseif (not rs("active")) then
								Response.Write "<TD valign=top nowrap align=left>Service Only</font></TD>"
							elseif isnull(rs("endoflifedate")) then
								Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>" & rs("endoflifedate") & "</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "software eol" then
							if rs("DelTypeID") = 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							elseif not rs("active") then
								Response.Write "<TD valign=top nowrap align=left>EOL</font></TD>"
							elseif isnull(rs("endoflifedate")) then
								Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>" & rs("endoflifedate") & "</font></TD>"
							end if
						end if

						if lcase(trim(ColumnArray(i))) = "service eoa" then
							if rs("DelTypeID") <> 1 then
								Response.Write "<TD valign=top nowrap align=left>N/A</font></TD>"
							elseif not rs("Serviceactive") then
								Response.Write "<TD valign=top nowrap align=left>Unavailable</font></TD>"
							elseif isnull(rs("serviceeoadate")) then
								Response.Write "<TD valign=top nowrap align=left>&nbsp;</font></TD>"
							else
								Response.Write "<TD valign=top nowrap align=left>" & rs("serviceeoaDate") & "</font></TD>"
							end if
						end if
						if lcase(trim(ColumnArray(i))) = "category" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("Category") & "</font></TD>"
						end if
						if lcase(trim(ColumnArray(i))) = "core team" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("CoreTeam") & "</font></TD>"
						end if
                        if lcase(trim(ColumnArray(i))) = "deliverable root name" then
							Response.Write "<TD valign=top nowrap align=left>" & rs("Name") & "</font></TD>"
						end if
					next

					'Response.Write "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & strlanguage & "</font></TD>"
                    if request("cboFormat") <> "1" and request("cboFormat") <> "2" then

                        if blnProductsColumns then
						    Response.Write "<TD ID=PATH" & trim(rs("ProductDeliverableID")) & " style=""Display:none"">" & rs("ImagePath") & "</TD>"
					    else
    						Response.Write "<TD ID=PATH" & trim(rs("ID")) & " style=""Display:none"">" & rs("ImagePath") & "</TD>"
					    end if
					end if
                    Response.Write  "</TR>"
				elseif request("txtFunction") = "5" then

					if trim(strlastroot) <> trim(rs("ID")) then

						if trim(strLastroot) <> "" then

							Response.Write "<TR><TD colspan=3>" & strProductTable & "</TD></TR></Table><BR><BR><BR>"

						end if

						Response.write "<font size=2 face=verdana><b>" & strDelName & " [" & strVersion & "] " & "Details</b></font><BR>"

						%>						

						<TABLE bordercolor=black cellspacing=0 Border=1 width="100%">

						<TR>

						<% if trim(rs("DelTypeID")) = "1"  then %>

							<TD valign=top><TABLE width="100%" >

							<TR><TD><font face=verdana size=1><b>ID:</b></font></TD><TD><font face=verdana size=1><a target=_blank href="../WizardFrames.asp?Type=1&ID=<%=rs("ID")%>"><%=rs("ID")%></a></font></TD></TR>

							<TR><TD><font face=verdana size=1><b>Hardware Version:</b></font></TD><TD><font face=verdana size=1><%=rs("Version") & "&nbsp;"%></font></TD></TR>

							<TR><TD nowrap><font face=verdana size=1><b>Firmware Version:</b></font></TD><TD><font face=verdana size=1><%=rs("Revision") & "&nbsp;"%></font></TD></TR>

							<TR><TD nowrap><font face=verdana size=1><b>Rev:</b></font></TD><TD><font face=verdana size=1><%=rs("Pass") & "&nbsp;"%></font></TD></TR>

							<%if trim(rs("DelTypeID"))= "1"  or (trim(rs("CategoryID")&"") <> "161" and rs("HFCN"))then%>

								<TR><TD><b>HFCN:</b></TD><TD><%=replace(replace(rs("HFCN") & "","True","Yes"),"False","No")%></TD><TR>

							<%end if%>

						</table></TD>

						<% elseif trim(rs("DelTypeID")) = "3" then %>

							<TD valign=top><TABLE width="100%">

							<TR><TD><font face=verdana size=1><b>ID:</b></font></TD><TD><font face=verdana size=1><a target=_blank href="../WizardFrames.asp?Type=1&ID=<%=rs("ID")%>"><%=rs("ID")%></a></font></TD></TR>

							<TR><TD><font face=verdana size=1><b>Version:</b></font></TD><TD><font face=verdana size=1><%=rs("Version") & "&nbsp;"%></font></TD></TR>

							<%if trim(rs("DelTypeID"))= "1"  or (trim(rs("CategoryID")&"") <> "161" and rs("HFCN"))then%>

								<TR><TD><b>HFCN:</b></TD><TD><%=replace(replace(rs("HFCN") & "","True","Yes"),"False","No")%></TD><TR>

							<%end if%>

						</table></TD>

						<% else %>

							<TD valign=top><TABLE width="100%" >

							<TR><TD><font face=verdana size=1><b>ID:</b></font></TD><TD><font face=verdana size=1><a target=_blank href="../WizardFrames.asp?Type=1&ID=<%=rs("ID")%>"><%=rs("ID")%></a></font></TD></TR>

							<TR><TD><font face=verdana size=1><b>Version:</b></font></TD><TD><font face=verdana size=1><%=rs("Version") & "&nbsp;"%></font></TD></TR>

							<TR><TD nowrap><font face=verdana size=1><b>Revision:</b></font></TD><TD><font face=verdana size=1><%=rs("Revision") & "&nbsp;"%></font></TD></TR>

							<TR><TD><font face=verdana size=1><b>Pass:</b></font></TD><TD><font face=verdana size=1><%=rs("Pass") & "&nbsp;"%></font></TD></TR>

						</table></TD>

						<% end if %>

						<TD valign=top><TABLE width="100%" ><TR><TD nowrap><font face=verdana size=1><b>Vendor:</b></font></TD><TD><font face=verdana size=1><%=rs("Vendor")& "&nbsp;"%></font></TD></TR>

						<TR><TD nowrap><font face=verdana size=1><b>Vendor Version:</b></font></TD><TD><font face=verdana size=1><%=rs("VendorVersion") & "&nbsp;"%></font></TD></TR>

						<% if trim(rs("DelTypeID"))= "1"  then %>

								<TR><TD nowrap><font face=verdana size=1><b>HP Part Number:</b></font></TD><TD><font face=verdana size=1><%=rs("PartNumber") & "&nbsp;"%></font></TD></TR>

								<TR><TD nowrap><font face=verdana size=1><b>Model Number:</b></font></TD><TD><font face=verdana size=1><%=rs("ModelNumber") & "&nbsp;"%></font></TD></TR>

							</table></TD>

						<% else %>

							<TR><TD nowrap><font face=verdana size=1><b>FileName:</b></font></TD><TD><font face=verdana size=1><%=rs("Filename") & "&nbsp;"%></font></TD></TR>

							</table></TD>

						<% end if %>

						<TD valign=top><TABLE width="100%" >

		

						<% if trim(rs("DelTypeID"))= "1" and trim(rs("CodeName") & "") <> "" then %>

								<TR><TD nowrap><font face=verdana size=1><b>Code Name:</b></font></TD><TD><font face=verdana size=1><%=rs("CodeName") & "&nbsp;"%></font></TD></TR>

						<% end if %>

						

						<TR><TD nowrap><font face=verdana size=1><b>Dev. Manager:</b></font></TD><TD><font face=verdana size=1><%=rs("DevManager") & "&nbsp;"%></font></TD></TR>

						<TR><TD nowrap><font face=verdana size=1><b>Developer:</b></font></TD><TD><font face=verdana size=1><%=rs("Developer") & "&nbsp;"%></font></TD></TR>

						<TR>

						<%if trim(rs("DelTypeID"))= "1"  then%>

							<td nowrap><font face=verdana size=1><b>Production&nbsp;Level:</b></font></td>    

						<%else%>

							<td nowrap><font face=verdana size=1><b>Build&nbsp;Level:</b></font></td>    

						<%end if%>

		

							<%

							set rs2 = server.CreateObject("ADODB.recordset")



							if trim(rs("DelTypeID"))= "1"  then

								rs2.Open "spListDeliverableLevels 1" ,cn,adOpenForwardOnly

							else

								rs2.Open "spListDeliverableLevels 2" ,cn,adOpenForwardOnly

							end if

			

							do while not rs2.EOF

								if trim(rs("LevelID") & "") = trim(rs2("ID") & "") then

									strLevel = rs2("name")

									exit do

								end if

								rs2.MoveNext

							loop

							rs2.Close

							set rs2 = nothing

							%>

						

							<TD><font face=verdana size=1><%=strLevel%></font></TD></TR>

						</table></TD>

						</TR>



				<TR><TD colspan=3><TABLE width="100%">



				<%

				Count = 0

				if trim(rs("ID") & "") <> "" then

					set rs2 = server.CreateObject("ADODB.recordset")

					strSQL = "spGetDelMilestoneList " & rs("RootID") & "," & rs("ID")

					rs2.Open strSQL,cn,adOpenForwardOnly

					do while not rs2.EOF

						strActualDate = rs2("Actual") & ""

						if strActualDate = "" then

							strActualDate = "&nbsp;"

						elseif instr(strActualDate," ") > 0 then

							strActualDate = left(strActualDate,instr(strActualDate," ") - 1)

						end if

		

						strMilestone = rs2("Milestone")

						strStatus = rs2("Status") 

						strPlanned = rs2("Planned")

		

					%>

				

					<% if Count = 0 then %>

						<TR><TD align=center nowrap><font face=verdana size=1><u><b>Workflow Step</b></u><BR>

					<% else %>

						<TR><TD align=center nowrap><font face=verdana size=1>

					<% end if %>

				

					<%=strMilestone %></font></TD>

		

					<% if Count = 0 then %>

						<TD align=center nowrap><font face=verdana size=1><u><b>Status</b></u><BR>

					<% else %>

						<TD align=center nowrap><font face=verdana size=1>

					<% end if %>

					<%=strStatus %></font></TD>

		

					<% if Count = 0 then %>

						<TD align=center nowrap ><font face=verdana size=1><u><b>Planned Date</b></u><BR>

					<% else %>

						<TD align=center nowrap><font face=verdana size=1>

					<% end if %>

					<%=strPlanned %></font></TD>



					<% if Count = 0 then %>

						<TD align=center nowrap><font face=verdana size=1><u><b>Actual Date</b></u><BR>

					<% else %>

						<TD align=center nowrap><font face=verdana size=1>

					<% end if %>

					<%=strActualDate %></font></TD>

		

					<% Count = Count + 1 

				

						rs2.MoveNext

					loop

					rs2.Close

					set rs2=nothing

				end if

	%>



				</TR></td>

				</TABLE></td></tr>







		    <%if trim(rs("DelTypeID"))= "1"  then%>

				<TR><TD colspan=3><TABLE width="100%" >

			<%else%>

				<TR style="Display:none"><TD colspan=3><TABLE width="100%" >

			<%end if%>

			<td nowrap><font face=verdana size=1><b>Samples&nbsp;Available&nbsp;Date:</b>&nbsp;&nbsp;&nbsp;

			<%if trim(rs("SampleDate")& "") <> "" then %>

				<%=rs("SampleDate")%>

			<% else %>

				Unknown&nbsp;

			<% end if %>

			&nbsp;&nbsp;&nbsp;<b>Confidence:</b>&nbsp;</font>

			<%if trim(rs("SamplesConfidence")) = "1" then%>
				<font face=verdana size=1 color=green>High</font>
			<%elseif trim(rs("SamplesConfidence")) = "2" then%>
				<font face=verdana size=1 color=black>Medium</font>
			<%elseif trim(rs("SamplesConfidence")) = "3" then%>
				<font face=verdana size=1 color=red>Low</font>
			<%else%>
				<font face=verdana size=1>Unknown&nbsp;</FONT>
			<%end if %>
			</td>
			</TABLE></td></tr>

	

			<%if trim(rs("IntroDate") & "") <> "" then%>

				<tr><TD colspan=3><TABLE width="100%" >

				<%if trim(rs("DelTypeID"))= "1" then%>

					<td nowrap><font face=verdana size=1><b>Mass&nbsp;Production:</b>&nbsp;&nbsp;

				<%else%>

					<td nowrap><font face=verdana size=1><b>Intro&nbsp;Date:</b>&nbsp;&nbsp;

				<%end if%>

				

				<%=strIntroDate %>



				&nbsp;&nbsp;<b>Confidence:</b>&nbsp;</font>

				<%if trim(rs("IntroConfidence") & "") = "1" then%>
					<font face=verdana size=1 color=green>High</font>
				<%elseif trim(rs("IntroConfidence") & "") = "2" then%>
					<font face=verdana size=1 color=black>Medium</font>
				<%elseif trim(rs("IntroConfidence") & "") = "3" then%>
					<font face=verdana size=1 color=red>Low</font>
				<%else%>
					<font face=verdana size=1>Unknown&nbsp;</FONT>
				<%end if%>
				</td>

			</TABLE></TD></tr>

		<%end if%>



	<% if trim(rs("EndOfLifeDate") & "") <> "" or (not rs("Active")) then %>

	<TR><TD colspan=3><TABLE width="100%" ><tr>

		<td nowrap><font face=verdana size=1><b>Available&nbsp;Until:</b>&nbsp;&nbsp;&nbsp;&nbsp;

		<%=rs("EndOfLifeDate")%>

			</FONT>	

		<% if not rs("Active") then %> 

			&nbsp;&nbsp;&nbsp;&nbsp;<font size=1 face=verdana>This version is End of Life.</font>

		<% end if %>

		</td>

	</TABLE></td></tr>

	<%end if%>    

  

    <% if trim(rs("InstallableUpdate") & "") <> "0" or trim(rs("packageForWeb") & "") <> "0" then %> 

		<%if left(rs("Filename") & "",5) <> "HFCN_" and  trim(rs("DelTypeID"))<> "1" then%>

			<TR><TD colspan=3><TABLE width="100%" >

		<%else%>

			<TR  style="Display:none"><TD colspan=3><TABLE width="100%" >

		<%end if%>

		<td nowrap><font face=verdana size=1><b>Special Notes:</b>&nbsp;&nbsp;

		<% if trim(rs("InstallableUpdate") & "") <> "0" then %> 

			&nbsp;&nbsp;&nbsp;&nbsp;Installable&nbsp;Update,

		<% end if %>

		<% if trim(rs("packageForWeb") & "") <> "0" then %> 

			&nbsp;&nbsp;Package&nbsp;For&nbsp;Web

		<% end if %>

		</font></td>

		</TABLE></td></tr>

	<%end if%>

	

	<%if trim(rs("DelTypeID"))= "2" then%>

		<tr><TD colspan=3><TABLE width="100%">

		<td nowrap><font face=verdana size=1><b>Packaging:</b>&nbsp;&nbsp;

		<% strTemp = "" 

			if rs("preinstallpackage") & "" = "1" then

				strTemp = strTemp & ",Preinstall"

			end if

			if rs("Diskettepackage") & "" = "1" then

				strTemp = strTemp & ",Diskette"

			end if

			if rs("scriptpaq") then

				strTemp = strTemp & ",Scriptpaq"

			end if

			if rs("cdimage") & "" = "1" then

				strTemp = strTemp & ",CD Files"

			end if

			if rs("isoImage") & "" = "1" then

				strTemp = strTemp & ",ISO Image"

			end if

			if rs("ar") & "" = "1" then

				strTemp = strTemp & ",Replicater Only"

				if trim(rs("Replicater") & "") <> "" then

					strTemp = strTemp & " (" & rs("Replicater") & ")"

				end if

			end if

			if strTemp <> "" then

				strTemp = mid(strTemp,2)

			end if



		%>

		<%=strTemp%>

		</font></td>

		</TABLE></td></tr>	

	<%elseif trim(rs("DelTypeID"))= "3" then%> 

		<tr><TD colspan=3><TABLE width="100%">

		<td nowrap><font face=verdana size=1><b>ROM Components:</b>

		<% strTemp = "" 

			if rs("binary") & "" = "1" then

				strTemp = strTemp & ",Binary"

			end if

			if rs("rompaq") & "" = "1" then

				strTemp = strTemp & ",Rompaq"

			end if

			if rs("preinstallrom") then

				strTemp = strTemp & ",Preinstall"

			end if

			if rs("cab") & "" = "1" then

				strTemp = strTemp & ",CAB"

			end if

			if strTemp <> "" then

				strTemp = mid(strTemp,2)

			end if



		%>

		<%=strTemp%>	

	

	</font></td>

	</TABLE></td></tr>	



	<%end if%>	

	<%if rs("icondesktop") or rs("iconmenu") or rs("icontray") or rs("iconpanel") or rs("iconinfocenter")  then %>



			<TR><TD colspan=3><TABLE width="100%">

			<td nowrap><font face=verdana size=1><b>Icons Installed:</b>&nbsp;&nbsp;&nbsp;

			<%

			strIcons = ""

			if rs("icondesktop") then 

				strIcons = strIcons & ", " & "Desktop"

			end if 

			if rs("iconmenu") then 

				strIcons = strIcons & ", " & "Start Menu"	

			end if 

			if rs("icontray") then

				strIcons = strIcons & ", " & "System Tray"

			end if

			if rs("iconpanel")   then

				strIcons = strIcons & ", " & "Control Panel"

			end if

			if rs("iconinfocenter") then

				strIcons = strIcons & ", " & "Info Center"

			end if

			if strIcons <> "" then

				strIcons = mid(strIcons,2)

			end if

			Response.Write strIcons

			%></font></td>

			</TABLE></td></tr>

	<%end if%>



	<%if trim(rs("PropertyTabs") & "") <> "" and trim(rs("DelTypeID")) <> "1" then%>

		<TR><TD colspan=3><TABLE width="100%" >

		<td nowrap><font face=verdana size=1><b>Property Tabs Added:</b>&nbsp;&nbsp;&nbsp;

		<%=replace(trim(rs("PropertyTabs") & ""),"""","&quot;")%>

		</font></td>

		</TABLE></td></tr>

	<%end if%>

  

	<%if trim(rs("DeliverableSpec") & "") <> "" then%>

		<tr><TD colspan=3><TABLE width="100%">

		<td nowrap><font face=verdana size=1><b>Functional&nbsp;Spec:</b>&nbsp;&nbsp;

		<%=rs("DeliverableSpec") & ""%>

	</font></td>

	</TABLE></td></tr>

	<%end if%>



	<%if trim(rs("ImagePath") & "") <> "" then%>

		<TR><TD colspan=3><TABLE width="100%">

		<td nowrap><font face=verdana size=1><b>Location/Path:</b>&nbsp;&nbsp;&nbsp; 

			<% if left(rs("ImagePath") & "",2) = "\\" then%>

				<a target=_blank href="<%=rs("ImagePath") & ""%>"><%=rs("ImagePath") & ""%></a>

			<%else%>

				<%=rs("ImagePath") & ""%>

			<%end if%>

	</font></td>

	</TABLE></td></tr>

	<%end if%>



	<%if trim(rs("Changes") & "") <> "" then%>

	<TR><TD colspan=3><TABLE width="100%">

	<td><font face=verdana size=1><b>Modifications, Enhancements, or Reason for Release:</b>&nbsp;&nbsp; 

		<BR><%=replace(rs("Changes"),vbcrlf,"<BR>")%>

	</font></td>

	</TABLE></td></tr>

	<%end if%>

	

	<%if trim(rs("Comments") & "") <> "" then%>

	<TR><TD colspan=3><TABLE width="100%">

	<td><font face=verdana size=1><b>Comments:</b>&nbsp;&nbsp; 

    	<%=rs("Comments") & ""%>&nbsp;&nbsp;

	</font></td>

	</TABLE></td></tr>

	<%end if%>







<%						

						strlastroot = trim(rs("ID"))

						strProductTable = ""

					end if

					strProductTable = strProductTable & "," & rs("Product")
				end if

				rs.MoveNext
			loop

			Response.Write "</Table><BR>"

			Response.Write   "<BR><font size=1 face=verdana>Deliverables Displayed: " & RowCount & "</font><BR>"

			rowcount=0

			rs.Close

		elseif request("txtFunction") = "6" then
			RowCount= 0
			Response.Write "<font size=2 face=verdana>"
			rs.Open "Select distinct  e.name as Developer, r.id as RootID, v.id, r.name, v.version,v.revision, v.pass, v.location, v.actualreleasedate " & strSQL & " order by r.name, v.id desc",cn,adOpenStatic
			strLastRoot = ""
			ReleaseCount = 0
			FailCount=0
			InTestCount=0
			OTSFixed = 0
			OTSFound = 0
			do while not rs.EOF
				RowCount = RowCount + 1
				if strLastRoot <>  rs("name") then
					if strLastRoot <> "" then
						Response.Write "</table>"
						Response.Write "<font size=1><Table><tr>"
						Response.Write "<td>Versions Released:</td><td>" & ReleaseCount & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
						Response.Write "<td>Versions In Test:</td><td>" & InTestCount & "</TD>"
						Response.Write "<td>OTS Fixed:</td><td>" & OTSFixed & "</td></tr>"
						Response.Write "<tr><td>Versions Failed:</td><td>" & FailCount & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
						Response.Write "<td>OTS Found:</td><td>" & OTSFound & "</TD>"
						Response.Write "</tr></table></font><BR><BR>"
						ReleaseCount = 0
						FailCount=0
						InTestCount=0
						OTSFixed = 0
						OTSFound = 0
						'Response.Flush
						if RowCount >=150 then
							Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Release Reports are limited to 150 records.</b><BR><BR></font>"
							exit do
						end if

					end if
					Response.write "<b>" & rs("name") & "</b><BR>"
					Response.Write "<Table width=600 bgcolor=ivory border=1 cellpadding=2 cellspacing=0><TR bgcolor=beige><TD>ID</TD><TD>Version</TD><TD>Released</TD><TD>OTS Fixed</TD><TD>OTS Found</TD><TD>Developer</TD></TR>"
					strLastRoot =  rs("name")
					
					'Get OTS Counts
					set rs2 = server.CreateObject("ADODB.recordset")
					rs2.Open "spListOTSFixedAndFound4Root " & rs("RootID") ,cn,adOpenForwardOnly	
					strVersionOTSIDList = ""
					strVersionOTSFixedList = ""
					strVersionOTSFoundList = ""
					do while not rs2.eof
						strVersionOTSIDList = strVersionOTSIDList & "," & trim(rs2("ID"))
						strVersionOTSFixedList = strVersionOTSFixedList & "," & trim(rs2("OTSFixed")&"")
						strVersionOTSFoundList = strVersionOTSFoundList & "," & trim(rs2("OTSFound")&"")
						rs2.movenext	
					loop
					rs2.close
					set rs2= nothing

                    
					VersionOTSIDArray = split(strVersionOTSIDList,",")
					VersionOTSFixedArray = split(strVersionOTSFixedList,",")
					VersionOTSFoundArray = split(strVersionOTSFoundList,",")
				end if

				strOTSFound = "0"
				strOTSFixed = "0"
				for i = lbound(VersionOTSIDArray) to ubound(VersionOTSIDArray)
					if trim(VersionOTSIDArray(i)) = trim(rs("ID")) then
						if trim(VersionOTSFoundArray(i))<>"" then
							strOTSFound = VersionOTSFoundArray(i)
							OTSFound = OTSFound + strOTSFound
						end if
						if trim(VersionOTSFixedArray(i))<>"" then
							strOTSFixed = VersionOTSFixedArray(i)
							OTSFixed = OTSFixed + strOTSFixed
						end if
						exit for
					end if
				next
				strVersion = rs("Version") & ""
				if trim(rs("Revision") & "") <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if trim(rs("Pass") & "") <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				ReleaseCount=ReleaseCount+1
				Response.Write "<TR>"
				Response.Write "<TD>" & rs("ID") & "</TD>"
				Response.Write "<TD>" & strVersion & "</TD>"
				if isnull(rs("actualreleasedate")) then
					Response.Write "<TD>" & rs("location") & "</TD>"
					if instr(rs("location") & "","Fail")>0 then
						FailCount=FailCount+1
					else
						InTestCount=InTestCount+1
					end if
				else
					Response.Write "<TD>" & formatdatetime(rs("ActualReleaseDate"),vbshortdate) & "</TD>"
				end if

				Response.Write "<TD align=middle>" & strOTSFixed & "</TD>"
				Response.Write "<TD align=middle>" & strOTSFound & "</TD>"
				Response.Write "<TD>" & shortname(rs("Developer")) & "</TD>"
				Response.Write "</TR>"
				rs.MoveNext
			loop
			rs.Close
			Response.Write "</table>"
			if rowcount < 150 then
				Response.Write "<Table><tr>"
				Response.Write "<td>Versions Released:</td><td>" & ReleaseCount & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.Write "<td>Versions In Test:</td><td>" & InTestCount & "</TD>"
				Response.Write "<td>OTS Fixed:</td><td>" & OTSFixed & "</td></tr>"
				Response.Write "<tr><td>Versions Failed:</td><td>" & FailCount & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.Write "<td>OTS Found:</td><td>" & OTSFound & "</TD>"
				Response.Write "</tr></table><BR><BR>"
			end if
			Response.Write "</font>"
		elseif request("txtFunction") <> "2" then
			RowCount=0
			dim strProductList
			dim strLastRoot
			dim strLastVersion

			dim OTSP1Investigating
			dim OTSP1InsufficientInfo
			dim OTSP1SingleUnit
			dim OTSP1FixInProgress
			dim OTSP1Retest
			dim OTSP1Other
			dim OTSP2Investigating
			dim OTSP2InsufficientInfo
			dim OTSP2SingleUnit
			dim OTSP2FixInProgress
			dim OTSP2Retest
			dim OTSP2Other

			dim OTSP3Investigating
			dim OTSP3InsufficientInfo
			dim OTSP3SingleUnit
			dim OTSP3FixInProgress
			dim OTSP3Retest
			dim OTSP3Other

			OTSP1Investigating = 0
			OTSP1InsufficientInfo = 0
			OTSP1SingleUnit = 0
			OTSP1FixInProgress = 0
			OTSP1Retest = 0
			OTSP1Other = 0

			OTSP2Investigating = 0
			OTSP2InsufficientInfo = 0
			OTSP2SingleUnit = 0
			OTSP2FixInProgress = 0
			OTSP2Retest = 0
			OTSP2Other = 0

			OTSP3Investigating = 0
			OTSP3InsufficientInfo = 0
			OTSP3SingleUnit = 0
			OTSP3FixInProgress = 0
			OTSP3Retest = 0
			OTSP3Other = 0

			strOTSIDList = ""
			'Response.Write "<BR><BR><font size=2 color=red face=verdana>This report is still under development.</font><BR><BR>"
			Response.Write "<DIV ID=OTSSummaryTableTop><font size=2 face=verdana>Accessing OTS.  Please wait...</font></DIV>"
			
			'Response.Flush
			
			rs.Open "SELECT c.name as category, r.ID as RootID, pv.typeid, pv.productstatusid, pv.id as ProductID, pv.DOTSName as Product, pd.id as ProductDeliverableID, v.ID, r.Name, v.version, v.revision, v.pass, v.vendorversion, v.Location, v.ImagePath, e2.name as DevManager, e.name as Developer, pd.targeted , pd.inimage, v.languages, r.softpaq, v.certificationstatus, r.certrequired " & strSQL,cn,adOpenForwardOnly
			if not (rs.EOF and rs.BOF) then
				Response.Write "<TABLE style=""BORDER-LEFT-STYLE: none;"" border=1 cellpadding=2 cellspacing = 0 width=""100%"" LANGUAGE=javascript onmouseover=""return ROW_onmouseover()"" onmouseout=""return ROW_onmouseout()"">"
			end if
			strlastRoot = ""
			strLastVersion = ""
			do while not rs.EOF
				rowcount = rowcount +1
				if rs("ID") <> strLastVersion and strLastVersion <> "" then
					if strProductList = "" then
						strProductList = "&nbsp;"
					else
						strProductList  = mid(strProductList,2)
					end if
					Response.Write strVersionRow & "<TD colspan=2>" & strproductList &  "</TD></TR>"
					strProductList = ""
					strVersionRow = ""
				end if

				if strLastRoot <> rs("RootID") then
					if strLastRoot <> "" then
						'Response.Write "</table></td></tr>"
						Response.Write "</tr>"
						Response.Write "<TR><td class=SectionHeader colspan=8><b>Open Observations:</b></td></tr>"
						Response.Write GetObservations(strLastRoot)
						Response.Write "<TR style=""HEIGHT:50;BORDER-LEFT-STYLE: none;BORDER-TOP-STYLE: none;"" bgcolor=white><TD  style=""BORDER-LEFT-STYLE: none;BORDER-RIGHT-STYLE: none"" colspan=8>&nbsp;</TD></TR>"
					end if
					Response.write "<TR><TD class=RootHeader colspan=6><b>Name:&nbsp;</b>" & rs("Name") & "</td><TD class=RootHeader nowrap><b>Manager:</b>&nbsp;" & rs("DevManager") & "</td><TD class=RootHeader nowrap><b>Category:</b>&nbsp;" & rs("Category") & "</td></tr>"
					Response.Write "<TR><td class=SectionHeader colspan=8><b>Deliverable&nbsp;Versions:</b></td></tr>"
					Response.Write "<TR bgcolor=beige><td colspan=3><b>Version</b></td><td><b>Developer</b></td><td colspan=2><b>Workflow</b></td><td colspan=2><b>Where Used (Targeted or In Image)</b> - Gray products are no longer in development</td></tr>"
					'Response.write "<TR bgcolor=ivory><TD><Table width=""100%"">"
					strLastRoot = rs("RootID")
				end if
				if rs("ID") <> strLastVersion then 'Capture the deliverable info once
					strVersion = rs("Version") & "," & rs("Revision") & "," & rs("Pass")
					strVersionRow =  "<TR bgcolor=ivory class=""Row"" LANGUAGE=javascript onclick=""return DelROW_onclick('" & clng(rs("id")) & "','" & clng(rs("Rootid")) & "')""><TD valign=top colspan=3>" & strversion & "</TD><TD valign=top nowrap>" & shortname(rs("Developer")) & "</TD><TD colspan=2 valign=top nowrap>" & rs("Location") & "</TD>"			
				end if

				if OTSRowCount >=500 then
					Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Status Reports are limited by total observation count.</b><BR><BR></font>"
					exit do
				end if

				if rs("Inimage") or rs("targeted") then
					if rs("TypeID") = 1 and rs("ProductStatusID") > 2 then
						strProductList = strProductList & ",<font color=darkgray>" &  rs("Product") & ""
					elseif rs("TypeID") = 1 then
						strProductList = "," &  rs("Product") & strProductList 
					end if
				end if
				strLastVersion = rs("ID")
				rs.MoveNext
			loop
			if rs.EOF and rs.BOF then
				rs.Close
			else
				rs.Close

				if strProductList = "" then
					strProductList = "&nbsp;"
				else
					strProductList  = mid(strProductList,2)
				end if
				Response.Write strVersionRow & "<TD colspan=2>" & strproductList &  "</TD></TR>"
				Response.Write "<TR bgcolor=gainsboro><td colspan=8 class=SectionHeader><b>Open Observations:</b></td></tr>"
				Response.Write GetObservations(strLastRoot)
				Response.Write "</TABLE>"
			end if

		end if	


		if strOTSIDList <> "" then
			strOTSIDList = mid(strOTSIDList,2)
			strSQL = "SELECT Priority, State, Count(*) as OTSCount " & _
					 "FROM HOUSIREPORT01.SIO.dbo.SI_Observation_Report_Simplified o (NOLOCK) " & _
					 "WHERE observationid in (" & strOTSIDLIst & ") "
			if request("txtFunction") = "4" then
				strSQl = strSQL & " and Priority in (0,1) "
			end if
			strSQL = strSQL & " group by Priority, State " & _
					 "order by priority;"
			rs.Open strSQL,cn,adOpenForwardOnly
			do while not rs.EOF
				Select Case lcase(rs("State"))
				case "insufficient information"
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1InsufficientInfo = OTSP1InsufficientInfo + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2InsufficientInfo = OTSP2InsufficientInfo + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3InsufficientInfo = OTSP3InsufficientInfo + rs("OTSCount")
					end if
				case "under investigation"
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1Investigating = OTSP1Investigating + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2Investigating = OTSP2Investigating + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3Investigating = OTSP3Investigating + rs("OTSCount")
					end if
				case "single unit failure"
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1SingleUnit = OTSP1SingleUnit + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2SingleUnit = OTSP2SingleUnit + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3SingleUnit = OTSP3SingleUnit + rs("OTSCount")
					end if
				case "fix in progress"
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1FixInProgress = OTSP1FixInProgress + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2FixInProgress = OTSP2FixInProgress + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3FixInProgress = OTSP3FixInProgress + rs("OTSCount")
					end if
				case "retest"
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1Retest = OTSP1Retest + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2Retest = OTSP2Retest + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3Retest = OTSP3Retest + rs("OTSCount")
					end if
				case else
					if rs("Priority") = 0 or rs("Priority") = 1 then
						OTSP1Other = OTSP1Other + rs("OTSCount")
					elseif rs("Priority") = 2 then
						OTSP2Other = OTSP2Other + rs("OTSCount")
					elseif rs("Priority") > 2 then
						OTSP3Other = OTSP3Other + rs("OTSCount")
					end if
				end select 
				
				rs.MoveNext
			loop
			rs.Close
			'Output OTS Summary Table

			dim OTSP1Total

			dim OTSP2Total
			dim OTSP3Total
			dim OTSTotal

			dim OTSInvestigatingTotal
			dim OTSInsufficientInfoTotal
			dim OTSSingleUnitTotal
			dim OTSFixInProgressTotal
			dim OTSRetestTotal
			dim OTSOtherTotal

			OTSP1Total = OTSP1Investigating + OTSP1Retest + OTSP1Other + OTSP1InsufficientInfo + OTSP1SingleUnit + OTSP1FixInProgress
			OTSP2Total = OTSP2Investigating + OTSP2Retest + OTSP2Other + OTSP2InsufficientInfo + OTSP2SingleUnit + OTSP2FixInProgress
			OTSP3Total = OTSP3Investigating + OTSP3Retest + OTSP3Other + OTSP3InsufficientInfo + OTSP3SingleUnit + OTSP3FixInProgress
			OTSTotal = OTSp1Total + OTSP2Total + OTSP3Total

			OTSInvestigatingTotal = OTSP1Investigating + OTSP2Investigating + OTSP3Investigating
			OTSInsufficientInfoTotal = OTSP1InsufficientInfo + OTSP2InsufficientInfo + OTSP3InsufficientInfo
			OTSSingleUnitTotal = OTSP1SingleUnit + OTSP2SingleUnit + OTSP3SingleUnit
			OTSFixInProgressTotal = OTSP1FixInProgress + OTSP2FixInProgress + OTSP3FixInProgress
			OTSRetestTotal = OTSP1Retest + OTSP2Retest + OTSP3Retest
			OTSOtherTotal = OTSP1Other + OTSP2Other + OTSP3Other

			Response.Write "<DIV ID=OTSSummaryTable><font size=3 face=verdana>Observation Summary</font>"
			Response.Write "<TABLE bgcolor=ivory border=1 cellpadding=2 cellspacing = 0>"
			Response.Write "<TR bgcolor=beige><TD valign=bottom><b>Priority</b></TD><TD align=middle width=80><b>Insufficient<BR>Information</b></TD><TD width=80 align=middle width=80><b>Under<BR>Investigation</b></TD><TD align=middle width=80><b>Single Unit<BR>Failure</b></TD><TD align=middle width=80><b>Fix In<BR>Progress</b></TD><TD valign=bottom width=80 align=middle><b>Retest</b></TD><TD valign=bottom width=80 align=middle><b>Other</b></TD><TD valign=bottom width=80 align=middle><b>Total</b></TD></TR>"
			Response.Write "<TR><TD>P1/P0</TD><TD align=middle>" & OTSP1InsufficientInfo & "</TD><TD align=middle>" & OTSP1Investigating & "</TD><TD align=middle>" & OTSP1SingleUnit & "</TD><TD align=middle>" & OTSP1FixInProgress & "</TD><TD align=middle>" & OTSP1Retest & "</TD><TD align=middle>" & OTSP1Other & "</TD><TD align=middle>" & OTSP1Total & "</TD></TR>"
			if request("txtFunction") = "3" then
				Response.Write "<TR><TD>P2</TD><TD align=middle>" & OTSP2InsufficientInfo & "</TD><TD align=middle>" & OTSP2Investigating & "</TD><TD align=middle>" & OTSP2SingleUnit & "</TD><TD align=middle>" & OTSP2FixInProgress & "</TD><TD align=middle>" & OTSP2Retest & "</TD><TD align=middle>" & OTSP2Other & "</TD><TD align=middle>" & OTSP2Total & "</TD></TR>"
				Response.Write "<TR><TD>P3+</TD><TD align=middle>" & OTSP3InsufficientInfo & "</TD><TD align=middle>" & OTSP3Investigating & "</TD><TD align=middle>" & OTSP3SingleUnit & "</TD><TD align=middle>" & OTSP3FixInProgress & "</TD><TD align=middle>" & OTSP3Retest & "</TD><TD align=middle>" & OTSP3Other & "</TD><TD align=middle>" & OTSP3Total & "</TD></TR>"
			end if
			Response.Write "<TR bgcolor=beige><TD>Total</TD><TD align=middle>" & OTSInsufficientInfoTotal & "</TD><TD align=middle>" & OTSInvestigatingTotal & "</TD><TD align=middle>" & OTSSingleUnitTotal & "</TD><TD align=middle>" & OTSFixInProgressTotal & "</TD><TD align=middle>" & OTSRetestTotal & "</TD><TD align=middle>" & OTSOtherTotal & "</TD><TD align=middle>" & OTSTotal & "</TD></TR>"
			Response.Write "</TABLE><BR><BR><font size=3 face=verdana>Deliverables</font></DIV>"

		end if

		if request("txtFunction") = "2"  or (request("txtFunction") = "1" and request("chkChangeType") <> "") then 
			if request("txtFunction") = "1" then
				Response.Write "<BR><BR><font size=2 face=verdana><b>Deliverable History</b></font>"
			end if

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

			if request("cboHistoryRange") = "Range" then
				if request("txtStartDate") = "" then
					strStartDate = formatdatetime(now,vbshortdate)
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
			strFinalSQL = ""

			if instr(request("chkChangeType"),"21") > 0 then
				strFinalSQL = strFinalSQL & " Union Select p.dotsname, v.deliverablename, L.ToInfo, L.FromInfo, t1.Status as FromStatus, t2.Status as ToStatus, l.Updated as DateUpdated, a.headername as ChangeType, v.version, v.revision, v.pass, v.partnumber, v.modelnumber, v.id as VersionID, p.dotsname as product, vd.name as Vendor, r.name as Deliverable from Actions a with (NOLOCK), ActionLog l with (NOLOCK), ProductVersion p with (NOLOCK), vendor vd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), TestStatus t1 with (NOLOCK), TestStatus t2 with (NOLOCK), product_deliverable pd with (NOLOCK) where pd.productversionid = p.id and pd.deliverableversionid = v.id and l.Updated between '" & strStartDate & "' and '" & Dateadd("d",1,cdate(strEndDate)) & "' " & strSpecificQualStatus & " and t1.id = l.FromID and t2.id = l.ToID and r.id = v.deliverablerootid and a.actionid = l.actionid and v.id = l.deliverableversionid and vd.id = v.vendorid and l.productversionid = p.id and l.actionid in(21) "
				strFinalSQL = strFinalSQL & " and pd.id in (" & "SELECT pd.id as ProductDeliverableID " & strSQL & ") "
			end if

			if instr(request("chkChangeType"),"22") > 0 then
				strFinalSQL = strFinalSQL & " Union Select p.dotsname, v.deliverablename, L.ToInfo, L.FromInfo, t1.name as FromStatus, t2.Name as ToStatus, l.Updated as DateUpdated, a.headername as ChangeType, v.version, v.revision, v.pass, v.partnumber, v.modelnumber, v.id as VersionID, p.dotsname as product, vd.name as Vendor, r.name as Deliverable from Actions a with (NOLOCK), ActionLog l with (NOLOCK), ProductVersion p with (NOLOCK), vendor vd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), PilotStatus t1 with (NOLOCK), PilotStatus t2 with (NOLOCK), product_deliverable pd with (NOLOCK) where pd.productversionid = p.id and pd.deliverableversionid = v.id and  l.Updated between '" & strStartDate & "' and '" & Dateadd("d",1,cdate(strEndDate)) & "' " & strSpecificPilotStatus & " and t1.id = l.FromID and t2.id = l.ToID and r.id = v.deliverablerootid and a.actionid = l.actionid and v.id = l.deliverableversionid and vd.id = v.vendorid and l.productversionid = p.id and l.actionid in (22) "
				strFinalSQL = strFinalSQL & " and pd.id in (" & "SELECT pd.id as ProductDeliverableID " & strSQL & ") "
			end if

			if strFinalSQL <> "" then
				strFinalSQL = mid(strFinalSQL,8)
			end if				
			rs.open strFinalSQL & " order by p.dotsname, v.deliverablename;",cn,adOpenStatic
			if rs.EOF and rs.BOF then
				Response.Write "<BR><font size=2 face=verdana>No History Records match your criteria.</font><BR><BR>"
			else		
				Response.Write "<TABLE ID=tblResults border=1 cellspacing=0 cellpadding=1 width=""100%"" bgcolor=ivory>"
				Response.Write "<THEAD><TR bgcolor=beige>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 0,1,3);"">ID</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 1,0,2);"">Type</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 2,0,2);"">From</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 3,0,2);"">To</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 4,2,2);"">Updated</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 5,0,2);"">Product</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 6,0,2);"">Vendor</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 7,0,2);"">Deliverable</a></b></font></TD>"
				Response.Write "<TD><font size=1 face=verdana><b><a href=""javascript: SortTable( 'tblResults', 8,0,2);"">Model</a></b></font></TD>"
				Response.Write "</TR></THEAD>"
			end if
			do while not rs.EOF
				RowCount = RowCount + 1
				if RowCount >=30000 then
					Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because History Reports are limited to 30000 records.</b><BR><BR></font>"
					exit do
				end if
				strVersion = rs("Version")
				if rs("Revision") & "" <> "" then
					strVersion = strVersion & "," &  rs("Revision")
				end if
				if rs("Pass") & "" <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				if trim(strVersion) <> "" then
					strVersion = " [" & strVersion & "]"
				end if

				strDate= trim(rs("DateUpdated") & "")
				if strDate <> "" then
					strDate = formatdatetime(strDate,vbshortdate)
				end if

				if not isnull(rs("FromInfo")) then
					strFrom = rs("FromInfo")
				else
					strFrom = rs("FromStatus") & ""
				end if

				if not isnull(rs("ToInfo")) then
					strTo = rs("ToInfo")
				else
					strTo = rs("ToStatus") & ""
				end if
				Response.Write "<TR>"
				Response.Write "<TD><font size=1 face=verdana><a href=""javascript: ShowDeliverableDetails(" & rs("VersionID") & ");"">" & rs("VersionID") & "</a></font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & rs("ChangeType") & "</font></TD>"
				Response.Write "<TD nowrap><font size=1 face=verdana>" & strFROM  & "</font></TD>"
				Response.Write "<TD nowrap><font size=1 face=verdana>" & strTo & "</font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & strDate & "</font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & rs("Product") & "</font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & rs("vendor") & "</font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & rs("Deliverable")  & "</font></TD>"
				Response.Write "<TD><font size=1 face=verdana>" & rs("ModelNumber") & "</font></TD>"
				Response.Write "</TR>"
				rs.MoveNext
			loop
			if not (rs.EOF and rs.BOF) then
				Response.Write "</TABLE>"
			end if
			Response.Write   "<BR><font size=1 face=verdana>History Records Displayed: " & RowCount & "</font>"
			rs.Close

		end if

		cn.Close
		set rs=nothing
		set cn=nothing
	end if

function GetObservations(ID)
	dim strOTS
	dim strOTSHeader
		if ID <> "" then
			set rs2 = server.CreateObject("ADODB.recordset")		
			if request("txtFunction") = "3" then
				rs2.Open "spListOTS4Root " & clng(ID),cn,adOpenForwardOnly
			else
				rs2.Open "spListOTS4Root " & clng(ID) & ",2",cn,adOpenForwardOnly
			end if
			strOTSHeader = "<TR bgcolor=beige>"
			strOTSHeader = strOTSHeader &  "<TD><b>ID</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD><b>Pr</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD><b>State</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD><b>Owner</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD nowrap><b>Version</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD nowrap><b>Product</b></TD>"
			strOTSHeader = strOTSHeader &  "<TD colspan=2><b>Summary</b></TD>"
			strOTSHeader = strOTSHeader &  "</TR>"
	
			if rs2.EOF and rs2.BOF then
				strOTS = "<TR bgcolor=ivory><TD colspan=8>none</TD></TR>"
			else
				strOTS = strOTSHeader 
				do while not rs2.EOF
					strOTS = strOTS & "<TR class=""Row"" LANGUAGE=javascript onclick=""return OTSROW_onclick('" & clng(rs2("Observationid")) & "')"" bgcolor=ivory ><TD>" & rs2("ObservationID") & "</TD><TD>" & rs2("Priority") & "</TD><TD>" & rs2("State") & "</TD><TD>" & shortname(rs2("OwnerName")& "") & "</TD><TD>" & rs2("OTSComponentVersion") & "</TD><TD nowrap>" & rs2("product") & "</TD><TD colspan=2>" & rs2("Summary") & "</TD></TR>"
					if len(strOTSIDList) <= 65000 then
						strOTSIDList = strOTSIDList & "," & rs2("ObservationID")
					end if
					OTSRowCount = OTSRowCount + 1
					rs2.MoveNext
				loop
			end if

			GetObservations = strOTS
			rs2.Close
			set rs2 = nothing
		else
			GetObservations = "<TR bgcolor=ivory><TD colspan=8>none</TD></TR>"
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
<font Size="2" Color="red"><p><strong>Confidential</strong></p></font>
</body>
</html>
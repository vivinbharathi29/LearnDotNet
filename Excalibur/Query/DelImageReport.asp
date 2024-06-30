<%@ Language=VBScript %>


	<%
	
	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
	end if

	%>

<HTML>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<Title>Deliverable Query Results</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<% 'Response.Flush %>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + RootID + "," + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Properties</FONT></SPAN></DIV>";

	popupBody = popupBody + "</DIV>";

    oPopup.document.body.innerHTML = popupBody; 
	
	oPopup.show(lefter, topper, 100, 50, document.body);
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
</SCRIPT>
</HEAD>
<STYLE>
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

<H3><font face=verdana><%= Server.HTMLEncode(request("txtTitle")) %></font></H3>
<!--<span ID=lblLoad><font size=2 face=verdana>Loading.  Please wait...</font></span>-->

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

	strSQL = "Select v.id, v.imagepath, pd.id as productdeliverableid, v.deliverablerootid as rootid, p.dotsname as product, v.deliverablename as Deliverable, v.version, v.revision, v.pass,a.Updated as AddedToImage " & _
			 "from ActionLog a with (NOLOCK), ProductVersion p with (NOLOCK), DeliverableVersion v with (NOLOCK), product_deliverable pd with (NOLOCK) " & _
			 "where a.actionid = 14 " & _
			 "and v.id = a.deliverableversionid " & _
			 "and pd.productversionid=p.id " & _
			 "and pd.deliverableversionid=v.id " & _
			 "and p.id = a.productversionid " & _
			 "and a.deliverableversionid in (Select ID from deliverableversion with (NOLOCK) where deliverablerootid=" & request("ID") & ") " & _
			 "order by p.dotsname,AddedToImage" 	

	dim RowCount
	
	RowCount=0
	
	if 	request("ID") = "" then
		Response.Write "<font size=2 face=verdana>No report criteria selected.  Please select the appropriate criteria and try again.</font>"
	else
		rs.Open strSQl,cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			Response.Write "<TABLE ID=tblResults width=""100%""cellspacing=0 cellpadding=2 bgcolor=ivory border=1>"
			Response.Write "<THead bgcolor=beige>"
			Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 0,1,1);"">ID</a></font></TD>"
			Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 1,0,1);"">Deliverable</a></font></TD>"
			Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 2,0,1);"">Version</a></font></TD>"
			Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 3,0,1);"">Product</a></font></TD>"
			Response.Write "<TD><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 4,0,1);"">Added To Image</a></font></TD>"
			Response.Write "</THead>"
		else
			Response.Write "<br><font size=2 face=verdana>No items match your query criteria</font>"
		end if
		RowCount = 0
	
		do while not rs.EOF
			RowCount = RowCount + 1
			strVersion = rs("Version") & ""
			if rs("revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("pass") & "" <> "" then
				strVersion = strVersion & "," & rs("pass")
			end if
				
			strDelName =  rs("Deliverable")
				
			Response.Write  "<TR>"
			Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick(" & rs("ProductDeliverableID") & "," & rs("RootID") & "," & rs("ID") & ")"" valign=top>" & rs("ID") & "</font></TD>"
			Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick(" & rs("ProductDeliverableID") & "," & rs("RootID") & "," & rs("ID") & ")"" valign=top nowrap>" & strDelName & "</font></TD>"
			Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick(" & rs("ProductDeliverableID") & "," & rs("RootID") & "," & rs("ID") & ")""  valign=top nowrap>" & strVersion & "</font></TD>"
			Response.Write "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick(" & rs("ProductDeliverableID") & "," & rs("RootID") & "," & rs("ID") & ")""  valign=top nowrap align=left>" & rs("Product") & "</font></TD>"
			Response.Write  "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick(" & rs("ProductDeliverableID") & "," & rs("RootID") & "," & rs("ID") & ")""  valign=top nowrap align=left>" & rs("AddedToImage") & "</font></TD>"
			Response.Write "<TD ID=PATH" & trim(rs("ProductDeliverableID")) & " style=""Display:none"">" & rs("ImagePath") & "</TD>"
			Response.Write  "</TR>"
			rs.MoveNext
		loop
		if (not (rs.EOF and rs.BOF)) then
			Response.Write   "</table>"
		end if
		Response.Write   "<BR><font size=1 face=verdana>Deliverables Displayed: " & RowCount & "</font><BR>"
		rs.Close

		end if	

		
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
	
%>

</BODY>
</HTML>

<%@ Language=VBScript %>
<!-- #include file = "./includes/Security.asp" --> 
<%
    if LCase(Request.QueryString("HideHeader")) <> "true" then
		Response.Redirect "/Pulsar/Component/Root/" & Request.QueryString("ID")
	End If
	
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<% 
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim DRID : DRID = regEx.Replace(Request.QueryString("ID"), "")
    Dim iShowBusiness : iShowBusiness = regEx.Replace(Request.QueryString("ShowBusiness"), "")
    Dim iShowProducts : iShowProducts = regEx.Replace(Request.QueryString("ShowProducts"), "")
    Dim iShowScr : iShowScr = regEx.Replace(Request.QueryString("ShowScr"), "")
    Dim iVersion : iVersion = regEx.Replace(Request.QueryString("Version"), "")
    
    regEx.Pattern = "[^0-9a-zA-Z_]"

    regEx.Pattern = "[^0-9a-zA-Z]"
    Dim hideHeader : hideHeader = regEx.replace(Request.QueryString("HideHeader"), "")
    Dim hideForPulsar : hideForPulsar = regEx.replace(Request.QueryString("HideHeader"), "")
    Dim sTab : sTab = regEx.Replace(Request.QueryString("Tab"), "")
    Dim sClass : sClass = regEx.Replace(Request.QueryString("Class"), "")
    Dim sStatus : sStatus = regEx.Replace(Request.QueryString("Status"), "")
    Dim sView : sView = regEx.Replace(Request.QueryString("View"), "")
    on error resume next
    Dim sDmStatus : sDmStatus = regEx.Replace(Request.Cookies("DMStatus"), "")
		on error goto 0

    Dim securityObj
    Set securityObj = New ExcaliburSecurity

    If sTab = "Certification" Then Server.Transfer("/Agency/dmview.asp")
    	
    if LCASE(hideForPulsar) = "true" Then hideForPulsar = "display:none;"
%>
<%

	dim strMyBrowser
    strBrowserWarning = ""
	'dim blnSkipPopup
	strMyBrowser = Request.ServerVariables("HTTP_User_Agent")
	if instr(strMyBrowser,"MSIE") > 0 then
		strMyBrowser = mid(strMyBrowser,instr(strMyBrowser,"MSIE")+5)
		if left(strMyBrowser,1) < 5 and left(strMyBrowser,2) <> "10" then
			strBrowserWarning =  "<font size=2 face=verdana><b>You browser may not be fully compatible with this application.  Please use IE 8 or better for best performance.<BR><BR> Contact <a href=""mailto:max.yu@hp.com"">Max Yu</a> if you think you should not be receiving this message.</B></font><BR>"
			blnSkipPopup = true
		elseif left(strMyBrowser,1) = 5 then
			if mid(strMyBrowser,3,1) < 5 then
			strBrowserWarning = "<font size=2 face=verdana><b>You browser may not be fully compatible with this application.  Please use IE 8 or better for best performance.<BR><BR> Contact <a href=""mailto:max.yu@hp.com"">Max Yu</a> if you think you should not be receiving this message.</B></font><BR>"
				blnSkipPopup = true
			end if
		else
			blnSkipPopup = false
		end if
	else
		blnSkipPopup = true			
		strBrowserWarning = "<font color=red size=2 face=verdana><b>Excalibur does not fully support your browser.  Please use IE 8 or later.</b><br></font>"
	end if

%>
<html>
<head>
<title>DM View</title>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<meta name="VI60_DefaultClientScript" content="JavaScript">
<!-- #include file="includes/bundleConfig.inc" -->
<!-- #include file="./Agency/AgencyPivot.asp" -->
<script src="includes/client/jquery.blockUI.js" type="text/javascript"></script>

<%if not blnSkipPopup then%>
	<script ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	
	var oPopup = window.createPopup();
	
	-->
	</script>
<%else%>
	<script ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	
	var oPopup;
	
	-->
	</script>


<%end if%>
<script language="javascript" src="_ScriptLibrary/jsrsClient.js"></script>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
<!-- #include file = "_ScriptLibrary/sort.js" -->

var SelectedRow;
var AddingID;

function UpdateUserAccess() {
	window.open("UpdateUserAccess.asp");
}
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

function contextMenu2(RootID, VersionID, TypeID, GroupID)
{
	if (window.event.srcElement.className != "text" && window.event.srcElement.className != "cell")
		return;

    // The variables "lefter" and "topper" store the X and Y coordinates
    // to use as parameter values for the following show method. In this
    // way, the popup displays near the location the user clicks. 
    var lefter = event.clientX;
    var topper = event.clientY;
    var popupBody;
    
		if (window.event.srcElement.className == "text")
			{
		    if (typeof(SelectedRow) != "undefined")
				if (SelectedRow != null)
					if (SelectedRow != window.event.srcElement.parentElement.parentElement)
						SelectedRow.style.color="black";
					
			SelectedRow = window.event.srcElement.parentElement.parentElement;
			SelectedRow.style.color="red";
			
			}
		else if (window.event.srcElement.className == "cell")
	    	{
		    if (typeof(SelectedRow) != "undefined")
				if (SelectedRow != null)
					if (SelectedRow != window.event.srcElement.parentElement)
						SelectedRow.style.color="black";
					
			SelectedRow = window.event.srcElement.parentElement;
			SelectedRow.style.color="red";
	    	}

		if(typeof(oPopup) == "undefined") 
			{
			DisplayVersion(RootID,VersionID);
			
			return;
			}
    
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:SendEmail(" + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Send&nbsp;Email&nbsp;...</SPAN></FONT></DIV>";
  
	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Component&nbsp;Detail&nbsp;Info&nbsp;&nbsp;</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Component&nbsp;History</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayEOLProperties(" + RootID + "," + VersionID + "," + GroupID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Update&nbsp;End&nbsp;of&nbsp;Availability&nbsp;Date...&nbsp;</FONT></SPAN></DIV>";

	popupBody = popupBody + "</DIV>";

    oPopup.document.body.innerHTML = popupBody; 
	
	oPopup.show(lefter, topper, 170, 66, document.body);

	//Adjust window size
	var NewHeight;
	var NewWidth;
	
	if (oPopup.document.body.scrollHeight> 66 || oPopup.document.body.scrollWidth> 170)
		{
		NewHeight = oPopup.document.body.scrollHeight;
		NewWidth = oPopup.document.body.scrollWidth;
		oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
		}


}


function contextMenu(RootID, VersionID, TypeID,isPM,isAccessoryPM, isCD)
{
	if (window.event.srcElement.className != "text" && window.event.srcElement.className != "cell")
		return;


    var lefter = event.clientX;
    var topper = event.clientY;
    var popupBody;
	var strFilename;
	var strPath = trim(document.all("Path" + VersionID).innerText);    
	var strISO = trim(document.all("ISO" + VersionID).innerText);    
    
		if (window.event.srcElement.className == "text")
			{
		    if (typeof(SelectedRow) != "undefined")
				if (SelectedRow != null)
					if (SelectedRow != window.event.srcElement.parentElement.parentElement)
						SelectedRow.style.color="black";
					
			SelectedRow = window.event.srcElement.parentElement.parentElement;
			SelectedRow.style.color="red";
			
			}
		else if (window.event.srcElement.className == "cell")
	    	{
		    if (typeof(SelectedRow) != "undefined")
				if (SelectedRow != null)
					if (SelectedRow != window.event.srcElement.parentElement)
						SelectedRow.style.color="black";
					
			SelectedRow = window.event.srcElement.parentElement;
			SelectedRow.style.color="red";
	    	}

		if(typeof(oPopup) == "undefined") 
			{
			DisplayVersion(RootID,VersionID);
			
			return;
			}
    
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	    if (TypeID == "1")
		    {
		    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		    popupBody = popupBody + "<FONT face=Arial size=2>";
		    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddVersion_onclick(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Clone&nbsp;Version...</SPAN></FONT></DIV>";
		    }
        else
            {
		    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		    popupBody = popupBody + "<FONT face=Arial size=2>";
		    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:CloneVersion(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Add&nbsp;New&nbsp;Version&nbsp...</SPAN></FONT></DIV>";
            }

    popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
   
    if (document.getElementById("ReleaseStep" + VersionID).innerText != "Complete")
    {
	    if (isPM==0 || isPM==1)
		    {
		    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		    popupBody = popupBody + "<SPAN style=\"white-space: nowrap\" onclick=\"parent.location.href='javascript:ReleaseVersion(" + RootID + "," + VersionID + ",1)'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Release&nbsp;From&nbsp;" + document.all("ReleaseStep" + VersionID).innerText + "...</FONT></SPAN></DIV>";
		
		    popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		    popupBody = popupBody + "<FONT face=Arial size=2>";
		    popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReleaseVersion(" + RootID + "," + VersionID + ",2)'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Fail&nbsp;Version...</FONT></SPAN></DIV>";

		    popupBody = popupBody + "<DIV>";
		    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
		    }
    }		
		
  popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
  popupBody = popupBody + "<FONT face=Arial size=2>";
  popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:SendEmail(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email&nbsp;...</SPAN></FONT></DIV>";
  
  if (strPath != "")
	{
	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:GetVersion(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Download</SPAN></FONT></DIV>";
	}


	if (TypeID != "1")
		{	
		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

		  if (strISO == "1")
			{
			popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
			popupBody = popupBody + "<FONT face=Arial size=2>";
			popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReleaseISO(" + VersionID + "," + txtUser.value + ")'\" >&nbsp;&nbsp;&nbsp;Request&nbsp;ISO&nbsp;Transfer...</SPAN></FONT></DIV>";
			}


		}
	else if (isPM==1)
		{	
		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:EditPartNumber(" + VersionID + "," + txtUser.value + ")'\" >&nbsp;&nbsp;&nbsp;Edit&nbsp;Part&nbsp;Number...</SPAN></FONT></DIV>";

		}
  if (strFilename == "!@#$#@!")
	{
	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
	}

  if (strFilename == "!@#$#@!")
	{
	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ReleaseDoc(" + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Release Doc</SPAN></FONT></DIV>";
	}
	
  if (strFilename == "!@#$#@!")
	{
	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:TestDoc(" + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Test Results</SPAN></FONT></DIV>";
	}

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
	if (txtPreinstallGroup.value=="1")
		{
		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:UpdateInternalRev(" + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Edit&nbsp;Preinstall&nbsp;Properties...</FONT></SPAN></DIV>";
		}
	if (isPM==0)
		{
		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:UpdateSchedule(" + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Update&nbsp;Schedule...</FONT></SPAN></DIV>";

		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
		}

	if (isAccessoryPM==1)
		{
		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:MultiUpdateTestStatus(0,0," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Batch&nbsp;Update&nbsp;Product&nbsp;Status&nbsp;...</SPAN></FONT></DIV>";
		
		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
		}
		
	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersionDetail(" + RootID + "," + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Component&nbsp;Detail&nbsp;Info&nbsp;&nbsp;</FONT></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDeliverableHistory(" + RootID + "," + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Component&nbsp;History</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + RootID + "," + VersionID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;&nbsp;Properties</FONT></SPAN></DIV>";

	popupBody = popupBody + "</DIV>";

    oPopup.document.body.innerHTML = popupBody; 
		oPopup.show(lefter, topper, 170, 66, document.body);

	//Adjust window size
	var NewHeight;
	var NewWidth;
	
	if (oPopup.document.body.scrollHeight> 66 || oPopup.document.body.scrollWidth> 170)
		{
		NewHeight = oPopup.document.body.scrollHeight;
		NewWidth = oPopup.document.body.scrollWidth;
		if(topper+NewHeight > document.body.clientHeight)
		    topper = document.body.clientHeight - NewHeight;
		if(topper < 0)
		    topper = 0;
		oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
		}


}

function MultiUpdateTestStatus(ProdID,RootID,VersionID){
	var NewTop;
	var NewLeft;
	
	NewLeft = (screen.width - 655)/2
	NewTop = (screen.height - 650)/2	
	var strResult;
	strResult = window.showModalDialog("Deliverable/Commodity/MultiTestStatusPulsar.asp?ProdID=" + ProdID + "&RootID=" + RootID + "&VersionList=" + VersionID,"","dialogWidth:900px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
}


function AddVersion_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function AddVersion_onmouseout() {
	window.event.srcElement.style.color = "blue";
}
function CloneVersion(VersionID){
	var NewTop;
	var NewLeft;
	
	NewLeft = (screen.width - 600)/2
	NewTop = (screen.height - 450)/2	
	
	var strResult;
    var strID;
	strResult = window.showModalDialog("SelectReleaseType.asp?ID=" + VersionID,"","dialogWidth:600px;dialogHeight:450px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strResult) != "undefined")
		{
        if (strResult=="1")
            {
            AddVersion_onclick(0);
            }
        else if (strResult=="2")
            {
    		strID = window.showModalDialog("WizardFrames.asp?RootID=" + txtID.value + "&CloneID=" + VersionID + "&CloneType=3","","dialogWidth:850px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No") 
        	if (typeof(strID) != "undefined")
		        {
		        window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Versions&ID=" + txtID.value);
		        window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + txtID.value);
		        }
            }
        else if (strResult=="4")
            {
    		strID = window.showModalDialog("WizardFrames.asp?RootID=" + txtID.value + "&CloneID=" + VersionID + "&CloneType=2","","dialogWidth:850px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No") 
        	if (typeof(strID) != "undefined")
		        {
		        window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Versions&ID=" + txtID.value);
		        window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + txtID.value);
		        }
            }
        }
}
function AddVersion_onclick(CopyID) {
	var strID;
	if (CopyID == 0)
		CopyID=""
	if (txtFilename.value == "HFCN")
		strID = window.showModalDialog("hfcn/hfcn.asp?RootID=" + txtID.value ,"","dialogWidth:700px;dialogHeight:650px;maximize:yes;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	else
		strID = window.showModalDialog("WizardFrames.asp?RootID=" + txtID.value + "&CopyID=" + CopyID,"","dialogWidth:850px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No") 

	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Versions&ID=" + txtID.value);
		window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + txtID.value);
		}
}

function AddRestrictions_onclick() {
	var strID;

	strID = window.showModalDialog("deliverable/restrict/restrict.asp?RootID=" + txtID.value,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No") 

	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Restriction&ID=" + txtID.value);
		window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + txtID.value);
		}
}


function window_onload() {
	LoadingMessage.style.display="none";
	var Found;
	var strFavorites;
	strFavorites = "," + txtFavs.value;
	Found = strFavorites.indexOf("," + txtID.value + ",");
	if (txtView.value!="1")
		{
		if (Found==-1)
			{
			RFLink.style.display="none";
			AFLink.style.display="";
			}
		else
			{
			RFLink.style.display="";
			AFLink.style.display="none";
			}
		}

	if (document.all['EOLLink'])
	{
		if (txtInactiveCount.value == "0" || txtInactiveCount.value == "")
			document.all['EOLLink'].style.display = "none";
		else
			document.all['EOLLink'].style.display = "";
	}

    //Instantiate modalDialog load
	modalDialog.load();

    //add datepicker
	load_datePicker();
}


function View_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function View_onmouseout() {
	window.event.srcElement.style.color = "blue";
}

function View_onclick() {
	window.showModalDialog("WizardFrames.asp?Type=1","","dialogWidth:850px;dialogHeight:650px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No") 
}

function DisplayVersionDetail(RootID, VerID){

	SelectedRow.style.color="black";
	SelectedRow=null;
	window.open("Query\\DeliverableVersionDetails.asp?Type=1&RootID=" + RootID + "&ID=" + VerID) 
}

function DisplayVersion(RootID, VerID){
	SelectedRow.style.color="black";
	SelectedRow=null;
	strID = window.showModalDialog("WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VerID,"","dialogWidth:850px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Versions&ID=" + RootID);
		window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + RootID);
		}
}

function DisplayEOLProperties(RootID, VerID, GroupID){
	SelectedRow.style.color="black";
	SelectedRow=null;
	strID = window.showModalDialog("Deliverable/EOLDate.asp?TypeID=" + GroupID + "&ID=" + VerID,"","dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.reload();
		}
}

function UpdateSchedule( VerID){
	SelectedRow.style.color="black";
	SelectedRow=null;

	strID = window.showModalDialog("deliverable/schedule.asp?ID=" + VerID,"","dialogWidth:500px;dialogHeight:350px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].location.reload();
		}
	//	}
}

function UpdateInternalRev( VerID){
	SelectedRow.style.color="black";
	SelectedRow=null;

   var strResult;		
    strResult = window.showModalDialog("Deliverable/PreinstallProperties.asp?ID=" + VerID,"","dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].location.reload();
        }
}

function ReleaseVersion(RootID, VerID,Action){
	if(typeof(SelectedRow) != "undefined") 
		{
		SelectedRow.style.color="black";
		SelectedRow=null;
		}
	strID = window.showModalDialog("Release.asp?ID=" + VerID + "&Action=" + Action,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].navigate ("dmview.asp?Tab=Versions&ID=" + RootID);
		window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + RootID);
		}
	
}

function GetVersion(VersionID){
	var strPath = trim(document.all("Path" + VersionID).innerText);    
	SelectedRow.style.color="black";
	SelectedRow=null;
	window.open ("FileBrowse.asp?ID=" + VersionID);	
}

function SendEmail(VersionID){
	var strResult;
	var NewTop;
	var NewLeft;
	
	NewLeft = (screen.width - 655)/2
	NewTop = (screen.height - 650)/2
	
	strResult = window.open("query/DelVerDetailSendEmail.asp?ID=" + VersionID); 
}

function ReleaseISO(VersionID,EmployeeID){
    window.open ("http://houcmitrel02.auth.hpicorp.net:81/iso_request.aspx?Type=v&EID=" + EmployeeID + "&ExcalID=" + VersionID);
	SelectedRow.style.color="black";
	SelectedRow=null;
}

function ReleaseDoc(RootID, VerID){
	window.open ("file:////ccafile10/PreRel$/ScriptPaq/SED98_F1.600/Release.Doc");
	SelectedRow.style.color="black";
	SelectedRow=null;
}

function TestDoc(RootID, VerID){
	window.open ("file:////ccafile10/PreRel$/ScriptPaq/SED98_F1.600/MAT.Doc");
	SelectedRow.style.color="black";
	SelectedRow=null;
}

function EditPartNumber(VersionID){
    var sURL = "deliverable/commodity/partnumber.asp?VersionID=" + VersionID;
    modalDialog.open({ dialogTitle: 'Edit Part Number', dialogURL: sURL, dialogHeight: 270, dialogWidth: 520, dialogResizable: true, dialogDraggable: true });
}

function CommodityResults(strID) {
    if (typeof(strID) != "undefined"){
        window.location.reload();
    }
}


function ChooseVersions(RootID, ProdID, ReleaseID){
	var strID;
	strID = window.showModalDialog("Deliverable/AdvancedSupport.asp?RootID=" + RootID + "&ProductID=" + ProdID + "&ProductDeliverableReleaseID=" + ReleaseID,"","dialogWidth:800px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	
	if (typeof(strID) != "undefined")
	{
		//window.location.reload();
		document.all("SupportCount" + ProdID).innerText = strID;
	}

}

function PreviewFiles(VerID){
	//window.alert("Download not implemented yet");
	//window.open ("Deliverable/generate.asp?VersionID=" + VerID);
	var strID;
	
	strID = window.showModalDialog("Deliverable/generate.asp?VersionID=" + VerID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	
	SelectedRow.style.color="black";
	SelectedRow=null;
}


function getCookieValue(cookieName)
{
	var cookieValue = document.cookie;
	var cookieStartsAt = cookieValue.indexOf(" " + cookieName + "=");
	if (cookieStartsAt == -1)
	{
		cookieStartsAt = cookieValue.indexOf(cookieName + "=");
	}
	if (cookieStartsAt == -1)
	{
		cookieValue="";
	}
	else
	{
		cookieStartsAt = cookieValue.indexOf("=",cookieStartsAt) + 1;
		var cookieEndsAt = cookieValue.indexOf(";",cookieStartsAt);
		if (cookieEndsAt == -1)
		{
			cookieEndsAt = cookieValue.length;
		}
		cookieValue=unescape(cookieValue.substring(cookieStartsAt,cookieEndsAt));
	}
	return cookieValue;
}

function AddFavorites(strID){
	var strFavorites;
	var FoundAt;
	var FavCount;
	
	AddingID = strID;
	
	strFavorites = txtFavs.value;
	FavCount = txtFavCount.value;
	if (FavCount == "" || FavCount == "NaN")
		FavCount=0;
	FavCount = Number(FavCount)+ 1;
	FoundAt = strFavorites.indexOf(strID + ",");		
	if (! FoundAt > -1)
	{
		strFavorites = strFavorites + strID + ","
		txtFavCount.value  = String(FavCount);
		txtFavs.value = strFavorites;
		//jsrsExecute("FavoritesRSupdate.asp", myCallback, "UpdateFavs",Array(strFavorites,String(FavCount),txtUser.value));
		ajaxurl = "FavoritesRSupdate.asp?CurrentUserID=" + txtUser.value + "&FavCount=" + String(FavCount) + "&Favorites=" + strFavorites;
		$.ajax({
		    url: ajaxurl,
		    type: "POST",
		    success: function (data) {
		        if (data == "1") {
		            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1&ID=" + AddingID);
		            RFLink.style.display="";
		            AFLink.style.display="none";
		        }
		    },
		    error: function (xhr, status, error) {
		        alert(error);
		    }

		});
	
	}
}


function myCallback( returnstring ){
		if (returnstring=="1")
			{
			window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1&ID=" + AddingID);
			RFLink.style.display="";
			AFLink.style.display="none";
			}
    } 
    
function ShowProperties(DisplayedID,strTab) {
	var strID;
	if (txtFilename.value == "HFCN")
		strID = window.showModalDialog("HFCN/HFCNAdd.asp?ID=" + DisplayedID,"","dialogWidth:600px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	else
		strID = window.showModalDialog("root.asp?ID=" + DisplayedID,"","dialogWidth:800px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (txtView.value!="1")
		{
			if (typeof(strID) != "undefined" )
			{
				window.parent.frames["RightWindow"].navigate ("dmview.asp?Prog=1&ID=" + DisplayedID);
				window.parent.frames["LeftWindow"].navigate ("tree.asp?Prog=1&ID=" + DisplayedID);
			}
		}
	else
		{
		window.location.reload();
		}	

}

function ShowScorecard(DisplayedID) {
	var strID;

	strID = window.showModalDialog("Deliverable/Scorecard/RootScorecard.asp?ID=" + DisplayedID,"","dialogWidth:700px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	//window.location.reload();
}

function ShowScorecardReport(DisplayedID) {
	var strID;

	window.open( "Deliverable/OTSCoreTeamDashboard.asp?RootID=" + DisplayedID,"_blank") 
}

function ShowScorecardCoreTeam(CoreTeamID,ReportID) {
	var strID;

	window.open( "Deliverable/OTSCoreTeamDashboard.asp?CoreTeamID=" + CoreTeamID + "&Report=" + ReportID,"_blank") 
}


function RemoveFavorites(strID){
	var strFavorites;
	var FoundAt;
	var FavCount;
	
	AddingID = strID;
	
	strFavorites = txtFavs.value;
	FavCount = txtFavCount.value;
	if (FavCount == "" || FavCount == "NaN")
		FavCount=0;
	FavCount = Number(FavCount)- 1;
	if (FavCount < 0)
		FavCount=0;
		
	FoundAt = strFavorites.indexOf(strID + ",");		
	if (! FoundAt > -1)
	{
		strFavorites = strFavorites.replace(strID+",","")
		txtFavCount.value  = String(FavCount);
		txtFavs.value = strFavorites;
	    //jsrsExecute("FavoritesRSupdate.asp", myCallback2, "UpdateFavs",Array(strFavorites,String(FavCount),txtUser.value));
		ajaxurl = "FavoritesRSupdate.asp?CurrentUserID=" + txtUser.value + "&FavCount=" + String(FavCount) + "&Favorites=" + strFavorites;
		$.ajax({
		    url: ajaxurl,
		    type: "POST",
		    success: function (data) {
		        if (data == "1") {
		            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1");
		            RFLink.style.display="none";
		            AFLink.style.display="";
		        }
		    },
		    error: function (xhr, status, error) {
		        alert(error);
		    }

		});
	}
}


function myCallback2( returnstring ){
		if (returnstring=="1")
			{
			window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1");
			RFLink.style.display="none";
			AFLink.style.display="";
			}
    } 


function versionrow_onmouseover() {
	var PopupOpen;
	if(typeof(oPopup) == "undefined") 
		PopupOpen = false;
		//return;
	else
		PopupOpen = oPopup.isOpen;
	
		
	if (! PopupOpen)
		{
		if (window.event.srcElement.className == "cell")
			{
    		window.event.srcElement.parentElement.style.color = "red";
			window.event.srcElement.parentElement.style.cursor = "hand";
			}
		else if (window.event.srcElement.className == "text")
			{
    		window.event.srcElement.parentElement.parentElement.style.color = "red";
			window.event.srcElement.parentElement.parentElement.style.cursor = "hand";		
			}

		if (typeof(SelectedRow) != "undefined")
			if (SelectedRow != null)
				SelectedRow.style.color="black";

		}
}

function versionrow_onmouseout() {
	var PopupOpen;
	if(typeof(oPopup) == "undefined") 
		PopupOpen = false;
		//return;
	else
		PopupOpen = oPopup.isOpen;

	if (! PopupOpen)
		{
		if (window.event.srcElement.className == "text")
	    	window.event.srcElement.parentElement.parentElement.style.color = "black";
		else if (window.event.srcElement.className == "cell")
	    	window.event.srcElement.parentElement.style.color = "black";
		}
}




function document_onmouseover() {
	var PopupOpen;
	if(typeof(oPopup) == "undefined") 
		PopupOpen = false;
		//return;
	else
		PopupOpen = oPopup.isOpen;

	if (! PopupOpen)
		{
		if (typeof(SelectedRow) != "undefined")
			{
			if (SelectedRow != null)
				SelectedRow.style.color="black";
			}
		}

}

function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}

function SetDisplayStatus(value) {

	var expireDate = new Date();

	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "DMStatus=" + value + ";expires=" + expireDate.toGMTString() + ";";

	window.location.reload(true);

}

function SetDMView(PageList, DelRootID, strClass) {

	var expireDate = new Date();
	window.location = "dmview.asp?Tab=" + PageList + "&ID=" + DelRootID + "&Class=" + strClass;
}

function openReport(ReportName, DelRootID, DelVerID)
{

	var nWindowHeight = screen.height
	var nWindowWidth = screen.width
	var sUrl = 'http://houhpqexcal01.auth.hpicorp.net/SQLReportServer?/ExcaliburReports/' + ReportName + '&rs:Command=Render&rc:Parameters=false&DeliverableRootID=' + DelRootID
	window.showModelessDialog(sUrl,'','dialogLeft:0;dialogRight:0;dialogHeight:'+nWindowHeight+';dialogWidth:'+nWindowWidth+';help:no;resizable:yes;status:no;');
}

function Export(intType){
	if (intType==1)
		ExportForm.txtData.value = "<TABLE BORDER=1>" + OTSTable.innerHTML + "</TABLE>";
	ExportForm.submit();
}

function OTSDetails(strID){
	var strID;
	var strDisplay;
	
	if (window.event.srcElement.className == "text")
    	{
    	strDisplay = window.event.srcElement.parentElement.parentElement.className;
    	window.event.srcElement.parentElement.parentElement.style.color = "black";
		ShowOTSDetails(strID);
		}
	else if (window.event.srcElement.className == "cell")
    	{
    	strDisplay = window.event.srcElement.parentElement.className;
    	window.event.srcElement.parentElement.style.color = "black";
		ShowOTSDetails(strID);
		}
}

function ShowOTSDetails(strID){
	var i;
	var strIDList="";
	var NewTop;
	var NewLeft;
	
	var strSort="";;
	
	if(sortedOn==-1 || sortedOn==1)
		strSort="&Sort1Column=o.observationid";
	else if (sortedOn==2)
		strSort="&Sort1Column=Priority";
	else if (sortedOn==3)
		strSort="&Sort1Column=State";
	else if (sortedOn==4)
		strSort="&Sort1Column=owner";
	else if (sortedOn==5)
		strSort="&Sort1Column=pm";
	else if (sortedOn==6)
		strSort="&Sort1Column=shortdescription";
	//+ ' ' +sortDirection
	if (strSort!="")
		{
			if(sortDirection==0)
				strSort = strSort + "&Sort1Direction=asc";
			else
				strSort = strSort + "&Sort1Direction=desc";
		}
	
	
	NewLeft = (screen.width - 655)/2;
	NewTop = (screen.height - 650)/2;
	if (typeof(strID) != "undefined")
		{
		strIDList = strID;
		}
	else if (typeof (chkOTSID.length) == "undefined" )
		{
		if (chkOTSID.checked)
			strIDList = chkOTSID.value;
		}
	else
		{
		for (i=0;i<chkOTSID.length;i++)
			if (chkOTSID(i).checked)
				strIDList = strIDList + "," + chkOTSID(i).value;
		if (strIDList != "")
			strIDList = strIDList.substr(1);
		}
	if (strIDList != "")
		{
		strResult = window.open("search/ots/Report.asp?txtReportSections=1&txtObservationID=" + strIDList + strSort,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No,scrollbars=Yes") 
		}
	else
		{
			alert("You must select at least one observation first"); 
		}
}

function onclick_ResetOTS(){
	var i;

	if (typeof (chkOTSID.length) == "undefined" )
		{
			chkOTSID.checked = chkAllOTS.checked;
		}
	else
		{
		for (i=0;i<chkOTSID.length;i++)
			chkOTSID(i).checked = chkAllOTS.checked;
		}
}


function ShowOTSAdvanced(strID){
	var i;
	var strIDList="";
	var NewTop;
	var NewLeft;
	
	NewLeft = (screen.width - 655)/2
	NewTop = (screen.height - 650)/2	
	
	if (typeof (chkOTSID.length) == "undefined" )
		{
		if (chkOTSID.checked)
			strIDList = chkOTSID.value
		}
	else
		{
		for (i=0;i<chkOTSID.length;i++)
			if (chkOTSID(i).checked)
				strIDList = strIDList + "," + chkOTSID(i).value
		if (strIDList != "")
			strIDList = strIDList.substr(1)
		}

	if (strIDList != "")
		{
		strResult = window.open("search/ots/default.asp?lstComponent=" + strID + "&txtObservationID=" + strIDList,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,status=yes,scrollbars=Yes") 
		}
	else
		{
		strResult = window.open("search/ots/default.asp?lstComponent=" + strID,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes") 
		}
}

function CloneRoot(CopyID){
	var strID;

	strID = window.showModalDialog("root.asp?Type=1&CopyID=" + CopyID,"","dialogWidth:800px;dialogHeight:650px;maximize:yes;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	
	if (typeof(strID) != "undefined")
		{
		window.parent.frames["RightWindow"].navigate ("dmview.asp?ID=" + strID);
		window.parent.frames["LeftWindow"].navigate ("tree.asp?ID=" + strID);
		}
}

function DisplayDeliverableHistory (RootID,VersionID){
	window.open ("Deliverable/DeliverableRootChange.asp?ID=" + RootID + "&VersionID=" + VersionID + "&ActionID=");
}

function ChooseColumns(strType){
    var strID;
    var SettingID;
    if (strType=="1")
        SettingID="6";
    else
        SettingID="5";
    
	strID = window.showModalDialog("ChooseColumns.asp?lstColumns=" + txtColumnList.value + "&UserSettingsID=" + SettingID,"","dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.reload();
        }
}

//-->
</script> 
<script LANGUAGE="javascript" FOR="document" EVENT="onmouseover">
<!--
 document_onmouseover()
//-->
</script>
<link href="style/wizard%20style.css" type="text/css" rel="stylesheet">
<link href="style/Excalibur.css" type="text/css" rel="stylesheet">
<style>
    BODY
    {
	    FONT-SIZE: xx-small;
    }
    P
    {
	    FONT-SIZE: xx-small;
    }
    A:visited
    {
        COLOR: blue
    }
    A:link
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }

    .MenuBar TD.ButtonSelected
    {
        COLOR: black;
        BACKGROUND-COLOR: wheat
    }
</style>
</head>
<body LANGUAGE="javascript" onload="return window_onload()">
<div style="display:none"><a href="UpdateUserAccess.asp"></a></div>
<div id="HideForPulsar" style="<%= hideForPulsar%>">
<font face="verdana">
<h3>
	<%
	
	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
		
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function

	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
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
	
	
	dim strID
	dim strDescription
	dim strNotes
	dim strManager
	dim strCategroy
	dim strVendor
	dim strPart
	dim strType
	dim strGreenSpec
	dim strFilename
	dim strPathCell
	dim strLeadfree
	dim strDisplayedList
	dim InactiveCount
	dim blnPM
	dim blnAccessoryPM
	dim blnSysAdmin
	dim strTester
	dim strSpec
	dim strManagerName
	dim strBuildLevel
	dim strDeveloper
	
	blnSysAdmin = 0
	blnPM = 0
	blnAccessoryPM=0

	strDisplayedList = "Versions"
	
	if trim(lcase(sTab)) = "versions" or trim(lcase(sTab)) = "ots" or trim(lcase(sTab)) = "agency" or trim(lcase(sTab)) = "restriction" or trim(lcase(sTab)) = "documents" or trim(lcase(sTab)) = "products" or trim(lcase(sTab)) = "naming" then
		strDisplayedList = sTab
	else
		strDisplayedList = "Versions"
	end if
	


	InactiveCount = "0"
	

	strVendor = ""
	strFilename = ""
	strBuildLevel = ""
	strID = clng(DRID)
	
	if strID <> "" and isnumeric(strID) then
		dim rs
		dim cn
        dim rs2
        
		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")


		dim CurrentUser 
		dim CurrentUserName
		dim CurrentUserID
		dim CurrentUserPartner
		dim CurrentUserSite
		dim strFavs
		dim strFavCount
		dim strDomainSite
		dim strTitleColor
		dim blnPreinstallGroup
		dim blnProcurementEngineer
        dim blnServiceCommodityManager
		dim blnLockReleases
		dim DevPMManagerID
		dim DevPMManagerName
		dim strTeamID
        dim blnShowOnStatus
        dim strProcurementGroup
        dim blnAgencyDataMaintainer

        strProcurementGroup = "0"

        DevPMManagerID = 0
        DevPMManagerName = ""

		blnProcurementEngineer=false
        blnServiceCommodityManager = false
		blnPreinstallGroup=false
		
		regEx.Pattern = "[^0-9a-fA-F#]"
		on error resume next
		strTitleColor = regEx.Replace(Request.Cookies("TitleColor"), "")
		if strTitleColor = "" then
			strTitleColor = "#0000cd"
		end if
		on error goto 0

		'Get User
		dim CurrentDomain
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
	
		set cm=nothing
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserName = rs("Name") & ""
			CurrentUserID = rs("ID")
			CurrentUserPartner = rs("PartnerID") & ""
			strFavs = trim(rs("Favorites") & "")
			strFavCount = trim(rs("FavCount") & "")
			if rs("workgroupid")= 15 or rs("workgroupid")= 22 then
				blnPreinstallGroup=true
			end if
			blnProcurementEngineer = rs("ProcurementEngineer") & ""
            blnServiceCommodityManager = rs("ServiceCommodityManager") 
			blnAccessoryPM = 0
			blnPM = 0
			blnSysAdmin = 0
			if rs("systemadmin") then
				blnSysAdmin = 1
				blnPM = 1
				blnAccessoryPM = 1
			end if
			if rs("CommodityPM") then
				blnPM = 1
			end if
			if rs("AccessoryPM") then
				blnAccessoryPM = 1
			end if
			if lcase(trim(rs("domain"))) = "asiapacific" then
				strDomainSite = 2
			else
				strDomainSite = 1
			end if
            if rs("AgencyDataMaintainer") > 0 then
    		    blnAgencyDataMaintainer = true
    		end if			
		end if
		rs.Close
        
       if blnServiceCommodityManager then
            strProcurementGroup = "2"
       elseif blnProcurementEngineer then
            strProcurementGroup = "1"
       else
            strProcurementGroup = "0"
       end if

		if currentdomain = "americas" then
			CurrentUserSite = "1"
		else
			CurrentUserSite = "2"
		end if

		strSQL = "spGetDelPropSummary " & clng(strID)
		rs.Open strSQL,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "Unable to find the selected component."
			Response.Write "<BR><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & Server.HTMLEncode(DRID) & ");""><font face=verdana size=1>Remove From Favorites</font></a>"
			Response.Write "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
			Response.Write "<font face=verdana size=2 id=LoadingMessage>Loading.  Please wait...</font>"
		else
			strDescription = server.HTMLEncode(rs("Description") & "") & "&nbsp;"
			strNotes = rs("Notes") & "&nbsp;"
			strManager = rs("Manager") & "&nbsp;"
			strDeveloper = rs("Developer") & "&nbsp;"
			strCategory = rs("Category") & "&nbsp;"
			blnLockReleases = rs("LockReleases")
			strVendor = rs("Vendor") & "&nbsp;"
			strPart = rs("BasePartNumber") & "&nbsp;"
			strTeamID = rs("TeamID") & ""
			strDevManagerID = rs("DevManagerID") 
			strTester = rs("Tester") & ""
			strType = rs("TypeID") & ""
			strSpec = trim(rs("DeliverableSpec") & "")
			strFilename = rs("Filename") & ""
			Response.Write rs("Name") & " Information"
			strDeliverableName = rs("Name") & ""
            strCoreTeamID = trim(rs("CoreTeamID") & "")
            blnShowOnStatus = rs("ShowOnStatus")
			if rs("Active")=0 then
				Response.write "<BR><BR><font size=1 face=verdana color=red>This component is inactive.</font><BR>"
			end if
		rs.Close
		if left(strSpec,2) = "\\" then
			strSpec = "<a href=""file://" & strSpec & """>" & strSpec & "</a>"
		end if


        'Lookup Manager ID and name
        strManagerName = ""
        if isnumeric(strDevManagerID) then
            rs.open "spGetManagerInfo " & clng(strDevManagerID),cn,adOpenForwardOnly
            if not (rs.eof and rs.bof) then
                DevPMManagerID  = rs("ID")
                DevPMManagerName = rs("Name")
                strManagerName = " or " & longname(DevPMManagerName)
            end if
            rs.close
        end if

	%>

 <br>
</h3>

	<%if strDisplayedList = "Products" then%>
		<font face="verdana" size="2" id="LoadingMessage">Counting Versions.  Please wait...</font>
	<%else%>
		<font face="verdana" size="2" id="LoadingMessage">Loading.  Please wait...</font>
	<%end if%>

<%	'Response.Flush%>


	<font size="1"><a href="javascript:ShowProperties(<%=DRID%>,'<%=Server.HTMLEncode(sTab)%>')">Edit Root Component</a></font>
	<span style="Display:none" id="RFLink"><font face="verdana" size="1" color="black">| </font><a href="javascript:RemoveFavorites(<%=DRID%>)"><font face="verdana" size="1">Remove From Favorites</font></a></span>
	<span style="Display:none" id="AFLink"><font face="verdana" size="1" color="black">| </font><a href="javascript:AddFavorites(<%=DRID%>)"><font face="verdana" size="1">Add To Favorites</font></a></span>
	<font face="verdana" size="1" color="black">| </font><a href="javascript:CloneRoot(<%=DRID%>)"><font face="verdana" size="1">Clone This Root</font></a> 
  <br> <br>
  <input type="hidden" id="txtTypeID" name="txtTypeID" value="<%=strType%>">
  <input type="hidden" id="txtFilename" name="txtFilename" value="<%=strFilename%>">
<table cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  
  <tr>
    <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Manager/OTS&nbsp;PM:</font></strong></td>
    <td><font size="1"><%=strManager%></font></td>
    <td bgColor="cornsilk"><strong><font size="1">Vendor:</font></strong></td>
    <td><font size="1"><%=strVendor%></font></td>
   </tr>
<%if strFilename = "HFCN" then%>
  <tr style="Display:none">
<%else%>
  <tr>
<%end if%>
    <%if trim(strType) = "1" then%>
        <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Execution&nbsp;Engineer:</font></strong></td>
    <%else%>
        <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Tester:</font></strong></td>
    <%end if%>
    <td><font size="1"><%=strTester%>&nbsp;</font></td>
    <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Category:</font></strong></td>
    <td><font size="1"><%=strcategory%></font></td></tr>

    <TR>
     <%if trim(strType) = "1" then%>
        <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Development&nbsp;Engineer:</font></strong></td>
     <%else%>
        <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Developer:</font></strong></td>
     <%end if%>
    <%if trim(strType) = "1" then%>
        <td colspan=3><font size="1"><%=strDeveloper%>&nbsp;</font></td>
    <%else%>
        <td><font size="1"><%=strDeveloper%>&nbsp;</font></td>
        <td nowrap width="100" bgColor="cornsilk"><strong><font size="1">Root&nbsp;Filename:</font></strong></td>
        <td><font size="1"><%=strFilename%></font></td>
    <%end if%>
    </tr>

	<%if trim(strSpec) <> "" then%>
	<tr>
    <td bgColor="cornsilk"><strong><font size="1">Functional Spec:</font></strong></td>
    <td colspan="4"><font size="1"><%=strSpec%></font></td>
	</tr>
	<%end if%>

  <tr>
    <td bgColor="cornsilk" valign="top"><strong><font size="1">Description:</font></strong></td>
    <td colspan="4"><font size="1"><%=replace(strDescription,vbcrlf,"<BR>")%></font></td></tr>
<%if strFilename = "HFCN" then%>
  <tr style="Display:none">
<%else%>
  <tr>
<%end if%>
    <td bgColor="cornsilk"><strong><font size="1">Notes:</font></strong></td>
    <td colspan="4"><font size="1"><%=strNotes%></font></td></tr>

    <tr>
    <td bgColor="cornsilk" valign="top"><strong><font size="1">Scorecard:</font></strong></td>
    <td colspan="4"><font face="verdana" size="1" color="black">
        </font><a href="javascript:ShowScorecard(<%=DRID%>)"><font face="verdana" size="1">Edit&nbsp;Scorecard</font></a>
        | <a href="javascript:ShowScorecardReport(<%=DRID%>)"><font face="verdana" size="1">Component&nbsp;Report</font></a>
        <%if trim(strCoreTeamID) <> "" and trim(strCoreTeamID) <> "0" and blnShowOnStatus then%>
        | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,0)"><font face="verdana" size="1">Core&nbsp;Team&nbsp;Report</font></a>
        | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,1)"><font face="verdana" size="1">Executive&nbsp;Summary</font></a>
        | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,2)"><font face="verdana" size="1">Action&nbsp;Items</font></a>
        <%end if%>
        </td>
    </tr>


</table><br>
<table border="1" bordercolor="Ivory" cellspacing="0" cellpadding="2" Id="menubar" Class="MenuBar"><tr bgcolor="<%=strTitleColor%>">
<%if strDisplayedList = "Versions" or strDisplayedList = ""  then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Version List&nbsp;&nbsp;</font></td>
<%else%>
	<td><font size="1" face="verdana"><a href="javascript:SetDMView('Versions', '<%= strID%>', '<%= sClass%>');">&nbsp;&nbsp;Version List</a>&nbsp;&nbsp;</font></td>
<%end if%>
<%if strDisplayedList = "OTS" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Observations&nbsp;&nbsp;</font></td>
<%else%>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('OTS', '<%= strID%>', '<%= sClass%>');">Observations</a>&nbsp;&nbsp;</font></td>
<%end if%>
<%if strDisplayedList = "Agency" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Agency&nbsp;&nbsp;</font></td>
<%elseif trim(strType) = "1" then %>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Agency', '<%= strID%>', '<%= sClass%>');">Agency</a>&nbsp;&nbsp;</font></td>
<%end if%>

<%if strDisplayedList = "Certification" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Certification&nbsp;<span style="color:Red; font-weight:bold;">(Beta)</span>&nbsp;&nbsp;</font></td>
<%elseif trim(strType) = "1" then %>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Certification', '<%= strID%>', '<%= sClass%>');">Certification&nbsp;<span style="color:Red; font-weight:bold;">(Beta)</span></a>&nbsp;&nbsp;</font></td>
<%end if%>

<%if strDisplayedList = "Restriction" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Restrictions&nbsp;&nbsp;</font></td>
<%elseif trim(strType) = "1" then %>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Restriction', '<%= strID%>', '<%= sClass%>');">Restrictions</a>&nbsp;&nbsp;</font></td>
<%end if%>

<td><font size="1" face="verdana">&nbsp;</font></td>

<%if strDisplayedList = "Documents" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Documents&nbsp;&nbsp;</font></td>
<%elseif trim(strTeamID) = "3" then%>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Documents', '<%= strID%>', '<%= sClass%>');">Documents</a>&nbsp;&nbsp;</font></td>
<%end if%>

<%if strDisplayedList = "Products" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Products&nbsp;&nbsp;</font></td>
<%else%>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Products', '<%= strID%>', '<%= sClass%>');">Products</a>&nbsp;&nbsp;</font></td>
<%end if%>

<%if strDisplayedList = "Naming" then%>
	<td class="ButtonSelected"><font size="1" face="verdana">&nbsp;&nbsp;Naming&nbsp;&nbsp;</font></td>
<%else%>
	<td><font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Naming', '<%= strID%>', '<%= sClass%>');">Naming</a>&nbsp;&nbsp;</font></td>
<%end if%>

</tr></table>
</div>

<%

'##############################################################################
'
' Display the SMR list Here
'
'##############################################################################
If strDisplayedList <> "SMR" Then
	Response.Write "<Table style=""Display:none"" ID=SMRTable><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
End if

'##############################################################################
'
' Display the observations here
'
'##############################################################################

If strDisplayedList <> "OTS" Then
	Response.Write "<Table style=""Display:none"" ID=OTSTable><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
Else

%>
<br><font face="verdana" size="2"><b>Open Observations</b></font>
<br><br>
<!--Links Here <BR><BR>-->

  <%
    dim blnOTSDown
    blnOTSDown = false
    on error resume next
	strSQL = "spListOTS4Root " & clng(strID) & ",1"
	rs.Open strSQL,cn,adOpenForwardOnly
    if cn.errors.count > 0 then
        blnOTSDown = true
    end if
    on error goto 0
  If blnOTSDown then
	Response.Write "<Table ID=OTSTable><TR><TD><font size=2>OTS is Currently Down</font></td></tr></table>"
  elseIf rs.EOF and rs.BOF then
	Response.Write "<Table ID=OTSTable><TR><TD><font size=2>None</font></td></tr></table>"
  else
  
  
  %><font face="verdana" size="1">

		<%if CurrentUserPartner=1 then%>
			<a target="DiretOTSLink" href="http://si.houston.hp.com/si/">Add New</a>
		<% else %>
			<a target="DiretOTSLink" href="https://prp.atlanta.hp.com/si/">Add New
            </a>
		<%end if%>
		&nbsp;|&nbsp;<a href="javascript:ShowOTSAdvanced('<%=strDeliverableName%>');">Search</a>
		&nbsp;|&nbsp;<a href="javascript:ShowOTSDetails();">Details</a>
		&nbsp;|&nbsp;<a href="javascript:window.print();">Print</a>
		&nbsp;|&nbsp;<a href="javascript:Export(1);">Export</a>
		<br><br>
</font>

<table ID="OTSTable" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <thead>
  <tr>
		<td nowrap width="20" bgColor="cornsilk" vAlign="center"><input type="checkbox" id="chkAllOTS" name="chkAllOTS" Language="javascript" onclick="onclick_ResetOTS();" style="WIDTH:16;HEIGHT:16"></td>	    
		<td onclick="SortTable( 'OTSTable', 1 ,1,2);" nowrap width="30" bgColor="cornsilk" vAlign="center"><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1"><strong>ID</strong></font></td>
	    <td onclick="SortTable( 'OTSTable', 2 ,0,2);" nowrap width="20" bgColor="cornsilk" vAlign="center"><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1"><strong>Found&nbsp;On</strong></font></td>
	    <td onclick="SortTable( 'OTSTable', 3 ,1,2);" nowrap width="20" bgColor="cornsilk" vAlign="center"><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1"><strong>PR</strong></font></td>
	    <td onclick="SortTable( 'OTSTable', 4 ,0,2);" nowrap width="20" bgColor="cornsilk"><strong><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1">State</font></strong></td>
	    <td onclick="SortTable( 'OTSTable', 5 ,0,2);" nowrap width="20" bgColor="cornsilk"><strong><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1">Owner</font></strong></td>
	    <td onclick="SortTable( 'OTSTable', 6 ,0,2);" nowrap width="110" bgColor="cornsilk"><strong><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1">Product</font></strong></td>
	    <td onclick="SortTable( 'OTSTable', 7 ,0,2);" nowrap width="110" bgColor="cornsilk"><strong><font onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();" size="1">Summary</font></strong></td>
 </tr>
 </thead>
<%
	do while not rs.EOF
		%>
		<tr LANGUAGE="javascript" onmouseout="return versionrow_onmouseout()" onmouseover="return versionrow_onmouseover()" onclick="OTSDetails('<%=rs("ObservationID")%>')">
			<td nowrap width="20" vAlign="center"><input type="checkbox" style="WIDTH:16;HEIGHT:16" id="chkOTSID" name="chkOTSID" value="<%=rs("ObservationID")%>"></td>
			<td nowrap class="cell"><font size="1" class="text"><%=rs("ObservationID")%></td>
			<td nowrap class="cell"><font size="1" class="text"><%=rs("OTSComponentVersion")%></td>
			<td nowrap class="cell"><font size="1" class="text"><%=rs("Priority")%></td>
			<td nowrap class="cell"><font size="1" class="text"><%=rs("State")%></td>
			<td nowrap class="cell"><font size="1" class="text"><%=shortname(rs("OwnerName") & "")%></td>
			<td nowrap class="cell"><font size="1" class="text"><%=rs("Product")%></td>
			<td class="cell"><font size="1" class="text"><%=rs("Summary")%></td>
		</tr>
		<%
		rs.MoveNext
	loop
%>

 
 </table>
<%
	end if
	if not blnOTSDown then
	    rs.Close
	end if
End if



'##############################################################################
'
' Display the Restrictions List Here
'
'##############################################################################

If strDisplayedList <> "Restriction" Then
	Response.Write "<Table style=""Display:none"" ID=RestrictionTable><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
Else

%>
<br><font face="verdana" size="2"><b>Restricted - Supply Chain</b></font>
<br><br>
<%if blnSysAdmin or blnPM then%>
	<font color="blue" size="1" id="AddRestrictions" LANGUAGE="javascript" onmouseover="return AddVersion_onmouseover()" onmouseout="return AddVersion_onmouseout()" onclick="return AddRestrictions_onclick()">
	<u>Update Restrictions</u></font> | 
<%end if%>
<font size="1" face="verdana"><a target="_blank" href="Deliverable/QuickReports.asp?Report=12">View All Restrictions</a></font>
 <br> <br>
 
 <%
	rs.open "spListCommodityRestrictions " & clng(DRID),cn,adOpenStatic
	if rs.eof and rs.bof then
		Response.write "No restrictions found for this component."
	else
		Response.Write "<TABLE style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bordercolor=tan bgcolor=ivory border=1><TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></TD><TD><font size=1 face=verdana><b>Product</b></font></TD><TD><font size=1 face=verdana><b>Vendor</b></font></TD><TD><font size=1 face=verdana><b>Component</b></font></TD><TD><font size=1 face=verdana><b>HW,FW,Rev</b></font></TD><TD><font size=1 face=verdana><b>Model</b></font></TD><TD><font size=1 face=verdana><b>Part</b></font></font></TD></TR>"
		do while not rs.eof
			strVersion = rs("Version")
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			Response.Write "<TR><TD><font size=1 face=verdana>" & rs("DelID") & "</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & rs("Product") & "</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & rs("Vendor") & "</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & rs("Deliverable") & "</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & strVersion & "&nbsp;</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & rs("ModelNumber") & "</font></TD>"
			Response.Write "<TD><font size=1 face=verdana>" & rs("PartNumber") & "</font></TD></tr>"
			rs.movenext
		loop
		rs.close
		Response.Write "</table>"
	end if
 %>
 
<%
end if

'##############################################################################
'
' Display the Version List Here
'
'##############################################################################

If strDisplayedList <> "Versions" Then
	Response.Write "<Table style=""Display:none"" ID=VersionTable><TR><TD><b><font color=red size=2>Not Displayed</font><b></td></tr></table>"
Else

%>
<br><font face="verdana" size="2"><b>Versions</b></font>
<br><br>
<%if clng(currentuserid) <> 31 and clng(currentuserid) <> 8 and clng(currentuserid) <> 1396 and clng(currentuserid)<>clng(strDevManagerID) and clng(currentuserid)<>clng(DevPMManagerID) and blnLockReleases then%>
<font color="blue" size="1" id="AddVersion" LANGUAGE="javascript" onmouseover="return AddVersion_onmouseover()" onmouseout="return AddVersion_onmouseout()" onclick="alert('New versions of this component can only be added by <%=longname(strmanager) & strManagerName%>.');">
<u>Add New Version</u></font>&nbsp;|&nbsp;
<%else%>
<font color="blue" size="1" id="AddVersion" LANGUAGE="javascript" onmouseover="return AddVersion_onmouseover()" onmouseout="return AddVersion_onmouseout()" onclick="return AddVersion_onclick(0)">
<u>Add New Version</u></font>&nbsp;|&nbsp;
<%end if%>
<font size="1">
<a target="blank" href="Deliverable\DeliverableRootChange.asp?ID=<%=DRID%>&amp;VersionID=0">View History</a>
<%if trim(strType) <> "1" then%>
 | <a target="blank" href="query/DelReport.asp?lstRoot=<%=DRID%>&txtFunction=1&chkInImage=on&txtAdvanced=pv.productstatusid=3&txtTitle=Versions%20In%20Image%20on%20Production%20Products">Production Images</a> | <a target="blank" href="Query/DelImageReport.asp?ID=<%=DRID%>&txtTitle=Image%20History">Image History</a>
<%end if%>

 | <a href="javascript: ChooseColumns(<%=strType%>);">Choose Columns</a>

<span ID="EOLLink" style="Display:none">&nbsp;|&nbsp;

<%if request("ShowEOL") = "" then%>
	<a href="dmview.asp?Tab=<%=Server.HTMLEncode(sTab)%>&amp;ID=<%=DRID%>&amp;ShowEOL=1">Show Inactive Versions</a></span>
<%else%>
	<a href="dmview.asp?Tab=<%=Server.HTMLEncode(sTab)%>&amp;ID=<%=DRID%>">Hide Inactive Versions</a></span>
<%end if%>
</font>
 <br> <br>
  <%
    strSQL = "spListDeliverableVersions " & clng(strID) & ",0"
	rs.Open strSQL,cn,adOpenStatic
  
  dim strVersion
  dim blnPMR
  dim blnCD
  dim strReleaseTeamStatus
  If not(rs.EOF and rs.BOF) then
	InactiveCount=0
  %>

<table ID="VersionTable" cellSpacing="1" cellPadding="1" width="100%" border="1" borderColor="tan" bgColor="ivory">
  <thead>
  <tr>
	<%if blnSkipPopup then%>
	    <td nowrap width="70" bgColor="cornsilk" vAlign="center"><font size="1"><strong>Function</font></td>
	<%end if%>
	
	
	<%
	
        dim ColumnCount
        dim strColumns
        dim ColumnArray
        dim MasterColumnArray
        dim ColumnPartArray
        
        ColumnCount=0
        if strFilename = "HFCN" then
            strColumns = "ID:1,Version:1,Supplier:1,Vendor Version:1,Title:1,Supported On:1"
        elseif trim(strType) = "1" then
            strColumns = "ID:1,Supplier:1,Vendor:1,Version:1,Model:1,HW:1,FW:1,Rev:1,Code Name:0,Part Number:1,Factory EOA:1,Service EOA:1,Workflow:1,Supported:1,Targeted:0,Samples Available:0,RoHS/Green Spec:0"
        else
            strColumns = "ID:1,Version:1,Vendor Version:1,Level:1,ISO:1,SMR:1,Workflow:1,Created:1,Completed:1,Targeted:1"
        end if
        MasterColumnArray = split(strColumns,",")
        
		set rs2 = server.CreateObject("ADODB.recordset")
        if trim(strType) = "1" then
            rs2.open "spGetEmployeeUserSettings " & currentuserid & ",6",cn
        else
            rs2.open "spGetEmployeeUserSettings " & currentuserid & ",5",cn
        end if
        if rs2.eof and rs2.bof then
        else
            strColumns = rs2("Setting") & ""
            for i = 0 to ubound(masterColumnArray)
                ColumnPartArray = split(mastercolumnarray(i),":")
                if instr("," & strcolumns,"," & trim(ColumnPartArray(0)) & ":") < 1 then
                    if strColumns = "" then
                        strColumns = trim(ColumnPartArray(0)) & ":1"
                    else
                        strColumns = strColumns & "," & trim(ColumnPartArray(0)) & ":1"
                    end if
                end if
            next
        end if

        response.write "<input id=""txtColumnList"" style=""display:none"" name=""txtColumnList"" type=""text"" value=""" & strColumns & """>"
        ColumnArray = split(strColumns,",")

        rs2.close	
        set rs2=nothing

        for i = 0 to ubound(ColumnArray)
            ColumnPartArray = split(ColumnArray(i),":")
            if trim(lcase(ColumnPartArray(1))) = "1" then
                if trim(lcase(ColumnPartArray(0))) = "id" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,1,2);"" nowrap width=""30"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>ID</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "version" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""100"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Version</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "supplier" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Supplier</strong></font></td>"
                elseif trim(lcase(ColumnPartArray(0))) = "vendor" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Vendor</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "model" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Model</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "code name" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""70"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Code&nbsp;Name</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "hw" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""20"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>HW</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "fw" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""20"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>FW</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "rev" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""20"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Rev</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "part number" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Part&nbspNumber</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "vendor version" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Vendor&nbsp;Version</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "title" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""110"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Title</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "level" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""40"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Level</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "iso" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""30"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>ISO</strong></font></td>"                 
                elseif trim(lcase(ColumnPartArray(0))) = "workflow" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Workflow</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "rohs/green spec" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,0,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>RoHS/Green&nbsp;Spec</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "created" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,2,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Created</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "samples available" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,2,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Samples&nbsp;Avail.</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "completed" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,2,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Completed</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "factory eoa" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,2,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Factory&nbsp;EOA</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "service eoa" then
                    response.Write "<td onclick=""SortTable( 'VersionTable', " & i & " ,2,2);"" nowrap width=""80"" bgColor=""cornsilk""><font onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" size=""1""><strong>Service&nbsp;EOA</strong></font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "targeted" then
                    response.Write "<td width=""100"" bgColor=""cornsilk""><strong><font size=""1"">Targeted&nbsp;On</font></strong></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "supported" then
                    response.Write "<td width=""100"" bgColor=""cornsilk""><strong><font size=""1"">Supported&nbsp;On</font></strong></td>"            
                end if
            end if
        next
        
        %>
	    <td style="display:none"><strong><font size="1">Filename</font></strong></td>
	    <td style="Display:none"><strong><font size="1">ISO</font></strong></td>
	    <td style="Display:none"><strong><font size="1">Image Path</font></strong></td>
  </tr></thead>
  <%do while not rs.EOF
    strReleaseStepName = ""
    if left(rs("location") & "",14) = "Release Team (" and rs("CommercialReleaseStatus") = 1 and rs("ConsumerReleaseStatus") = 2 then 
        strReleaseTeamStatus = "Release&nbsp;Team&nbsp;(TDC:&nbsp;Done)"
        strReleaseStepName = "Release"
    elseif left(rs("location") & "",14) = "Release Team (" and rs("CommercialReleaseStatus") = 2 and rs("ConsumerReleaseStatus") = 1 then
        strReleaseTeamStatus = "Release&nbsp;Team&nbsp;(Houston:&nbsp;Done)"
        strReleaseStepName = "Release"
    elseif trim(rs("location") & "") = "Workflow Complete" then
        strReleaseTeamStatus = "Complete"
        strReleaseStepName = "Complete"
    else
        strReleaseTeamStatus = rs("location") & ""
        if instr(strReleaseTeamStatus,",") <> 0 then
            strReleaseStepName = "Workflow&nbsp;Step"
        else
            if instr(strReleaseTeamStatus,"(") <> 0 then
                strReleaseStepName  = trim(left(strReleaseTeamStatus,instr(strReleaseTeamStatus,"(")-1))
            else
                strReleaseStepName  = strReleaseTeamStatus
            end if
        end if
    end if 

    'Lookup Build Level
    if trim(rs("LevelID") & "") = "" then
        strBuildLevel = "&nbsp;"
    elseif trim(rs("certificationstatus") & "") = "2" or trim(rs("certificationstatus") & "") = "4" then
        strBuildLevel = "WHQL"
    else
        set rs2 = server.CreateObject("ADODB.recordset")
	    rs2.open "spGetDeliverableBuildLevel " & rs("LevelID"),cn
    	if rs2.eof and rs2.bof then
    	    strBuildLevel = "&nbsp;"
    	else
    	    strBuildLevel = trim(rs2("name") & "")
    	end if
    	rs2.close
    	set rs2=nothing
    end if


	
	if request("ShowEOL") = "" and not rs("Active") then
		InactiveCount=InactiveCount + 1
	else
		strTitle = rs("comments") & "&nbsp;"

		strVersion = rs("Version")
		if rs("Revision") & "" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass") & "" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if
		
		strReleaseDate = trim(rs("actualreleasedate") & "")
		if strReleaseDate <> "" then
		    if isdate(strReleaseDate) then
		        strReleaseDate = formatdatetime(strReleaseDate,vbshortdate)
		    end if
		elseif trim(rs("Location") & "") = "Workflow Complete" then
		    strReleaseDate = formatdatetime(rs("Created"),vbshortdate)
		else
		    strReleaseDate = "&nbsp;"
		end if
		if blnPreinstallGroup and trim(strType) <> "1" and trim(rs("PreinstallInternalRev") & "") <> "" then
			strVersion = strVersion & "&nbsp;[Rev:&nbsp;" & rs("PreinstallInternalRev") & "" & "]"
		end if
		if rs("Vendor") & "" = "< Multiple Suppliers >" then
			strVendor = "&nbsp;"
		else
			strVendor = rs("Vendor") & "&nbsp;"
		end if
  
	
			blnCD=0
	
  
  dim strColor
  if rs("Active") then
	strColor = "ivory"
  else
	strcolor = "gainsboro"
	InactiveCount=InactiveCount + 1
  end if
  
 if trim(strProcurementGroup) <> "0" then%>
		<tr bgcolor="<%=strColor%>" LANGUAGE="javascript" onmouseout="return versionrow_onmouseout()" onmouseover="return versionrow_onmouseover()" oncontextmenu="javascript:contextMenu2(<%=strID%>,<%=rs("VersionID")%>,<%=trim(strType)%>,<%=trim(strProcurementGroup)%>);return false;" onclick="javascript:contextMenu2(<%=strID%>,<%=rs("VersionID")%>,<%=trim(strType)%>,<%=trim(strProcurementGroup)%>);">
	<%else%>
		<tr bgcolor="<%=strColor%>" LANGUAGE="javascript" onmouseout="return versionrow_onmouseout()" onmouseover="return versionrow_onmouseover()" oncontextmenu="javascript:contextMenu(<%=strID%>,<%=rs("VersionID")%>,<%=trim(strType)%>,<%=blnPM%>,<%=blnAccessoryPM%>,<%=blnCD%>);return false;" onclick="javascript:contextMenu(<%=strID%>,<%=rs("VersionID")%>,<%=trim(strType)%>,<%=blnPM%>,<%=blnAccessoryPM%>,<%=blnCD%>);">
	<%end if
	
	dim strCaption
	strCaption = "Edit"

	if  blnSkipPopup then
	%>
		<td valign="top"><font size="1" class="link"><a href="javascript:ReleaseVersion(<%=strID%>,<%=rs("VersionID")%>)">Release</a></font></td>
	<%end if
	
	set rs2 = server.CreateObject("ADODB.recordset")
	strSQL = "spGetTargetedProductsForVersion " & rs("VersionID")
	rs2.Open strSQL,cn,adOpenStatic
  
	strTargetedProducts = ""
	strAllProducts = ""
	do while not rs2.EOF
		strAllProducts = strAllProducts & ", " & rs2("Family") & "&nbsp;" & rs2("Version")
		if rs2("Targeted") then
			strTargetedProducts = strTargetedProducts & ", " & rs2("Family") & "&nbsp;" & rs2("Version")
		end if
		rs2.MoveNext
	loop
	rs2.close
	set rs2 = nothing			
	
	if strAllProducts = "" then
		strAllProducts = "None"
	elseif left(strAllProducts,1) = "," then
		strAllProducts = mid(strAllProducts,2)
	end if
	if strTargetedProducts = "" then
		strTargetedProducts = "None"
	elseif left(strTargetedProducts,1) = "," then
		strTargetedProducts = mid(strTargetedProducts,2)
	end if

    if trim(rs("Rohs") & "") = "" and trim(rs("Greenspec") & "") = "" then
    	strGreenSpec = "&nbsp;"
    elseif trim(rs("Rohs") & "") <> "" and trim(rs("Greenspec") & "") <> "" then
    	strGreenSpec = rs("Rohs") & "_" & rs("GreenSpec") 
    elseif trim(rs("Rohs") & "") <> "" then
    	strGreenSpec = rs("Rohs")
    else
    	strGreenSpec = rs("GreenSpec")
    end if    
    
    if trim(strType) = "1" then
	    strVendor = rs("Vendor") & ""
	    if rs("SupplierCode") & "" <> "" and rs("SupplierCode") & "" <> "TBD" then
		    strVendor = strVendor & " ("  & rs("SupplierCode") & ")"
	    end if

	    strVersion = rs("Version") & ""
    	
        if trim(rs("EOLDate") & "") = "" then
            strEOLDate = "&nbsp;"
        else
            strEOLDate = rs("EOLDate")
        end if
        if trim(rs("ServiceEOADate") & "") = "" then
            strServiceEOADate = "&nbsp;"
        else
            strServiceEOADate = rs("ServiceEOADate")
        end if
    end if
    
        for i = 0 to ubound(ColumnArray)
            ColumnPartArray = split(ColumnArray(i),":")
            if trim(lcase(ColumnPartArray(1))) = "1" then
                if trim(lcase(ColumnPartArray(0))) = "id" then
                    response.Write "<td class=""cell""><font size=""1"" class=""text"">" & rs("VersionID") & "</font></td>"
                elseif trim(lcase(ColumnPartArray(0))) = "supplier" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("Supplier") & "&nbsp;</font></td>"             
                elseif trim(lcase(ColumnPartArray(0))) = "vendor" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strVendor & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "model" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("ModelNumber") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "version" then
                    response.Write "<td class=""cell""><font size=""1"" class=""text"">" & strVersion & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "vendor version" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("VendorVersion") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "rohs/green spec" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strGreenSpec & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "code name" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("CodeName") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "level" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strBuildLevel & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "iso" then
                    response.Write "<td align=center nowrap class=""cell""><font size=""1"" class=""text"">" & replace(replace(rs("ISOImage")  & "","1","X"),"0","") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "hw" then
                    response.Write "<td class=""cell""><font size=""1"" class=""text"">" & strVersion & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "fw" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("Revision") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "rev" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("Pass") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "part number" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & rs("PartNumber") & "&nbsp;</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "factory eoa" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strEOLDate & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "service eoa" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strServiceEOADate & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "workflow" then
                    if trim(strType) = "1" then
                        response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & replace(rs("location") & "","Workflow Complete","Complete") & "&nbsp;</font></td>"            
                    else
                        response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strReleaseTeamStatus & "&nbsp;</font></td>"            
                    end if
                elseif trim(lcase(ColumnPartArray(0))) = "supported" then
                    response.Write "<td nowrap class=""cell"" ID=""Products" & trim(rs("VersionID"))& """><font size=""1"" class=""text"">" & strAllProducts & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "targeted" then
                    response.Write "<td nowrap class=""cell"" ID=""Products" & trim(rs("VersionID"))& """><font size=""1"" class=""text"">" & strTargetedProducts & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "title" then
                    response.Write "<td class=""cell""><font size=""1"" class=""text"">" & strTitle & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "created" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & formatdatetime(rs("created"),vbshortdate) & "</font></td>"            
                elseif trim(lcase(ColumnPartArray(0))) = "samples available" then
                    if isnull(rs("SampleDate")) then
                        response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">&nbsp;</font></td>"            
                    else
                        response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & formatdatetime(rs("SampleDate"),vbshortdate) & "</font></td>"            
                    end if
                elseif trim(lcase(ColumnPartArray(0))) = "completed" then
                    response.Write "<td nowrap class=""cell""><font size=""1"" class=""text"">" & strReleaseDate & "</font></td>"            
                end if
            end if
        next
        %>
		<td style="Display:none" nowrap class="cell"><font size="1" class="text" ID="File<%=rs("VersionID")%>"><%=rs("Filename") & "&nbsp;"%></font></td>
		<td style="Display:none" class="cell" ID="ISO<%=rs("VersionID")%>"><font size="1" class="text"><%=rs("IsoImage")%></font></td>
		<td style="Display:none" class="cell" ID="Path<%=rs("VersionID")%>"><font size="1" class="text"><%=rs("ImagePath")%></font></td>
        <td style="Display:none" class="cell" ID="ReleaseStep<%=rs("VersionID")%>"><%=replace(strReleaseStepName," ","&nbsp;")%></td>
  </tr>
  
  <%
	end if	
	rs.MoveNext
	loop
	%>

</table>
<%
	if InactiveCount = 1 then
		Response.Write "<font size=1 color=red><BR><BR>There is one inactive version of this component.</font>"
	elseif Inactivecount > 0 then
		Response.Write "<font size=1 color=red><BR><BR>There are " & InactiveCount & " inactive versions of this component.</font>"
	end if

%>

<%else%>
<b><font face="Verdana" size="2">No versions found for this component.</font></b>
 
<%end if
		rs.Close
		set rs = nothing
		cn.close
		set cn= nothing
end if
End If 'Display Versions

'##############################################################################
'
' Display the Agency Matrix Here
'
'##############################################################################

If strDisplayedList = "Agency" Then
%>
	<br>	
	<%
	If sDmStatus = "All" Then
	%>
        <font face="verdana" size="2"><strong>Certification - All Countries</strong></font><br><br>
	    <font size="1">
		<a id="AddVersion" HREF="javascript:void(0);" onclick="return SetDisplayStatus('Supported');">Show Supported Countries</a></font><br><br>
	<%
	Else
	%>
        <font face="verdana" size="2"><strong>Certification - Supported Countries</strong></font><br><br>
	    <font size="1">
		<a id="AddVersion" HREF="javascript:void(0);" onclick="return SetDisplayStatus('All');">Show All Countries</a></font><br><br>
	<%
	End If
	%>
	
	</font>
	<%
    Response.Write "<div id=""GridViewContainer"" class=""GridViewContainer"" style=""width: 100%; height: 525px;"">"
	Response.Write "<Table ID=TableAgency width=100% border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>"
	Call DrawDMViewMatrix(DRID, blnAgencyDataMaintainer)
	Response.Write "</table><p>* Country added after POR by DCR</p>"
    Response.Write "</div>"

End If

'##############################################################################
'
' Display the Document List Here
'
'##############################################################################

If strDisplayedList = "Documents" Then
%>
	<br>
	<font face="verdana" size="2"><strong>Component Documents</strong></font><br><br>

	<table ID="TableAgency" width="100%" cellpadding="2" cellspacing="1">
	<tr><td><img SRC="images/ICON-DOC-HTML.GIF" WIDTH="16" HEIGHT="16">&nbsp;<a href="javascript: openReport('DelVerSupportedCountriesByRoot', <%=strID%>);">Countries Supported by Version</a></td></tr>
	</table>
<%
End If
%> 
<%else%>
<a id="RFLink" href="javascript:RemoveFavorites(<%=DRID%>)"><font face="verdana" size="1"></font></a>
<a id="AFLink" style="Display:none" href="javascript:AddFavorites(<%=DRID%>)"><font face="verdana" size="1"></font></a>

<%end if%>
<input type="hidden" id="txtID" name="txtID" value="<%=DRID%>">
<input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserID%>">
<input type="hidden" id="txtUserSite" name="txtUserSite" value="<%=CurrentUserSite%>">
<input type="hidden" id="txtUserPartner" name="txtUserPartner" value="<%=trim(CurrentUserPartner)%>">
<input type="hidden" id="txtFavs" name="txtFavs" value="<%=strFavs%>">
<input type="hidden" id="txtFavCount" name="txtFavCount" value="<%=strFavCount%>">

<form id="ExportForm" target="new" method="post" action="mobilese/today/excelexport.asp">
<textarea ID="txtData" name="txtData" style="Display:none" rows="100" cols="20"></textarea>
</form>

<%
'##############################################################################
'
' Display the Product List Here
'
'##############################################################################

If strDisplayedList = "Products" Then

	dim strProductStatus
	dim strProductStatusName
	dim strSelectedBusiness
	dim strSelectedBusinessName
	
	strProductStatusName=""
	strSelectedBusinessName = ""
	if iShowProducts <> "1" and iShowProducts <> "2" and trim(iShowProducts) <> "" then 'All
		strProductStatus = "All"
		strProductStatusName = "All"
	else
		strProductStatus = "<a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowProducts=0&ShowBusiness=" & Server.HTMLEncode(iShowBusiness) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>All</a>"
	end if
	if iShowProducts = "1" or iShowProducts = "" then 'Active
		strProductStatus = strProductStatus & " , Active"
		strProductStatusName = "Active"
	else
		strProductStatus = strProductStatus & " , <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowProducts=1&ShowBusiness=" & Server.HTMLEncode(iShowBusiness) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>Active</a>"
	end if
	if iShowProducts = "2" then 'Inavtive
		strProductStatus = strProductStatus & ", Inactive"
		strProductStatusName = "Inactive"
	else
		strProductStatus = strProductStatus & " , <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowProducts=2&ShowBusiness=" & Server.HTMLEncode(iShowBusiness) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>Inactive</a>"
	end if

	if iShowBusiness <> "1" and iShowBusiness <> "2" then
		strSelectedBusiness = "All"
		strSelectedBusinessName = " "
	else
		strSelectedBusiness = "<a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowBusiness=0&ShowProducts=" & Server.HTMLEncode(iShowProducts) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>All</a>"
	end if
	if iShowBusiness = "1"  then
		strSelectedBusiness = strSelectedBusiness & ", Commercial"
		strSelectedBusinessName = " Commercial "
	else
		strSelectedBusiness = strSelectedBusiness & ", <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowBusiness=1&ShowProducts=" & Server.HTMLEncode(iShowProducts) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>Commercial</a>"
	end if
	if iShowBusiness = "2"  then
		strSelectedBusiness = strSelectedBusiness & ", Consumer"
		strSelectedBusinessName = " Consumer "
	else
		strSelectedBusiness = strSelectedBusiness & ", <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowBusiness=2&ShowProducts=" & Server.HTMLEncode(iShowProducts) & "&HideHeader=" & Server.HTMLEncode(hideHeader) & """>Consumer</a>"
	end if

	%>
	<table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
	<tr>
		<td valign="top">
			<table><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		</td>
		<td width="100%">
			<table>
				<tr><td nowrap><b>Business:</b></td><td width="100%"><%=strSelectedBusiness%></td></tr>
				<tr><td><b>Status:</b></td><td width="100%"><%=strProductStatus%></td></tr>
			</table>
		</td>
	</tr>
	</table>

	<br>
  <%
	strID = clng(DRID)
	if iShowProducts = "" then
		strSQL = "spGetProductsForRoot " & clng(strID) & ",1"
	elseif iShowProducts = "1" or iShowProducts = "2" then
		strSQL = "spGetProductsForRoot " & clng(strID) & "," & Server.HTMLEncode(iShowProducts)
	else
		strSQL = "spGetProductsForRoot " & clng(strID) & ",0" 
	end if

	rs.Open strSQL,cn,adOpenForwardOnly
	strAvailableProducts = ""
	strInactiveProducts = ""
	dim strProductIDList
	dim strProductNameList
	dim strProductDevCenter
	dim strProductSeriesSummary
	strProductIDList = ""
	strProductNameList = ""
	strProductDevCenter = ""
	strProductSeriesSummary = ""

	do while not rs.EOF
		if trim(strSelectedBusinessName) = "" or (trim(strSelectedBusinessName)="Consumer" and trim(rs("DevCenter")) = "2")  or (trim(strSelectedBusinessName)="Commercial" and trim(rs("DevCenter")) <> "2" and trim(rs("DevCenter")) <> "0") then
			
            dim strBusinessName
            if rs("devcenter") = 0 then
				strBusinessName = "Not Specified"
			elseif rs("devcenter") = "2" then
				strBusinessName = "Consumer"
			else
				strBusinessName = "Commercial"
			end if

            strAvailableProducts = strAvailableProducts & "<tr bgColor=""ivory""><td nowrap vAlign=""center""><font size=""1"">" & rs("Name") & "&nbsp;" & "</font></td><TD nowrap><font size=""1""><a ID=""SupportCount" & rs("ID") & """ href=""javascript:ChooseVersions(" & clng(strID) & "," & rs("ID") & "," & rs("ReleaseID") & ");"">" & rs("VersionCount") & "</a></font></TD><td nowrap vAlign=""center""><font size=""1"">" & rs("VersionTargetedCount") & "&nbsp;" & "</font></td><td nowrap vAlign=""center""><font size=""1"">" & strBusinessName & "&nbsp;" & "</font></td><td nowrap vAlign=""center""><font size=""1"">" & rs("SeriesSummary") & "&nbsp;" & "</font></td></tr>"
		end if
		rs.MoveNext
	loop
	rs.close		
	
%>
	
  <% if strAvailableProducts = "" then %>
    <p><font size="2"><strong>No product is using this component now</strong></font></p>
  <% else %>
    <p><font size="2"><strong><%=strProductStatusName%><%=strSelectedBusinessName%>products using this component</strong></font></p>
    <table ID="TableProducts" width="60%" bordercolor="tan" border="1" cellpadding="2" cellspacing="1">
        <thead>
            <tr><td nowrap bgColor="cornsilk" vAlign="center"><font size="1"><b>Products</b></td><td nowrap bgColor="cornsilk" vAlign="center"><font size="1"><b>Versions</b></td><td nowrap bgColor="cornsilk" vAlign="center"><font size="1"><b>Targeted</b></td><td nowrap bgColor="cornsilk" vAlign="center"><font size="1"><b>Business</b></td><td nowrap bgColor="cornsilk" vAlign="center"><font size="1"><b>Series</b></td></tr>
        </thead>
        <tbody>
            <%=strAvailableProducts%>
        </tbody>
    </table>
  <% end if %>
 
<%
End If
%> 

<%
'##############################################################################
'
' Display the Change Requests Here
'
'##############################################################################

If strDisplayedList = "ChangeRequests" Then

	dim strScrStatus
	dim strScrStatusName
	
	strScrStatusName=""

	if iShowScr <> "1" and iShowScr <> "2" and trim(iShowScr) <> "" then 'All
		strScrStatus = "All"
		strScrStatusName = "All"
	else
		strScrStatus = "<a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowScr=0"">All</a>"
	end if
	if iShowProducts = "1" or iShowProducts = "" then 'Active
		strScrStatus = strScrStatus & " , Active"
		strScrStatusName = "Active"
	else
		strScrStatus = strScrStatus & " , <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowScr=1"">Open</a>"
	end if
	if iShowProducts = "2" then 'Inavtive
		strScrStatus = strScrStatus & ", Inactive"
		strScrStatusName = "Inactive"
	else
		strScrStatus = strScrStatus & " , <a href=""dmview.asp?Tab=" & Server.HTMLEncode(sTab) & "&ID=" & Server.HTMLEncode(DRID) & "&ShowScr=2"">Closed</a>"
	end if


	%>
	<table class="DisplayBar" Width="100%" CellSpacing="0" CellPadding="2">
	<tr>
		<td valign="top">
			<table><tr><td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		</td>
		<td width="100%">
			<table>
				<tr><td><b>Status:</b></td><td width="100%"><%=strScrStatus%></td></tr>
			</table>
		</td>
	</tr>
	</table>
<br />
<table id="TableChangeRequests" width="50%" bordercolor="tan" border="1" cellpadding="2" cellspacing="1" style="display:none;">
<tr><th>Number</th><th>Submitter</th><th>Owner</th><th>Status</th><th>Summary</th></tr>
<tr></tr>
</table>
<% End If 

'##############################################################################
'
' Display Deliverable Names
'
'##############################################################################

If strDisplayedList = "Naming" Then
    Dim Names : Names = "Engineering,GPG (40-char SA),GPG-PhWeb (40-char AV),ZSRP (29-char AV),GPSy (40-char AV),GPSy (200-char AV),PMG (100-char AV),PMG (250-char AV)"
    Names = split(Names,",")
    Dim i : i = 0
	rs.open "usp_SelectDeliverableNames " & clng(DRID),cn,adOpenStatic
	Response.Write "<TABLE style=""WIDTH:100%"" cellpadding=2 cellspacing=1 bordercolor=tan bgcolor=ivory border=1><TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>Type</b></font></TD><TD><font size=1 face=verdana><b>Name</b></font></TD></TR>"
	for i = 0 to 7
		Response.Write "<TR><TD><font size=1 face=verdana>" & Names(i) & "</font></TD>"
		    if i = 0 then
			    Response.Write "<TD><font size=1 face=verdana>" & rs("Name") & "</font></TD></TR>"
			else
			    Response.Write "<TD><font size=1 face=verdana>" & rs("Name" & i+1) & "&nbsp;</font></TD></TR>"
			end if
	next
	rs.close
	Response.Write "</table>"
End If
%>
<input type="hidden" id="txtView" name="txtView" value="<%=sView%>">

<%if blnPreinstallGroup then%>
	<input type="hidden" id="txtPreinstallGroup" name="txtPreinstallGroup" value="1">  
<%else%>
	<input type="hidden" id="txtPreinstallGroup" name="txtPreinstallGroup" value="0">  
<%end if%>
<div id=FirefoxPopup style="display:none;position:absolute;width:2px;height:2px;left:0px;top:0px;padding:0px;background:white;border:1px solid gainsboro;z-index:100"></div>
</body>
</html>

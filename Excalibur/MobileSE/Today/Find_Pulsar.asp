<%@ Language="VBScript" %>
<%
Response.CacheControl = "Private"
Response.Expires = 0
%>
<% 
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^ 0-9a-zA-Z,#.-]/()!"
    Dim sType : sType = regEx.Replace(Request.QueryString("Type"), "")
    Dim sFind : sFind = regEx.Replace(Request.QueryString("Find"), "")

%>
<html>
<head>
    <meta name="VI60_defaultClientScript" content="JavaScript" />
    <title>Find <%=sType%> - HP Restricted</title>
    <link rel="stylesheet" type="text/css" href="../Style/programoffice.css" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <script id="clientEventHandlersJS" language="javascript" type="text/javascript">
    <!--

    var oPopup = window.createPopup();

    function window_onload() {
        var strID;
        Progress.style.display = "none";
	   if (txtOnlyID.value != "0")
		    {
            strID = window.showModalDialog("action.asp?" + txtOnlyID.value, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No");
		    if (typeof(strID) != "undefined")
			    {

                txtOnlyID.value = "0";
                window.history.back(1);
            }
        }
    }

    function DeliverableAlertMenu(RootID, VersionID, PathID)
    {
        if (window.event.srcElement.className != "text" && window.event.srcElement.className != "cell")
            return;

        // The variables "lefter" and "topper" store the X and Y coordinates
        // to use as parameter values for the following show method. In this
        // way, the popup displays near the location the user clicks. 
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;
        var strFilename;
        var strPath = document.all("Path" + PathID + "_" + VersionID).innerText;

        //	if(typeof(oPopup) == "undefined") 
        //		{
        //		DisplayVersion(RootID,VersionID);
        //		
        //		return;
        //		}

        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayChanges(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Display Changes</SPAN></font></DIV>";


      if (strPath != "")
	    {
            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<font face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:GetVersion(" + VersionID + ")'\" >&nbsp;&nbsp;&nbsp;Download</SPAN></font></DIV>";
        }

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayRoot(" + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Display&nbsp;Versions</SPAN></font></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayRootProperties(" + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Root&nbsp;Properties</SPAN></font></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + RootID + "," + VersionID + ")'\" ><font face=Arial size=2>&nbsp;&nbsp;&nbsp;Version&nbsp;Properties</font></SPAN></DIV>";

        popupBody = popupBody + "</DIV>";


        var NewHeight;
        var NewWidth;

        if (typeof (oPopup) == "undefined") {
            mnuPopup.style.display = "";
            mnuPopup.innerHTML = popupBody;

            mnuPopup.style.width = mnuPopup.scrollWidth + 10;
            mnuPopup.style.height = mnuPopup.scrollHeight;
            mnuPopup.style.left = lefter;
            mnuPopup.style.top = topper;
        }
        else {
            oPopup.document.body.innerHTML = popupBody;
            oPopup.show(lefter, topper, 170, 86, document.body);

            NewHeight = oPopup.document.body.scrollHeight;
            NewWidth = oPopup.document.body.scrollWidth;
            oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
        }

    }


    function RootAlertMenu(RootID, ReportType)
    {
        if (window.event.srcElement.className != "text" && window.event.srcElement.className != "cell")
            return;

        // The variables "lefter" and "topper" store the X and Y coordinates
        // to use as parameter values for the following show method. In this
        // way, the popup displays near the location the user clicks. 
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;
        var strFilename;

        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	    if (ReportType!=1)
		    {
            popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
            popupBody = popupBody + "<font face=Arial size=2>";
            popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayMatrix(" + RootID + "," + ReportType + ")'\" >&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Matrix</SPAN></font></DIV>";

            popupBody = popupBody + "<DIV>";
            popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";
        }

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayRoot(" + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Display&nbsp;Versions</SPAN></font></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayRootProperties(" + RootID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></font></DIV>";

        popupBody = popupBody + "</DIV>";

        oPopup.document.body.innerHTML = popupBody;
        //	if (strPath != "")
        //		oPopup.show(lefter, topper, 170, 102, document.body);
        //	else
        oPopup.show(lefter, topper, 170, 86, document.body);

        //Adjust window size
        var NewHeight;
        var NewWidth;

        NewHeight = oPopup.document.body.scrollHeight;
        NewWidth = oPopup.document.body.scrollWidth;
        oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
    }

    function DisplayMatrix(RootID, ReportType) {

        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 900) / 2
        NewTop = (screen.height - 100) / 2

        window.open("../../Deliverable/HardwareMatrix.asp?lstRoot=" + RootID + "&Subassembly=1", "_blank", "Width=900,Height=500,menubar=yes,toolbar=yes,resizable=Yes,status=yes,scrollbars=yes")

    }

    function DisplayVersion(RootID, VersionID) {
        var strResult;
    strResult = window.showModalDialog("../../WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID, "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	    if (typeof(strResult) != "undefined")
		    {
            window.location.reload(true);
        }
    }

    function DisplayRoot(RootID) {
        window.location = "../../DMView.asp?ID=" + RootID;
    }

    function DisplayRootProperties(RootID) {
        var strResult;
        strResult = window.showModalDialog("../../Root.asp?ID=" + RootID, "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
	    if (typeof(strResult) != "undefined")
		    {
            window.location.reload(true);
        }
    }

    function DisplayChanges(ID) {
        strID = window.showModalDialog("../../Properties/FT.asp?ID=" + ID, "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
    }

    function GetVersion(VersionID) {
        window.open("../../FileBrowse.asp?ID=" + VersionID);
    }

    function ProductMenu(ID) {
        window.location.href = "../../pmview.asp?Class=1&ID=" + ID;
    }

    function contextMenu(RootID, TypeID)
    {
        // The variables "lefter" and "topper" store the X and Y coordinates
        // to use as parameter values for the following show method. In this
        // way, the popup displays near the location the user clicks. 
        var lefter = event.clientX;
        var topper = event.clientY;
        var popupBody;

		    if (window.event.srcElement.className == "text")
			    {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                    SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
		    else if (window.event.srcElement.className == "cell")
	    	    {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement)
                    SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }


        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionPrint(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Print&nbsp;Preview...</SPAN></font></DIV>";


        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionMail(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Send&nbsp;Email...</SPAN></font></DIV>";


        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";



        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<font face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ActionProperties(" + RootID + "," + TypeID + ")'\" >&nbsp;&nbsp;&nbsp;Properties</SPAN></font></DIV>";

        popupBody = popupBody + "</DIV>";

        oPopup.document.body.innerHTML = popupBody;

        oPopup.show(lefter, topper, 130, 85, document.body);
    }


    function ActionProperties(strID, strType) {
        var strResult;
        strResult = window.showModalDialog("action.asp?ID=" + strID + "&Type=" + strType, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;maximize:Yes;;center:Yes; help: No;resizable: Yes;status: No")
	    if (typeof(strResult) != "undefined")
		    {
            window.location.reload(true);
        }
    }

    function ActionPrint(strID, strType) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        //strResult = window.showModalDialog("mobilese/today/actionReport.asp?ID=" + strID + "&Type=" + strType,"","dialogWidth:655px;dialogHeight:650px;maximize:Yes;edge: Sunken;center: Yes; help: No;resizable: Yes;status: No") 
        strResult = window.open("actionReport.asp?Action=0&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No")
    }

    function ActionMail(strID, strType) {
        var strResult;
        var NewTop;
        var NewLeft;

        NewLeft = (screen.width - 655) / 2
        NewTop = (screen.height - 650) / 2

        //strResult = window.showModalDialog("mobilese/today/actionReport.asp?ID=" + strID + "&Type=" + strType,"","dialogWidth:655px;dialogHeight:650px;maximize:Yes;edge: Sunken;center: Yes; help: No;resizable: Yes;status: No") 
        strResult = window.open("actionReport.asp?Action=1&ID=" + strID + "&Type=" + strType, "_blank", "Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=no,toolbar=no,resizable=Yes,status=No")
    }



    function rows_onmousemove() {
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
	    else if (window.event.srcElement.className == "bold")
    	    {
            window.event.srcElement.parentElement.parentElement.style.color = "red";
            window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
        }

    }

    function rows_onmouseout() {
        if (window.event.srcElement.className == "text")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "bold")
            window.event.srcElement.parentElement.parentElement.style.color = "black";
        else if (window.event.srcElement.className == "cell")
            window.event.srcElement.parentElement.style.color = "black";
    }

    function rows_onclick() {
        var strID;
        var strDisplay;

	    if (window.event.srcElement.className == "text")
		    {
            strDisplay = window.event.srcElement.parentElement.parentElement.className;
            strID = window.showModalDialog("action.asp?" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No");
        }
	    else if (window.event.srcElement.className == "bold")
		    {
            strDisplay = window.event.srcElement.parentElement.parentElement.parentElement.className;
            strID = window.showModalDialog("action.asp?" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No");
        }
	    else if (window.event.srcElement.className == "cell")
		    {
            strDisplay = window.event.srcElement.parentElement.className;
            strID = window.showModalDialog("action.asp?" + strDisplay, "", "dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes;maximize:Yes; help: No;resizable: Yes;status: No");
        }
        if (typeof (strID) != "undefined")
            window.location.reload();
    }

    function window_onmouseup() {
        mnuPopup.style.display = "none";
    }
    //-->
    </script>
    <style type="text/css">
        td{
            FONT-FAMILY: Verdana;
            FONT-SIZE: xx-small
        }
    </style>
</head>
<body onload="return window_onload()" onmouseup ="window_onmouseup()">
<%
     
    set cn = server.CreateObject("ADODB.Connection")
    set rs = server.CreateObject("ADODB.recordset")
    cn.ConnectionString = Session("PDPIMS_ConnectionString")
    cn.IsolationLevel=256
    cn.Open

    function IsInteger( strValue)
        Set re = new RegExp
        re.Pattern = "^\d+$"
        re.Global = true
        IsInteger = re.test(strValue)
    end function

    dim CurrentDomain
    dim CurrentUser
    dim CurrentUserPartner
    dim CurrentUserPartnerName
    dim CurrentPartnerOTSGroupID
    dim CurrentUserOtherPartnerNames
    CurrentUserOtherPartnerNames = ""


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
    Set	rs = cm.Execute 

    set cm=nothing	
		
    if not (rs.EOF and rs.BOF) then
	
		CurrentUserID = rs("ID")
'		CurrentUserName = rs("Name") & ""
'		CurrentUserEmail = rs("Email") & ""
'		CurrentUserGroup = rs("WorkgroupID") & ""
'		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = rs("PartnerID") & ""
        CurrentUserOtherPartnerNames = rs("OtherPartnerNames") & ""
	else
		Response.Redirect "Excalibur.asp"
	end if
	rs.Close
	
	CurrentPartnerOTSGroupID = ""
	if CurrentUserPartner = 1 then
		CurrentUserPartnerName = "HP"
	elseif trim(CurrentUserPartner) = "9" then
		Response.Redirect "/mobilese/modusmain.asp"
	else
		rs.Open "spGetPartnerName " & CurrentUserPartner,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			CurrentUserPartnerName = ""
		else
			CurrentUserPartnerName = rs("Name") & ""
            CurrentPartnerOTSGroupID = trim(rs("OTSVendorGroupID") & "")
		end if		
		rs.Close
	end if
    set rs = nothing
    if CurrentPartnerOTSGroupID = "" then
       CurrentPartnerOTSGroupID = -1
    end if
    
    function LoopupPartNumberID(strPart)
		dim cnRel 
		dim rsRel
		
		set cnRel = server.CreateObject("ADODB.Connection")
		if CurrentDomain = "asiapacific" then
			cnRel.ConnectionString = "Provider=sqloledb;Data Source=16.159.144.33;Network Library=DBMSSOCN;Initial Catalog=OneStop;User ID=Excalibur;Password=Airport;" 
		else
			cnRel.ConnectionString = "Provider=sqloledb;Data Source=houhpqexcal02.americas.hpqcorp.net;Network Library=DBMSSOCN;Initial Catalog=OneStop;User ID=Excalibur;Password=Airport;" 
		end if
		cnRel.Open 

		set rsRel = server.CreateObject("ADODB.recordset")
		rsRel.Open "spGetCDs '" & strPart & "'",cnRel,adOpenForwardOnly		
		if rsRel.eof and rsRel.bof then
			LoopupPartNumberID = 0
		else
			LoopupPartNumberID = rsRel("ExcalID")
		end if
		rsRel.Close
		set rsRel = nothing
		cnRel.close
		set cnRel = nothing
	end function


	function isPartNumber(strPart)
		dim strTemp
		dim strBase
		dim strSpin
		
		strTemp = trim(strPart)
		isPartNumber = true
		if len(strTemp) <> 10 then
			isPartNumber = false
		else
			strBase = left(strTemp,6)
			strSpin = right(strTemp,3)
			if mid(strTemp,7,1) <> "-" then
				isPartNumber = false
			elseif not isnumeric(strBase) then			
				isPartNumber = false
			elseif not isnumeric(strSpin) then			
				isPartNumber = false
			end if
		end if
	end function

function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("%", "select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
	
	
  function GetVal(strText)
	if isnumeric(strText) then
		GetVal = clng(strText)
	else
		GetVal=0
	end if
  end function
  

  set rs = server.CreateObject("ADODB.recordset")
  set rs2 = server.CreateObject("ADODB.recordset")
  dim ColorIndex 
  dim strStatus
  dim strTargetDate
  dim strActualDate
  dim strType
  dim strDeliverable
  dim strProduct
  dim strFind
  dim intID
  dim strItems
  dim ItemCount
  dim LastID
  dim IDList
  dim i
  dim IDString
  dim blnAllNumeric
  dim cm
  dim p
  dim strRecordCount
  dim strFindScrubbed
  dim MaxMatchedDisplayed

    MaxMatchedDisplayed = 5000
  ItemCount = 0
  LastID = 0
  strRecordCount = -1
  
  
  
   strFind = replace(sFind," ",",")
'if currentuser="dwhorton" then
'    response.Write strFind
'    response.flush
'end if
if sType = "Issues" or  sType = "Change Requests" or sType = "Action Items" or sType = "AllActions" then

  'intID = getval(strFind)
  Dim LookupTypeID
  Select case sType
	case "Issues"
		LookupTypeID = "1"
	case "Change Requests"
		LookupTypeID = "3"
	case "Action Items"
		LookupTypeID = "2"
	case "AllActions"
		LookupTypeID = "0"
  end select

	IDList = split(strFind,",")
	  
	IDString = ""
	blnAllNumeric = true
	for i = 0 to Ubound(IDList) 
			if not (isinteger(IDList(i)) or trim(IDList(i))="") then
				blnAllNumeric = false
			end if
			if  trim(IDList(i))<>"" then
				IDString = IDString & "," & IDList(i)
			end if
	next 
	if blnallnumeric then
		if trim(IDString) <> "" then
			IDString = mid(IDString,2)
			IDString = " or i.id in (" & IDString & ") "
		end if
	else
		IDString = ""
	end if


  'Create a recordset
  rs.ActiveConnection = cn  
  if lookuptypeid = 0 then
	strSQl = "SELECT COUNT(1) AS RecordCount FROM DeliverableIssues AS i WITH (NOLOCK) LEFT OUTER JOIN DeliverableRoot AS r WITH (NOLOCK) ON i.DeliverableRootID = r.ID WHERE (i.description like '%" & ScrubSQL(sFind) & "%' " & ScrubSQL(IDString) & " or i.summary like '%" & ScrubSQL(sFind) & "%')"
	rs.Open strSQL, cn,adOpenForwardOnly
	if rs("RecordCount") > MaxMatchedDisplayed then		
		strSQl = "SELECT top " & MaxMatchedDisplayed & " i.ID,i.Summary, r.name as deliverable,ProductVersionID,Type, Status, TargetDate, ActualDate, i.Description FROM DeliverableIssues AS i WITH (NOLOCK) LEFT OUTER JOIN DeliverableRoot AS r WITH (NOLOCK) ON i.DeliverableRootID = r.ID WHERE (i.description like '%" & ScrubSQL(sFind) & "%' " & ScrubSQL(IDString) & " or i.summary like '%" & ScrubSQL(sFind) & "%') order by i.id desc"
	else
		strSQl = "SELECT i.ID,i.Summary, r.name as deliverable,ProductVersionID,Type, Status, TargetDate, ActualDate, i.Description FROM DeliverableIssues AS i WITH (NOLOCK) LEFT OUTER JOIN DeliverableRoot AS r WITH (NOLOCK) ON i.DeliverableRootID = r.ID WHERE (i.description like '%" & ScrubSQL(sFind) & "%' " & ScrubSQL(IDString) & " or i.summary like '%" & ScrubSQL(sFind) & "%')"

'        strSQL = strSQl & " union Select id, Summary, null as deliverable, projectid,0,statusid, null as TargetDate, DateClosed as ActualDate, Details from SupportIssues i with (NOLOCK) WHERE (i.details like '%" & ScrubSQL(sFind) & "%' " & ScrubSQL(IDString) & " or summary like '%" & ScrubSQL(sFind) & "%')"
	end if
	strRecordCount = rs("RecordCount") & ""
	rs.Close
	
	rs.Open strSQL
  else
	rs.Open "SELECT i.ID,i.Summary, r.name as deliverable,ProductVersionID,Type, Status, TargetDate, ActualDate, i.Description FROM DeliverableIssues AS i WITH (NOLOCK) LEFT OUTER JOIN DeliverableRoot AS r WITH (NOLOCK) ON i.DeliverableRootID = r.ID WHERE Type=" & ScrubSQL(lookuptypeid) & " and (i.description like '%" & ScrubSQL(sFind) & "%' " & ScrubSQL(IDString) & " or i.summary like '%" & ScrubSQL(sFind) & "%')"
  end if
  if rs.EOF and rs.EOF then
		Response.Write "<LABEL class=Progress id=Progress name=Progress>&nbsp;</LABEL>"
		if trim(lcase(sType)) = "allactions" then
			Response.Write "<h2>No Actions match search criteria</h2>"
		else
			Response.Write "<h2>No " & sType & " match search criteria</h2>"
		end if
		Response.Write "<font size=2 face=verdana><b>Fields searched:</b>  ID, Summary, Description"
		Response.write "<br /><b>Search Criteria:</b> " & ScrubSQL(sFind) &  "</font>"
		
  else
  %> 
  
<label class=Progress id=Progress>&nbsp;</label>	  
<font face=verdana size=3 color=black><b>Action Items matching search criteria</b></font>
<table id="rows" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2  LANGUAGE=javascript onmousemove="return rows_onmousemove()" onmouseout="return rows_onmouseout()" onclick="return rows_onclick()">
	
    <tr bgcolor=cornsilk>
		<td width=40><b><font face=verdana size=1>ID</font></b></td>
		<td><b><font face=verdana size=1>Program</font></b></td>
		<td><b><font face=verdana size=1>Type</font></b></td>
		<td><b><font  face=verdana size=1>Status</font></b></td>
		<td><b><font  face=verdana size=1>TargetDate</font></b></td>
		<td><b><font  face=verdana size=1>ActualDate</font></b></td>
		<td><b><font  face=verdana size=1>Summary</font></b></td>
	</tr>
<% 
	ColorIndex = 1
   do while not rs.eof	
		ItemCount = ItemCount + 1
%>
    <tr 
		<% if colorindex = 1 then 
		colorindex = 2 %>
		bgcolor="ivory"			
		<%else
			
			colorindex = 1
		%>
		bgcolor="ivory"
		<%end if
		
		strproduct = ""
		if rs("ProductversionID") > 0 then
			rs2.ActiveConnection = cn
			rs2.Open "Select f.Name + ' ' + v.version as Name from ProductFamily f with (NOLOCK),ProductVersion v with (NOLOCK) Where v.Productfamilyid = f.id and 	v.id = " & rs("ProductversionID"),cn,adOpenForwardOnly
			if rs2.EOF and rs2.BOF then
				strProduct = "No Product Specified"
			else
				strProduct = rs2("Name")
			end if
			rs2.Close
		end if
			
			
		strTargetdate = rs("TargetDate")
		if isnull(strtargetdate) then
			strtargetdate = "&nbsp;"	
		end if
		
		strActualDate = rs("ActualDate") & ""
		if stractualdate = "" then
			stractualdate = "&nbsp;"	
		else
			strActualDate = formatdatetime(strActualDate,vbshortdate)
		end if
		
		select case rs("Status")
            case 1
                strstatus = "Open"
            case 2
                strstatus = "Closed"
            case 3
                strstatus = "Need More Info"
            case 4
                strstatus = "Approved"
            case 5
                strstatus = "Disapproved"
            case 6
                strstatus = "Investigating"
		    case else
                strstatus = "N/A"
		end select			
		
		Select Case rs("type")
		    case 1
			    strType = "Issue"
		    case 2
			    strType = "Action Item"
		    case 3
			    strType = "Change Request"
		    case 4
			    strType = "Note"
		    case 5
			    strType = "Improvement"
		    case 6
			    strType = "Test Request"
	        case 7
	            strType = "Service ECR"
		end select
		
		if isnull(rs("Deliverable")) then
			strDeliverable = strProduct
		else
			strDeliverable = rs("Deliverable") 
		end if
		
		LookupTypeID = rs("Type")

		strSummary = replace(rs("Summary") & "&nbsp;",strFind,"<b class=""bold"">" & strfind & "</b>",1, -1,1)
		%>
		Class="ID=<%=rs("ID")%>&Type=<%=LookupTypeID%>" oncontextmenu="contextMenu(<%=rs("ID")%>,<%=LookupTypeID%>);return false;">
		
		<td class="cell"><%=replace(rs("ID"),strFind,"<b class=""bold"">" & strfind & "</b>")%></font></td>
		<td nowrap class="cell"><%=strDeliverable%></td>
		<td nowrap class="cell"><%=strType%></td>
		<td class="cell"><%=strStatus%></td>
		<td class="cell"><%=strTargetdate%></td>
		<td class="cell"><%=strActualDate%></td>
		<td class="cell"><%=strSummary%></td>
	</tr>
<%
	LastID = "ID=" & rs("ID")& "&Type=" & LookupTypeID
	rs.MoveNext
	loop
	rs.Close
%>
</table>
<%if isnumeric(strRecordcount) then
	if clng(strRecordcount)> MaxMatchedDisplayed then%>
<br /><br />
<font size=2 color=black face=verdana>Only the newest <%=MaxMatchedDisplayed%> records matching your criteria have been displayed.  There are <%=strRecordcount - MaxMatchedDisplayed%> other records matching your request. </font>
		
	<%end if
end if%>
<br />
<br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<h3></h3>
<h3></h3>
<%end if 
elseif sType = "Problem Reports" then
  'Create a recordset
%>
	<LABEL class=Progress id=Progress>Loading Problem Reports.  Please wait...</LABEL>	

<%
    rs.ActiveConnection = cn  
  rs.Open "Select component + ' ' + ComponentVersion as Deliverable, state as Status, ObservationID as ID, Owner,shortDescription as Summary, PrimaryProduct as product, Priority FROM HOUSIREPORT01.SIO.dbo.SI_Observation_Report o (NOLOCK) Where o.shortdescription like '%" & scrubsql(sFind) & "%' or o.observationid like '%" & scrubsql(sFind) & "%'",cn,adOpenForwardOnly
	if rs.EOF and rs.EOF then
		Response.Write "<h2>No Problem Reports match search criteria</h2>"
	else
%>
<font face=verdana size=3 color=black><b>Problem Reports matching search criteria</font>
<table id="otsrows" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2>
    <tr bgColor=cornsilk>
		<td><b><font face=verdana size=1>Number</font></b></td>
		<td><b><font face=verdana size=1>Product</font></b></td>
		<td><b><font  face=verdana size=1>Deliverable</font></b></td>
		<td><b><font  face=verdana size=1>Pr</font></b></td>
		<td><b><font  face=verdana size=1>Status</font></b></td>
		<td><b><font  face=verdana size=1>Owner</font></b></td>
		<td><b><font  face=verdana size=1>Summary</font></b></td>
	</tr>
<% 
	ColorIndex = 1
	rowcount=0
	do while not rs.eof	
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if
        rowcount = rowcount +1
%>
    <tr Class="flag://!!PRID:<%= rs("ID")%> ">
		<td class="cell"><%=rs("ID")%></td>
		<td class="cell"><%=rs("Product") & "&nbsp;"%></td>
		<td class="cell"><%=rs("Deliverable") & "&nbsp;"%></td>
		<td class="cell"><%=rs("Priority") & "&nbsp;"%></td>
		<td class="cell"><%=rs("Status") & "&nbsp;"%></td>
		<td class="cell"><%=rs("Owner") & "&nbsp;"%></td>
		<td class="cell"><%=rs("Summary") & "&nbsp;"%></td>
	</tr>
<%
	rs.MoveNext
	loop
	rs.Close
%>
</table>
<%
    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if
%>
<br /><br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<H3></H3>
<H3></H3>
<%
  end if
elseif sType = "Products" then
  'Create a recordset
%>
	<LABEL class=Progress id=Progress>Searching For Products.  Please wait...</LABEL>	
<%
    strFindScrubbed = scrubsql(sFind)

	IDList = split(strFind,",")
	  
	IDString = ""
	blnAllNumeric = true
	for i = 0 to Ubound(IDList) 
		if (not (isinteger(IDList(i)) or trim(IDList(i))="")) or (isnumeric(IDList(i)) and (instr(lcase(IDList(i)),"d") > 0 or instr(lcase(IDList(i)),"e") > 0)) then
			blnAllNumeric = false
		end if
		if  trim(IDList(i))<>"" then
			IDString = IDString & "," & IDList(i)
		end if
	next 
	
	if blnallnumeric and trim(IDString) <> "," then
		if trim(IDString) <> "" then
			IDString = mid(IDString,2)
			IDString = " or id in (" & IDString & ") "
		end if
	else
		IDString = ""
		RootIDString = ""
	end if

  rs.ActiveConnection = cn  
  '  rs.Open "Select v.Id, v.DotsName, v.SystemBoardID, dbo.concatenate(b.seriessummary) as SeriesSummary FROM productversion v, Product_Brand b Where v.id = b.productversionid and (dotsname like '%" & scrubsql(sFind) & "%' or SystemboardID like '%" & scrubsql(sFind) & "%' or b.seriessummary like '%" & scrubsql(sFind) & "%' " & IDString & ") group by v.id, v.dotsname, v.systemboardid order by dotsname",cn,adOpenForwardOnly
  
  strSQl = "SELECT * FROM (" & _
                "SELECT v.ID, v.DOTSName, p.Name AS Partner, v.SystemBoardID, dbo.Concatenate(b.SeriesSummary) AS SeriesSummary, dbo.Concatenate(COALESCE (bn.Name, '') + COALESCE (bn.Streetname, '') + COALESCE (bn.Suffix, '')) AS Brand, " & _ 
                        "CASE v.FusionRequirements WHEN 1 THEN 'Pulsar' ELSE 'Legacy' END AS Requirements " & _
								"FROM ProductVersion AS v WITH (NOLOCK) " & _
								"INNER JOIN Product_Brand AS b WITH (NOLOCK) ON v.ID = b.ProductVersionID " & _
								"INNER JOIN Brand AS bn WITH (NOLOCK) ON b.BrandID = bn.ID " & _
								"LEFT OUTER JOIN Partner AS p WITH (NOLOCK) ON v.PartnerID = p.ID " & _
			        "WHERE  (v.ID NOT IN (100, 188, 209, 230, 232, 252, 257)) " & _ 
			        "GROUP BY v.ID, v.DOTSName, p.Name, v.SystemBoardID, v.FusionRequirements " & _ 
			        "UNION " & _ 
			        "SELECT  v.ID, v.DOTSName, p.Name AS Partner, v.SystemBoardID, '' AS SeriesSummary, '' AS Brand, " & _ 
                             "CASE v.FusionRequirements WHEN 1 THEN 'Pulsar' ELSE 'Legacy' END AS Requirements " & _
								"FROM ProductVersion AS v WITH (NOLOCK) LEFT OUTER JOIN Partner AS p WITH (NOLOCK) ON v.PartnerID = p.ID " & _
								"WHERE (v.ID NOT IN (SELECT DISTINCT ProductVersionID FROM Product_Brand WITH (NOLOCK))) " & _
			        "AND (v.Division = 1)  " & _ 
                    "AND (v.ID NOT IN (170, 100, 250, 133, 255, 258, 323, 343, 344, 446))" & _ 
                 ") p " & _ 
                 "WHERE dotsname like '%" & scrubsql(sFind) & "%'  or SystemboardID like '%" & scrubsql(sFind) & "%' or Partner like '%" & scrubsql(sFind) & "%' or Brand like '%" & scrubsql(sFind) & "%' or seriessummary like '%" & scrubsql(sFind) & "%' " & IDString & " " & _
                 "ORDER BY dotsname"
    rs.Open strSQl,cn,adOpenStatic
    if rs.EOF and rs.EOF then
		Response.Write "<font family=verdana size=4><b>No Products match search criteria</font></b><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:</b> ID, Name, SysID, Series, Partner"
		Response.write "<br /><b>Search Criteria:</b> " & server.HTMLEncode(strFindScrubbed) &  "</font><br />"
		rs.Close
	else
%>
  
<font face=verdana size=3><b>Products matching search criteria</b></font><br />
<br /><font size=2 face=verdana><b>Fields searched:</b>  ID, Name, SysID, Series, Partner, Brand
<br /><b>Search Criteria:</b>&nbsp;<%=server.HTMLEncode(strFindScrubbed)%></font><br />
<table id="TABLE1" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2 LANGUAGE=javascript onmousemove="return rows_onmousemove()" onmouseout="return rows_onmouseout()">
    <tr bgcolor=cornsilk>
		<td><b>ID</b></td>
		<td><b>Product</b></td>
        <td><b>Requirements</b></td>
		<td><b>Partner</b></td>
		<td><b>SysID</b></td>
		<td><b>Series</b></td>
	</tr>
<% 
	ColorIndex = 1
    RowCount = 0
	do while not rs.eof	
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if
        RowCount = RowCount +1 
%>
	<tr LANGUAGE=javascript onclick="return ProductMenu(<%=rs("ID")%>)">
		<td class="cell"><%=rs("ID")%></td>
		<td class="cell"><%=rs("Dotsname") & ""%></td>
        <td class="cell"><%=rs("Requirements") & ""%></td>
		<td class="cell"><%=rs("Partner") & ""%></td>
		<td class="cell"><%=replace(replace(Ucase(rs("Systemboardid") & ""),"H",""),"0X","")%>&nbsp;</td>
		<td class="cell"><%=rs("SeriesSummary") & ""%>&nbsp;</td>
	</tr>
<%
	rs.MoveNext
	loop
	rs.Close
%>
</TABLE>
<%
    if RowCount > MaxMatchedDisplayed then	
           response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if
%>
<br /><br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<H3></H3>
<H3></H3>
<%
	end if
elseif sType = "OTS" then
  'Create a recordset
%>
	<LABEL class=Progress id=Progress>Loading Observation.  Please wait...<br /></LABEL>	
<%
    dim strOTS 
    dim OTSArray
    dim blnOTSDown
      'strOTS = replace(replace(sFind,"'",""),"""","")
      strOTS = scrubsql(sFind)
     
    '  OTSArray = Split(strOTS,",")
     ' strOTS = ""
     ' for i = lbound(OTSArray) to ubound(OTSArray)
    '	strOTS = strOTS & ",'" & right("0000000" & trim(OTSArray(i)),7) & "'"
     ' next
     ' if strOTS <> "" then
    '	strOTS = mid(strOTS,2)
    '  end if
      rs.ActiveConnection = cn  
      blnOTSDown = false

    dim strLimitPartner
    strLimitPartner = ""

    if CurrentUserPartnerName <> "HP" then
        strLimitPartner = " ("
        strLimitPartner = strLimitPartner & " o.ownergroup like '%" & currentuserpartnername & "%'"
        strLimitPartner = strLimitPartner & " or o.originatorgroup like '%" & currentuserpartnername & "%'"
        strLimitPartner = strLimitPartner & " or o.developergroup like '%" & currentuserpartnername & "%'"
        strLimitPartner = strLimitPartner & " or o.ComponentPMgroup like '%" & currentuserpartnername & "%' "


        if trim(CurrentUserOtherPartnerNames) <> "" then
            dim strpartnername
            OtherPartnerNameArray = split(CurrentUserOtherPartnerNames,",")
            
            for each strpartnername in OtherPartnerNameArray
                strpartnername = trim(strpartnername)
                if strpartnername <> "" then
                    strLimitPartner = strLimitPartner & " or o.ownergroup like '%" & strpartnername & "%'"
                    strLimitPartner = strLimitPartner & " or o.originatorgroup like '%" & strpartnername & "%'"
                    strLimitPartner = strLimitPartner & " or o.developergroup like '%" & strpartnername & "%'"
                    strLimitPartner = strLimitPartner & " or o.ComponentPMgroup like '%" & strpartnername & "%' "
                end if
            next
        end if

        strLimitPartner = strLimitPartner & " ) "
    end if
    on error resume next
  strSQl = "Select component + ' ' + ComponentVersion as Deliverable, state as Status, ObservationID as ID, Owner,shortDescription as Summary, PrimaryProduct as product, Priority FROM HOUSIREPORT01.SIO.dbo.SI_Observation_Report o (NOLOCK) Where o.DivisionID = 6 and (o.observationid like '%" & strOTS & "%' or o.shortdescription like '%" & strOTS & "%' or o.longdescription like '%" & strOTS & "%') and o.status <> 'Closed'"
    if CurrentUserPartnerName <> "HP" then
        strSQl = strSQL & " and ( " & strLimitPartner & ") " 
    end if

    rs.Open  strSQl,cn,adOpenForwardOnly
      
   if cn.Errors.Count > 0 then
       blnOTSDown = true
   end if
    on error goto 0
  
     if blnOTSDown then
        response.write "OTS is Currently Unavailable"
	elseif rs.EOF and rs.EOF then
		Response.Write "<font face=verdana size=3><b>No Open Mobile Observations match search criteria</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Observation ID, Short Description, Long Description"
		Response.write "<br /><b>Search Criteria:&nbsp;&nbsp;</b> " & replace(strOTS,"''","'") &  "</font>"
	else
%>
  
<font face=verdana size=3><b>Open Mobile Observations matching search criteria</b></font><br />
<font size=2 face=verdana><b><br />Fields searched:&nbsp;&nbsp;</b>Observation ID, Short Description, Long Description
<br /><b>Search Criteria:&nbsp;&nbsp;</b><%=replace(strOTS,"''","'")%></font><br /><br />
<table id="TABLE2" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2>
    <tr bgColor=cornsilk>
		<td><b>Number</b></td>
		<td><b>Product</b></td>
		<td><b>Deliverable</b></td>
		<td><b>Pr</b></td>
		<td><b>Status</b></td>
		<td><b>Owner</b></td>
		<td><b>Summary</b></td>
	</tr>

<%
    RowCount = 0
	do while not rs.EOF
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if
        RowCount = RowCount + 1	
		Response.Write "<TR><TD valign=top><a target=OTS href=""http://" & Application("Excalibur_ServerName") & "/search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ID") & """><font size=1 face=verdana>" & rs("ID") & "</a></td>"
		Response.Write "<TD valign=top>" & rs("product") & "</td>"
		Response.Write "<TD valign=top>" & rs("Deliverable") & "&nbsp;"  & "</td>"
		Response.Write "<TD valign=top>" & rs("Priority") & "</td>"
		Response.Write "<TD valign=top>" & rs("status")  & "</td>"
		Response.Write "<TD valign=top>" &  rs("owner") & "</td>"
		Response.Write "<TD valign=top>" & rs("summary") & "</td>"
		Response.Write "</TR>"
	
	
		rs.MoveNext
	loop
%>
</table>
<%
    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if
%>
<br /><br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<H3></H3>
<H3></H3>   
<%
    end if
    if not blnOTSDown then
        rs.Close
    end if    
 ElseIf sType = "Part" Then
 'Create a recordset
%>    
	<label class="Progress" id="Progress">Searching Material Data.  Please wait...<br /></label>	
<%
'response.Flush
'
' Search Material Master
'

    dim strParts 
    dim PartsArray
    strParts = scrubsql(sFind)
 
    dim strPartSql 
    strPartSql = "select gd.PartNumber, md.GPGDescription, gd.MaterialType, gd.Division, gd.RevisionLevel, gd.CrossPlantStatus, ian.ian_upc, ian.ian_category_ean, ian.ean_indicator " & _
        "from datawarehouse.dbo.ihub_materialmastergeneraldata gd  with (NOLOCK) " & _
        "inner join datawarehouse.dbo.ihub_materialdescription md  with (NOLOCK) on gd.partnumber = md.partnumber " & _
        "left outer join datawarehouse.dbo.iHUB_InternationalArticleNumbers ian  with (NOLOCK) on gd.partnumber = ian.partnumber " & _
        "where gd.partnumber like '%" & strParts & "%'"
     if currentuser="dwhorton" then
	    'Response.Write strPartSQl
	    'Response.Flush
    end if

    rs.ActiveConnection = cn  
    rs.Open  strPartSql ,cn, adOpenForwardOnly
	if rs.EOF and rs.EOF then
		Response.Write "<font face=verdana size=3><b>No Description Found in the Corporate Data</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Part Number"
		Response.write "<br /><b>Search Criteria:&nbsp;&nbsp;</b> " & replace(strParts,"''","'") &  "</font>"
	else

%>
  
<font face=verdana size=3><b>Material Description Found in the Corporate Data</b></font><br />
<font size=2 face=verdana><b><br />Fields searched:&nbsp;&nbsp;</b>Part Number
<br /><b>Search Criteria:&nbsp;&nbsp;</b><%=replace(strParts,"''","'")%></font><br /><br />
<table id="partsRows" bgColor=ivory borderColor=tan width="95%" border=1 cellspacing=1 cellpadding=2>
    <tr bgColor=cornsilk>
		<td><b>Number</b></td>
		<td><b>GPG Description</b></td>
		<td><b>Material Type</b></td>
		<td><b>Division</b></td>
		<td><b>Revision Level</b></td>
		<td><b>Cross Plant Status</b></td>
		<td><b>UPC/JAN</b></td>
	</tr>


<%
   RowCount = 0
	do while not rs.EOF
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if	
        RowCount = RowCount + 1
		Response.Write "<TD valign=top><a href=""find.asp?type=Part&find=" & Server.urlencode(Trim(rs("PartNumber"))) & """>" & rs("PartNumber") & "</a>&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("GPGDescription") & "&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("MaterialType") & "&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("Division")  & "&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("RevisionLevel") & "&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("CrossPlantStatus") & "&nbsp;</td>"
		Response.Write "<TD valign=top>" & rs("IAN_UPC") & "&nbsp;</td>"
		Response.Write "</TR>"	
	
	
		rs.MoveNext
	loop
%>
</table>
<% 
    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if

  end if
  rs.Close
%>
<br />
<hr />
<br />
<%
'response.Flush
'
' Search Bom as Parent
'
    strPartSql = "select distinct bom.partnumber, bom.childpartnumber, md.gpgdescription, mm.RevisionLevel, mm.CrossPlantStatus, bom.qtyper, bom.sortstring " & _
        "from datawarehouse.dbo.iHUB_BillOfMaterial bom with (NOLOCK) " & _
        "inner join datawarehouse.dbo.ihub_materialdescription md with (NOLOCK) on bom.childpartnumber = md.partnumber " & _
        "inner join datawarehouse.dbo.iHUB_MaterialMasterGeneralData mm with (NOLOCK) on bom.childpartnumber = mm.partnumber " & _
        "where bom.partnumber like '%" & strParts & "%' " & _
        "and DATEDIFF(d, GETDATE(), bom.enddt) > 0 " & _
        "Order By bom.PartNumber, bom.SortString"

 
    rs.ActiveConnection = cn  
    rs.ActiveConnection.CommandTimeout = 120
    rs.Open  strPartSql ,cn, adOpenForwardOnly
	if rs.EOF and rs.EOF then
		Response.Write "<font face=verdana size=3><b>No BOM Structure Found in the Corporate Data</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Part Number"
		Response.write "<br /><b>Search Criteria:&nbsp;&nbsp;</b> " & replace(strParts,"''","'") &  "</font>"
	else

%>
  
<font face=verdana size=3><b>BOM Structure Found in the Corporate Data</b></font><br />
<font size=2 face=verdana><b><br />Fields searched:&nbsp;&nbsp;</b>Part Number
<br /><b>Search Criteria:&nbsp;&nbsp;</b><%=replace(strParts,"''","'")%></font><br /><br />
<table id="Table3" bgColor=ivory borderColor=tan width="95%" border=1 cellspacing=1 cellpadding=2>
    <tr bgColor=cornsilk>
		<td><b>Number</b></td>
		<td><b>GPG Description</b></td>
		<td><b>Rev</b></td>
		<td><b>Cross Plant Status</b></td>
		<td><b>Qty</b></td>
		<td><b>Sort String</b></td>
	</tr>



<%
    Dim strLastPart
    strLastPart = ""
    RowCount = 0
	do while not rs.EOF
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if
        RowCount = RowCount + 1
	    If strLastPart <> rs("PartNumber") Then
	        Response.Write "<tr><td valign=top colspan=7 style=""background-color:LightSteelBlue;font-family:verdana;font-size:xx-small;font-weight:bold;"">" & rs("PartNumber") & "</td></tr>"
	        strLastPart = rs("PartNumber")
	    End If
	
		Response.Write "<tr><td valign=top><font size=1 face=verdana><a href=""find.asp?type=Part&find=" & Server.urlencode(Trim(rs("ChildPartNumber"))) & """>" & rs("ChildPartNumber") & "</a>&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("GPGDescription") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("RevisionLevel") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("CrossPlantStatus") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("QtyPer") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("SortString") & "&nbsp;</td>"
		Response.Write "</tr>"
	
	
		rs.MoveNext
	loop
%>
</TABLE>
<% 

        if RowCount > MaxMatchedDisplayed then	
            response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
        end if


    end if 
	rs.Close
%>
<br />
<hr />
<br />
<%
'response.Flush
'
' Search Bom as Child
'
    strPartSql = "select distinct bom.partnumber, bom.childpartnumber, md.gpgdescription, bom.startdt, bom.enddt, bom.qtyper, bom.usage, bom.sortstring, bom.bomitemnumber " & _
        "from datawarehouse.dbo.iHUB_BillOfMaterial bom with (NOLOCK) " & _
        "inner join datawarehouse.dbo.ihub_materialdescription md with (NOLOCK) on bom.partnumber = md.partnumber " & _
        "where bom.childpartnumber like '%" & strParts & "%' " & _
        "Order By bom.childpartnumber, bom.PartNumber, bom.Usage, bom.BomItemNumber"

 
    rs.ActiveConnection = cn  
    rs.Open  strPartSql ,cn, adOpenForwardOnly
	if rs.EOF and rs.EOF then
		Response.Write "<font face=verdana size=3><b>No Where Used Information Found in the Corporate Data</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Child Part Number"
		Response.write "<br /><b>Search Criteria:&nbsp;&nbsp;</b> " & replace(strParts,"''","'") &  "</font>"
	else

%>
  
<font face=verdana size=3><b>Where Used Information Found in the Corporate Data</b></font><br />
<font size=2 face=verdana><b><br />Fields searched:&nbsp;&nbsp;</b>Part Number
<br /><b>Search Criteria:&nbsp;&nbsp;</b><%=replace(strParts,"''","'")%></font><br /><br />
<table id="Table4" bgColor=ivory borderColor=tan width="95%" border=1 cellspacing=1 cellpadding=2>
    <tr bgColor=cornsilk>
		<td><b>Number</b></td>
		<td><b>GPG Description</b></td>
		<td><b>Start Dt Type</b></td>
		<td><b>End Dt</b></td>
		<td><b>Qty</b></td>
		<td><b>Usage</b></td>
		<td><b>Sort String</b></td>
	</tr>



<%
    strLastPart = ""
    RowCount=0
	do while not rs.EOF
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if
        RowCount=RowCount + 1
	    If strLastPart <> rs("ChildPartNumber") Then
	        Response.Write "<tr><td valign=top colspan=7 style=""background-color:LightSteelBlue;font-family:verdana;font-size:xx-small;font-weight:bold;"">" & rs("ChildPartNumber") & "</td></tr>"
	        strLastPart = rs("ChildPartNumber")
	    End If

		Response.Write "<tr><td valign=top><a href=""find.asp?type=Part&find=" & Server.urlencode(Trim(rs("PartNumber"))) & """>" & rs("PartNumber") & "</a>&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("GPGDescription") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("StartDt") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("EndDt")  & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("QtyPer") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("Usage") & "&nbsp;</td>"
		Response.Write "<td valign=top>" & rs("SortString") & "&nbsp;</td>"
		Response.Write "</tr>"
	
	
		rs.MoveNext
	loop

%>
</TABLE>
<% 

        if RowCount > MaxMatchedDisplayed then	
            response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
        end if


    end if 
	rs.Close%>

<!-----  NEW SECTION RON 1/10/2012 ------->
<% 
  
sqlSparekits = " SELECT DISTINCT SKBOM.CategoryName AS [Spare Kit Category], " & _ 
" SKBOM.SpareKitNo AS [Spare Kit Part No], SKBOM.Description AS " & _ 
" [Spare Kit Desc]  FROM (SELECT DISTINCT ProductVersionID, DevCenter, " & _ 
" PVName, PVPB.ServiceFamilyPn AS ServiceFamilyPn, ProductBrandID, BrandName, " & _ 
" PVPB.KMAT AS KMAT, PVPartnerID, SKU, ConfigCode, GPGDescription, AVs, " & _ 
" SKUS.ServiceGeoID AS SKUServiceGeoID, SKUS.ServiceGeoShortName AS " & _ 
" SKUServiceGeo, Released, Rev, CrossPlantStatus, MaterialType, " & _ 
" IncludeInReports, SFMultPartnerID, OSSPS.SFMultPartnerName, " & _ 
" OSSPS.SFMultPartnerGeoID, OSSPS.SFMultPartnerGeo FROM " & _ 
" ((SELECT DISTINCT PV.ID AS ProductVersionID, PV.DevCenter, " & _ 
" PV.DOTSName AS PVName, PV.ServiceFamilyPn, PB.ID AS ProductBrandID, " & _ 
" dbo.Brand.Name AS BrandName, PB.KMAT, PV.PartnerID AS PVPartnerID " & _ 
" FROM dbo.ServiceFamilyDetails WITH (NOLOCK) RIGHT OUTER JOIN " & _ 
" dbo.ProductVersion AS PV WITH (NOLOCK) ON dbo.ServiceFamilyDetails.ServiceFamilyPn = " & _ 
" PV.ServiceFamilyPn LEFT OUTER JOIN dbo.Product_Brand AS PB WITH (NOLOCK) " & _ 
" ON PV.ID = PB.ProductVersionID LEFT OUTER JOIN dbo.Brand WITH (NOLOCK) " & _ 
" ON PB.BrandID = dbo.Brand.ID WHERE (PV.ProductStatusID <> 5) " & _ 
" AND (dbo.ServiceFamilyDetails.AutoPublishRsl = 1)) PVPB LEFT JOIN " & _ 
" (SELECT DISTINCT SKU, ConfigCode, GPGDescription, AVs, A.ServiceGeoID, " & _ 
" Released, Rev, CrossPlantStatus, MaterialType, IncludeInReports, KMAT, " & _ 
" ServiceGeoShortName FROM dbo.BTOSSSKUAVs A WITH (NOLOCK) LEFT JOIN " & _ 
" dbo.ServiceGeo B WITH (NOLOCK) ON A.ServiceGeoID=B.ServiceGeoID) SKUS " & _
" ON PVPB.KMAT=SKUS.KMAT LEFT JOIN (SELECT DISTINCT dbo.ServiceFamily_Partner.ServiceFamilyPn, " & _ 
" dbo.ServiceFamily_Partner.PartnerID AS SFMultPartnerID, dbo.Partner.Name " & _ 
" AS SFMultPartnerName, dbo.ServiceFamily_Partner.ServiceGeoID AS " & _ 
" SFMultPartnerGeoID, dbo.ServiceGeo.ServiceGeoShortName AS " & _ 
" SFMultPartnerGeo FROM dbo.ServiceGeo WITH (NOLOCK) RIGHT OUTER JOIN " & _ 
" dbo.ServiceFamily_Partner WITH (NOLOCK) ON dbo.ServiceGeo.ServiceGeoID = " & _ 
" dbo.ServiceFamily_Partner.ServiceGeoID RIGHT OUTER JOIN dbo.Partner WITH " & _ 
" (NOLOCK) ON dbo.ServiceFamily_Partner.PartnerID = dbo.Partner.ID " & _ 
" WHERE ServiceFamily_Partner.Status='A') OSSPS ON PVPB.ServiceFamilyPn=OSSPS.ServiceFamilyPn " & _ 
" AND SKUS.ServiceGeoID=OSSPS.SFMultPartnerGeoID)) PVPBSKUS LEFT JOIN " & _ 
" (SELECT ProductVersionID,ProductBrandID,KMAT,ServiceGeoID,ConfigCode,SKU,AVs,SpareKitNo,ServiceFamilyPn,Active, " & _ 
" LastUpdated,ServiceSpareKitMapId,GeoNa,GeoLa,GeoApj,GeoEmea,SpareKitId,Revision,LastChangeType,DATE_STR,DATE_INT " & _ 
" FROM dbo.BTOServiceSpares WITH(NOLOCK)) BTOSS ON (PVPBSKUS.SKU=BTOSS.SKU " & _ 
" AND PVPBSKUS.SKUServiceGeoID=BTOSS.ServiceGeoID) LEFT JOIN " & _ 
" (SELECT SpareKitID,dbo.BTOSSSKBOM.SpareKitNo,SA,Component,RegionalComponents, " & _ 
" LastChangeType,LastChangeTime,SADesc,ComponentDesc,ISNULL(dbo.ServiceSpareKit.Description,SKDesc) As " & _ 
" [Description],ChangeSource,Released,Rev,SKCrossPlantStatus,SKMaterialType, " & _ 
" SKIncludeInReports,SACrossPlantStatus,SAMaterialType,SAIncludeInReports, " & _ 
" CompCrossPlantStatus,CompMaterialType,CompIncludeInReports,SARev,CompRev, " & _ 
" SAReleased,CompReleased, CategoryName, SpareCategoryID AS " & _ 
" ServiceSpareCategoryID, " & _ 
" CompQtyPer, CompPriAltGen, SAQtyPer, SAPriAltGen  FROM dbo.BTOSSSKBOM WITH(NOLOCK) LEFT JOIN dbo.ServiceSpareKit WITH(NOLOCK) ON dbo.BTOSSSKBOM.SpareKitID=dbo.ServiceSpareKit.ID LEFT JOIN ServiceSpareCategory WITH(NOLOCK) ON ServiceSpareKit.SpareCategoryID=ServiceSpareCategory.ID WHERE SKIncludeInReports=1) SKBOM ON (BTOSS.SpareKitID=SKBOM.SpareKitID) WHERE " & _ 
" (SKBOM.Component IN('" & strParts &  "')) AND (CompIncludeInReports=1) ORDER BY SKBOM.CategoryName ASC "

   
    rs.ActiveConnection = cn  
    rs.Open  sqlSparekits ,cn, adOpenForwardOnly
    if rs.eof then
%>
<br /><br />
<hr /><br />
<font face=verdana size=3><b>No Where Used in Sparekits Found in the Corporate Data</b></font><br /><br />
<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Child Part Number
<br /><b>Search Criteria:&nbsp;&nbsp;</b> <%=replace(strParts,"''","'") %></font><br />
<%
 else%>
 <br /><br />
<hr /><br />
<font face=verdana size=3><b>Where Used in Sparekits Found in the Corporate Data</b></font><br /><br />
<font size=2 face=verdana><b>Fields searched:&nbsp;&nbsp;</b>Part Number
<br /><b>Search Criteria:&nbsp;&nbsp;</b> <%=replace(strParts,"''","'") %></font>  
          <div><br />
	<table id="Table12" bgColor=ivory borderColor=tan width="95%" border=1 cellspacing=1 cellpadding=2>
      <tr  align="left" style="color:Black;background-color:cornsilk;">
			<th >Spare Kit Part No </th><th >Spare Kit Desc</th><th >Spare Kit Category</th>
	  </tr>
 <%while not rs.eof%>

 <tr class="SKURow">
			    <td>
                <a href="find.asp?type=Part&find=<%=rs("Spare Kit Part No")%>">
                <%=rs("Spare Kit Part No")%></a></td><td><%=rs("Spare Kit Desc")%></td><td><%=rs("Spare Kit Category")%></td>
		    </tr>
 <%rs.movenext
 wend 
 end if%>
 </div>
   </table>

 <%rs.Close%> 


<%
'
' Search Pulsar Deliverables
'
 

	rs.Open "SELECT r.ID as RootID, v.id,v.DeliverableName, v.irspartnumber, v.imagepath, v.location,v.version, v.revision, v.pass, v.active, v.partnumber, r.description from deliverableRoot r with (NOLOCK), deliverableVersion v with (NOLOCK) Where  v.deliverableRootID=r.ID and (v.partnumber like '%" & strParts & "%' or v.irspartnumber like '%" & strParts & "%' or v.cdpartnumber like '%" & strParts & "%' )"

	ColorIndex = 1
	EOLRowCount=0
	HeaderDrawn = false
    RowCount = 0
	do while not rs.eof	
        if RowCount > MaxMatchedDisplayed then	
            exit do
        end if    	
	
		if (not rs("active")) and request("ShowEOL") <> "1" then
			EOLRowCount=EOLRowCount + 1
		else
			if not headerdrawn then
				HeaderDrawn=true
            end if 'not headerdrawn
	    end if 'not active
    rs.MoveNext
	loop
	rs.Close
	Response.Write "</Table>"
	
	if EOLRowCount= 1 then
            if RowCount > MaxMatchedDisplayed then	
			    Response.Write "<br /><font color=red size=1>There is at least " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
            else
			    Response.Write "<br /><font color=red size=1>There is " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
		    end if
		elseif EOLRowCount > 1 then
            if RowCount > MaxMatchedDisplayed then	
			    Response.Write "<br /><font color=red size=1>There are at least " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
            else
			    Response.Write "<br /><font color=red size=1>There are " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
		    end if
		end if

        if RowCount > MaxMatchedDisplayed then	
            response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
        end if
	
%>
  



	
<br /><br /><p style="font-family:verdana; font-size:x-small; color:Red; font-weight:bold;">HP Restricted</p>
<%
else 'sType = Other
	Response.write "<LABEL class=Progress id=Progress name=Progress></LABEL>"
	dim RootIDString 
	
	IDList = split(strFind,",")
	  
	IDString = ""
	blnAllNumeric = true
	for i = 0 to Ubound(IDList) 
		if (not isinteger(IDList(i))) and (not ispartnumber(IDList(i))) and trim(IDList(i))<>"" then
			blnAllNumeric = false
		elseif trim(IDList(i))<>"" then
			if ispartnumber(IDList(i)) then
				IDString = IDString & "," & LoopupPartNumberID(IDList(i))
			else 
				IDString = IDString & "," & IDList(i)
			end if
		end if
	next 
	
	if blnallnumeric and trim(IDString) <> "," then
		if trim(IDString) <> "" then
			IDString = mid(IDString,2)
			RootIDString = " or r.id in (" & IDString & ") "
			IDString = " or v.id in (" & IDString & ") "
		end if
	else
		IDString = ""
		RootIDString = ""
	end if
 'Create a recordset
  rs.ActiveConnection = cn  
  strFindScrubbed = scrubsql(sFind)
 

  
  rs.Open "SELECT r.ID as RootID, v.id,v.DeliverableName, v.irspartnumber, v.imagepath, v.location,v.version, v.revision, v.pass, v.active, v.modelnumber, v.partnumber, r.description from deliverableroot r with (NOLOCK), deliverableversion v with (NOLOCK) Where  v.deliverablerootid=r.id and (v.partnumber like '%" & strFindScrubbed & "%' or v.irspartnumber like '%" & strFindScrubbed & "%' or v.modelnumber like '%" & strFindScrubbed & "%' or v.pnpdevices like '%" & strFindScrubbed & "%' or v.cdpartnumber like '%" & strFindScrubbed & "%' or r.description like  '%" & strFindScrubbed & "%' or v.deliverablename like  '%" & strFindScrubbed & "%' " & IDString & ") order by v.deliverablename, v.id" '"v.version, v.revision, v.pass"

	
	if rs.EOF and rs.EOF then
		Response.Write "<font family=verdana size=4><b>No Deliverable Versions match search criteria</font></b><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:</b>  ID, Deliverable Name, Part Number, IRS Part Number, Model Number, CD Part Number, PNPID/HWID, and Description"
		Response.write "<br /><b>Search Criteria:</b> " & server.HTMLEncode(strFindScrubbed) &  "</font><br />"
		rs.Close
	else
%>
		

<font face=verdana size=3><b>Deliverable Versions matching search criteria</b></font><br />
<br /><font size=2 face=verdana><b>Fields searched:</b>  ID, Name, Part Number, IRS Part Number, Model Number, CD Part Number, Softpaq Number, PNPID/HWID, and Description
<br /><b>Search Criteria:</b>&nbsp;<%=server.HTMLEncode(strFindScrubbed)%></font><br />
<table id="TABLE8" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2 LANGUAGE=javascript onmousemove="return rows_onmousemove()" onmouseout="return rows_onmouseout()" >
<% 
	ColorIndex = 1
	EOLRowCount=0
	HeaderDrawn = false
	RowCount=0
	do while not rs.eof	
	    if RowCount > MaxMatchedDisplayed then
	        exit do
	    end if
	    RowCount = RowCount + 1
		if (not rs("active")) and request("ShowEOL") <> "1" then
			EOLRowCount=EOLRowCount + 1
		else
			if not headerdrawn then
				HeaderDrawn=true
%>
    <tr bgcolor=cornsilk>
		<td><b>ID</b></td>
		<td><b>Name</b></td>
		<td><b>Version</b></td>
		<td><b>Revision</b></td>
		<td><b>Pass</b></td>
		<td><b>IRS&nbsp;Part&nbsp;Number</b></td>
		<td><b>Workflow</b></td>
		<TD style=display:none>Path</td>
	</tr>
			<%end if%>
    <tr LANGUAGE=javascript onclick="return DeliverableAlertMenu(<%=rs("RootID")%>,<%=rs("ID")%>,2)"> 
		<TD class="cell"><%=rs("ID")%></td>
		<TD class="cell"><%=rs("DeliverableName")%></td>
		<TD class="cell"><%=rs("Version")%></td>
		<TD class="cell"><%=rs("Revision")%>&nbsp;</td>
		<TD class="cell"><%=rs("Pass")%>&nbsp;</td>
		<TD class="cell"><%=rs("IRSPartNumber")%>&nbsp;</td>
		<TD class="cell"><%=rs("location")%></td>
		<TD  style=display:none class="cell" ID="Path2_<%=trim(rs("ID"))%>"><%=rs("ImagePath")%>&nbsp;</td>
	</tr>
<%
		end if

	rs.MoveNext
	loop
	rs.Close
	Response.Write "</Table>"
	
	if EOLRowCount= 1 then
        if RowCount > MaxMatchedDisplayed then	
    		Response.Write "<br /><font color=red size=1>There is at least " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    else
    		Response.Write "<br /><font color=red size=1>There is " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	elseif EOLRowCount > 1 then
        if RowCount > MaxMatchedDisplayed then	
    		Response.Write "<br /><font color=red size=1>There are at least " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    else
    		Response.Write "<br /><font color=red size=1>There are " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	end if

    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if
end if

'-----------------------------------------
Response.Write "<br /><HR><br />"
	rs.Open "SELECT r.ID,r.name, r.description, r.active from deliverableroot r with (NOLOCK) Where (r.description like  '%" & strFindScrubbed & "%' or r.name like  '%" & strFindScrubbed & "%' " & RootIDString & ")"
	if rs.EOF and rs.EOF then
		Response.Write "<font family=verdana size=4><b>No Root Deliverables match search criteria</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:</b>  ID, Name, and Description"
		Response.write "<br /><b>Search Criteria:</b> " & server.HTMLEncode(strFindScrubbed) &  "</font><br />"
		rs.Close
	else
%>
<font face=verdana size=3><b>Root Deliverables matching search criteria</b></font><br />
<br /><font size=2 face=verdana><b>Fields searched:</b>  ID, Name, and Description
<br /><b>Search Criteria:</b>&nbsp;<%=server.HTMLEncode(strFindScrubbed)%></font><br />
<TABLE id="TABLE10" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2 LANGUAGE=javascript onmousemove="return rows_onmousemove()" onmouseout="return rows_onmouseout()">
<% 
	ColorIndex = 1
	EOLRowCount=0
	HeaderDrawn = false
	RowCount = 0
	do while not rs.eof	
	    if RowCount > MaxMatchedDisplayed then
	        exit do
	    end if
	    RowCount = RowCount + 1
		if (rs("active")=0) and request("ShowEOL") <> "1" then
			EOLRowCount=EOLRowCount + 1
		else
			if not headerdrawn then
				HeaderDrawn=true
			
%>
    <tr bgcolor=cornsilk>
		<td><b>ID</b></td>
		<td><b>Name</b></td>
		<td><b>Description</b></td>
	</tr>
			
			<%end if%>
			
    <tr LANGUAGE=javascript onclick="return RootAlertMenu(<%=rs("ID")%>,1)"> 
		<TD class="cell"><%=rs("ID")%></td>
		<TD class="cell"><%=rs("Name")%></td>
		<TD class="cell"><%=rs("Description")%></td>
	</tr>
<%
		end if

	rs.MoveNext
	loop
	rs.Close
	Response.Write "</Table>"
	
	if EOLRowCount= 1 then
        if RowCount > MaxMatchedDisplayed then
	    	Response.Write "<br /><font color=red size=1>There is at least " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
    	else
    		Response.Write "<br /><font color=red size=1>There is " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	elseif EOLRowCount > 1 then
        if RowCount > MaxMatchedDisplayed then
		    Response.Write "<br /><font color=red size=1>There are at least " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
		else
		    Response.Write "<br /><font color=red size=1>There are " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	end if
    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if
	
	
	end if
'-----------------------------------------

Response.Write "<br /><HR><br />"

  rs.Open "SELECT r.ID,r.name, r.active, s.base as baseSubassembly, v.dotsname as Product, s.subassembly from deliverableroot r with (NOLOCK), product_delroot s with (NOLOCK), productversion v with (NOLOCK) Where v.id = s.productversionid and r.id = s.deliverablerootid and s.subassembly is not null and (s.base like  '%" & strFindScrubbed & "%' or s.subassembly like  '%" & strFindScrubbed & "%' or r.name like  '%" & strFindScrubbed & "%') "
	if rs.EOF and rs.EOF then
		Response.Write "<font family=verdana size=4><b>No Subassemblies match search criteria</b></font><br /><br />"
		Response.Write "<font size=2 face=verdana><b>Fields searched:</b>  Base Subassembly Number, Subassembly Number, Subassembly Name"
		Response.write "<br /><b>Search Criteria:</b> " & server.HTMLEncode(strFindScrubbed) &  "</font><br /><br />"
	else
%>

<font face=verdana size=3><b>Subassemblies matching search criteria</b></font><br />
<br /><font size=2 face=verdana><b>Fields searched:</b>  Base Subassembly Number, Subassembly Number, Subassembly Name
<br /><b>Search Criteria:&nbsp;</b><%=server.HTMLEncode(strFindScrubbed) %></font><br />
<table id="TABLE11" bgColor=ivory borderColor=tan WIDTH="95%" BORDER=1 CELLSPACING=1 CELLPADDING=2 LANGUAGE=javascript onmousemove="return rows_onmousemove()" onmouseout="return rows_onmouseout()" >
<% 
	ColorIndex = 1
	EOLRowCount=0
	HeaderDrawn = false
	RowCount=0
	do while not rs.eof	
	    if RowCount > MaxMatchedDisplayed then
	        exit do
	    end if
	    RowCount = RowCount + 1
		if (rs("active")=0) and request("ShowEOL") <> "1" then
			EOLRowCount=EOLRowCount + 1
		else
			if not headerdrawn then
				HeaderDrawn=true
%>
    <tr bgcolor=cornsilk>
		<td><b>ID</b></td>
		<td><b>Name</b></td>
		<td><b>Product</b></td>
		<td><b>Base</b></td>
		<td><b>Subassembly</b></td>
	</TR>
        <%end if%>
			


	<TR LANGUAGE=javascript onclick="return RootAlertMenu(<%=rs("ID")%>,2)"> 
		<TD class="cell"><%=rs("ID")%></td>
		<TD class="cell"><%=rs("Name")%></td>
		<TD class="cell"><%=rs("Product")%></td>
		<TD class="cell"><%=rs("BaseSubassembly")%>&nbsp;</td>
		<TD class="cell"><%=rs("Subassembly")%>&nbsp;</td>
	</TR>
<%
    end if
	rs.MoveNext
	loop
	rs.Close
%>
</table>
<%
	if EOLRowCount= 1 then
        if RowCount > MaxMatchedDisplayed then	
	    	Response.Write "<br /><font color=red size=1>There is at least " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
		else
		    Response.Write "<br /><font color=red size=1>There is " & EolRowCount & " inactive deliverable.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	elseif EOLRowCount > 1 then
        if RowCount > MaxMatchedDisplayed then	
	    	Response.Write "<br /><font color=red size=1>There are at least " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
		else    
		    Response.Write "<br /><font color=red size=1>There are " & EolRowCount & " inactive deliverables.</font> <font size=1><a href=""Find.asp?" &  Request.QueryString & "&ShowEOL=1" & """>Show Inactive Deliverables</a></font>"
	    end if
	end if
    if RowCount > MaxMatchedDisplayed then	
        response.Write "<BR><BR><font color=red><b>Note: Only the first " & MaxMatchedDisplayed & " matches are shown in this section</b></font><BR><BR>"
    end if

%>
</table></TABLE>
<br />
<br />
<font face="verdana" Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
<H3></H3>
<H3></H3>
<%	end if %>

<% end if %>


<%	if ItemCount <> 1 then
		LastID = 0
	end if
%>
 
<input type="hidden" id="txtOnlyID" name="txtOnlyID" value="<%=LastID%>" />
<div id="mnuPopup" style="display:none;position:absolute;width:2px;height:2px;left:0px;top:0px;padding:0px;background:white;border:1px solid gainsboro;z-index:100"></div>

</body>
</html>

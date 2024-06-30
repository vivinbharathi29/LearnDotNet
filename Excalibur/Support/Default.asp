<%@  language="VBScript" %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<HEAD>
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="includes/client/json2.js"></script>
    <script type="text/javascript" src="includes/client/json_parse.js"></script>

    <SCRIPT id="clientEventHandlersJS" language="javascript">
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->
    var CurrentState;
    var ItemDisplayed;

    function row_onmouseover() {
        var MyElement=window.event.srcElement;
        if (MyElement.tagName == "TD")
            MyElement = window.event.srcElement.parentElement;

	        MyElement.style.color = "red";
	        MyElement.style.cursor = "hand";
    }

    function row_onmouseout() {
        var MyElement=window.event.srcElement;
        if (MyElement.tagName == "TD")
            MyElement = window.event.srcElement.parentElement;

	        MyElement.style.color = "black";
	        MyElement.style.cursor = "default";
    }

    function window_onload(Tab){
	var strPath;
    if (Tab!="")
        SelectTab(Tab);
	//strPath = window.showModalDialog("support.asp","","dialogWidth:600px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 

        //Instantiate modalDialog load - PBI 27768 (Showmodaldialog to jquery dialog)
        modalDialog.load();
    }
	

function ProcessState() {
	var steptext;

	switch (CurrentState)
	{
	    case "Tickets":
	        tabTickets.style.display = "";
	        tabArticles.style.display = "none";
	        tabCategories.style.display = "none";
	        //txtQueryString.value = "Tab=Tickets";
	        window.scrollTo(0, 0);
	        break;

		case "Articles":
			tabTickets.style.display="none";
			tabArticles.style.display="";
			tabCategories.style.display="none";
			//txtQueryString.value = "Tab=Articles";
			window.scrollTo(0, 0);		
		break;

        case "Categories":
			tabTickets.style.display="none";
			tabArticles.style.display="none";
			tabCategories.style.display="";
			//txtQueryString.value = "Tab=Categories";
			window.scrollTo(0, 0);		
        break;
	}

}

function SelectTab(strStep) {
    var i;

    
	//Reset all tabs
	document.all("CellTicketsb").style.display="none";
	document.all("CellTickets").style.display="";
	document.all("CellArticlesb").style.display="none";
	document.all("CellArticles").style.display="";
	document.all("CellCategoriesb").style.display="none";
	document.all("CellCategories").style.display="";

	//Highight the selected tab
	document.all("Cell"+strStep).style.display="none";
	document.all("Cell"+strStep+"b").style.display="";

	CurrentState = strStep;
	
	ProcessState();

}


    function window_onmouseup() {
        if(typeof(mnuPopup)!="undefined")
            mnuPopup.style.display = "none";
        }


    function ShowMenu(ID){
	    var NewHeight;
	    var NewWidth;
        if (typeof(mnuPopup) != "undefined")
            {
            mnuPopup.style.display = "";
            mnuPopup.style.width = mnuPopup.scrollWidth;
            mnuPopup.style.height = mnuPopup.scrollHeight;
            mnuPopup.style.left = document.body.scrollLeft + ((event.clientX - event.offsetX) - 3);
            mnuPopup.style.top = (document.body.scrollTop-1) + (event.clientY - event.offsetY) + event.srcElement.offsetHeight;

            ItemDisplayed = ID;
            }
    }


    function EmailSubmitter(){
        alert(ItemDisplayed);
       // window.open("mailto:" & strTo & "?subject=Excalibur Ticket");
    }


    function ShowTicket(ID){
        var url = "Ticket.asp?ID=" + ID;
        modalDialog.open({dialogTitle:'Support Ticket', dialogURL:''+url+'', dialogHeight:650, dialogWidth:700, dialogResizable:true, dialogDraggable:true});
    }

    function ShowTicket_return(strResult){
	if (typeof(strResult) != "undefined")
		{
		    //window.location.reload();
		    window.location.href = "default.asp?" + txtQueryString.value;
		}
    }

    function AddTicket(ID){
        var url = "Support.asp";
        modalDialog.open({dialogTitle:'Mobile Tools Support', dialogURL:''+url+'', dialogHeight:600, dialogWidth:700, dialogResizable:true, dialogDraggable:true});
    }

    function AddTicket_return(strResult){
	if (typeof(strResult) != "undefined")
		{
                //window.location.reload();
  	      window.location.href = "default.asp?" + txtQueryString.value;
		}
    }

    function ShowArticle(ID){
        var url = "Article.asp?ID=" + ID;
        modalDialog.open({dialogTitle:'Support Article', dialogURL:''+url+'', dialogHeight:650, dialogWidth:850, dialogResizable:true, dialogDraggable:true});
    }

    function ShowArticle_return(strResult){
	if (typeof(strResult) != "undefined")
		{
            //window.location.reload();
		    window.location.href = "default.asp?" + txtQueryString.value;
		}
    }

    function AddArticle(ID){
        var url = "Article.asp";
        modalDialog.open({dialogTitle:'Support Article', dialogURL:''+url+'', dialogHeight:650, dialogWidth:850, dialogResizable:true, dialogDraggable:true});
    }

    function AddArticle_return(strResult){
	if (typeof(strResult) != "undefined")
		{
		    //window.location.href = "default.asp?Tab=Articles";
		    window.location.href = "default.asp?" + txtQueryString.value;
		}
    }


    function AddCategory(){
        var url = "Category.asp";
        modalDialog.open({dialogTitle:'Support Category', dialogURL:''+url+'', dialogHeight:700, dialogWidth:750, dialogResizable:true, dialogDraggable:true});
    }

    function AddCategory_return(strResult)
    {
	if (typeof(strResult) != "undefined")
		{
		    //window.location.href = "default.asp?Tab=Categories";
		    window.location.href = "default.asp?" + txtQueryString.value;

		}
    }

    function ShowCategory(ID){
        var url = "Category.asp?ID=" + ID;
        modalDialog.open({dialogTitle:'Support Category', dialogURL:''+url+'', dialogHeight:700, dialogWidth:750, dialogResizable:true, dialogDraggable:true});
    }

    function ShowCategory_return(strResult){
        if (typeof(strResult) != "undefined")
        {
		    //window.location.href = "default.asp?Tab=Categories";
		    window.location.href = "default.asp?" + txtQueryString.value;
		}
    }

    function ChangeTab(TabName) {
        window.location.href = "default.asp?Tab=" + TabName;
    }

    function goFind_onmouseover(e) {
        if (window.event)
            window.event.srcElement.style.cursor = "hand";
        else
            e.target.style.cursor = "pointer";
    }

    function GoTicketSearch_onclick() {
        //if (txtTicketSearch.value != "") {
       // if (cboSearchOptions.value == "1")
            window.location.href = ReplaceURLParameter("default.asp?" + txtQueryString.value, "TicketSearch", txtTicketSearch.value);
        //else
         //   window.location.href = "default.asp?TicketOwner=All&TicketStatus=All&TicketSearch=" + txtTicketSearch.value;
        //}

    }

    function txtTicketSearch_onkeypress() {
        if (window.event.keyCode == 13)
            GoTicketSearch_onclick();
    }

    function cboSearchOptions_onclick() {
        GoTicketSearch_onclick();
    }

    function ReplaceURLParameter(strURL, strKey, strValue) {
        if (strURL.indexOf("?") == -1)
            strURL = strURL + "?";

        var blnFound = false;
        var NewParameters = "";
        var MyArray = strURL.split("?");
        var URL = MyArray[0];
        var Parameters = MyArray[1];

        if (URL) {
            var MyArray = Parameters.split("&");
            for (var i in MyArray) {
                NewParameters = NewParameters + "&";
                if (MyArray[i] == "") {
                    NewParameters = NewParameters + strKey + "=" + strValue;
                    blnFound = true;
                }
                else if (MyArray[i].indexOf(strKey) == -1)
                    NewParameters = NewParameters + MyArray[i];
                else {
                    NewParameters = NewParameters + strKey + "=" + strValue;
                    blnFound = true;
                }
            }
        }
        if (!blnFound)
            NewParameters = NewParameters + "&" + strKey + "=" + strValue;

        return URL + "?" + NewParameters.substr(1);

    }

function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}
//-->
</SCRIPT>

<STYLE>
.MenuBar
{
}
.MenuBar A:link
{
    COLOR: white;
    FONT-FAMILY: Verdana;
    TEXT-DECORATION: none;
}
.MenuBar A:visited
{
    COLOR: white;
    FONT-FAMILY: Verdana;
    TEXT-DECORATION: none;
}
.MenuBar A:hover
{
    COLOR: yellow;
    FONT-FAMILY: Verdana;
    TEXT-DECORATION: none;
}
    
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
.DisplayBar
{
    width: 100%;
    padding: 2px;
    border: solid 2px gray;
    background-color: #DCDCDC; /* Gainsboro */
    white-space: nowrap;
    margin-top: 5px;
    margin-bottom: 10px;
}
.DisplayBar TD
{
    font-size: xx-small;
    font-family: Verdana;
}  

A:link
{
    color: blue;
}
A:visited
{
    color: blue;
}
A:hover
{
    color: red;
}
</STYLE>

</HEAD>


<body language="javascript" onmouseup="window_onmouseup()" onload="window_onload('<%=server.htmlencode(request("Tab"))%>');">
<%
	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0
    dim AnyFiltersSelected

    AnyFiltersSelected = false

%>
    <h3>Pulsar Support</h3>

    <table class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0">
        <tr bgcolor="<%=strTitleColor%>">
            <td id="CellTickets" style="display: none" width="10">
                        <font size="2" color="black"><b>&nbsp;<a href="javascript:ChangeTab('Tickets')">Tickets</a>&nbsp;</b></font></td>
            <td id="CellTicketsb" style="display: " width="10" bgcolor="wheat">
                        <font size="2" color="black"><b>&nbsp;Tickets&nbsp;</b></font></td>

            <td id="CellArticles" width="10">
                        <font size="2" color="white"><b>&nbsp;<a href="javascript:ChangeTab('Articles')">Articles</a>&nbsp;</b></font></td>
            <td id="CellArticlesb" style="display: none" width="10" bgcolor="wheat">
                        <font size="2" color="black"><b>&nbsp;Articles&nbsp;</b></font></td>
            <td id="CellCategories" width="10">
                        <font size="2" color="white"><b>&nbsp;<a href="javascript:ChangeTab('Categories')">Categories</a>&nbsp;</b></font></td>
            <td id="CellCategoriesb" style="display: none" width="10" bgcolor="wheat">
                        <font size="2" color="black"><b>&nbsp;Categories&nbsp;</b></font></td>
                </tr>
            </table>

<%
    dim cn, rs, cm
    dim CurrentUserID
    dim CurrentUser
    dim CurrentDomain

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


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
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = ""
	end if
	
	rs.Close



    dim strTicketOwner
    dim strTicketProject
    dim strTicketCategory
    dim strTicketCategoryDisplay
    dim strTicketStatus
    dim strTicketType
    dim strTicketSearch
    dim strSQL

    if trim(request("TicketOwner")) = "" then
        strTicketOwner = "Mine , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus"))  & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=All&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>All</a>"
        AnyFiltersSelected = true
    else
        strTicketOwner = "<a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus"))  & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>Mine</a> , All"
    end if

    strTicketProject = ""
    rs.open "spSupportProjectsListSelect",cn
    do while not rs.EOF
        if trim(rs("ID")) = trim(request("TicketProject")) then
            strTicketProject = strTicketProject & " , " & replace(rs("Name")," ","&nbsp;")
        else
            strTicketProject = strTicketProject & " , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & rs("ID") & """>" & replace(rs("Name")," ","&nbsp;")  & "</a>"
        end if
        rs.MoveNext
    loop
    rs.Close
    if trim(request("TicketProject")) = "" then
        strTicketProject = strTicketProject & " , All"
    else
        strTicketProject = strTicketProject & " , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject="">All</a>"
        AnyFiltersSelected = true
    end if
    if strTicketProject <> "" then
        strTicketProject = mid(strTicketProject,4) 
    end if

    if trim(request("TicketProject")) = "" then
        strTicketCategoryDisplay = "none"
    else
        strTicketCategory = ""
        rs.open "spSupportCategoryListSelect " & clng(request("TicketProject")),cn
        do while not rs.EOF
            if trim(rs("ID")) = trim(request("TicketCategory")) then
                strTicketCategory = strTicketCategory & " , " & replace(rs("Name")," ","&nbsp;")
            else
                strTicketCategory = strTicketCategory & " , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketCategory=" & rs("ID") & """>" & replace(rs("Name")," ","&nbsp;")  & "</a>"
            end if
            rs.MoveNext
        loop
        rs.Close

        if  strTicketCategory = "" then
            strTicketCategoryDisplay = "none"
        else
            strTicketCategoryDisplay = ""
            if trim(request("TicketCategory")) = "" then
                strTicketCategory = strTicketCategory & " , All"
            else
                strTicketCategory = strTicketCategory & " , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketCategory="">All</a>"
                AnyFiltersSelected = true
            end if
        
            if strTicketCategory <> "" then
                strTicketCategory = mid(strTicketCategory,4) 
            end if
        end if
    end if 

    if trim(request("TicketStatus")) = "" then 'Open
        strTicketStatus = "Open , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=New"">New</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=Closed"">Closed</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=In Progress"">In Progress</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=All"">All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketStatus")) = "All" then 'All
        strTicketStatus = "<a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>Open</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=New"">New</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=Closed"">Closed</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=In Progress"">In Progress</a> , All"
    elseif trim(request("TicketStatus")) = "Closed" then 'Closed
        strTicketStatus = "<a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>Open</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=New"">New</a> , Closed , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=In Progress"">In Progress</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=All"">All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketStatus")) = "New" then 'New
        strTicketStatus = "<a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>Open</a> , New , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=Closed"">Closed</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=In Progress"">In Progress</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=All"">All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketStatus")) = "In Progress" then 'Inprogress
        strTicketStatus = "<a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & """>Open</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=New"">New</a> , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=Closed"">Closed</a> , In Progress , <a href=""Default.asp?TicketType=" & server.HTMLEncode(request("TicketType")) & "&TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner")) & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=All"">All</a>"
        AnyFiltersSelected = true
    end if

    strTicketType = ""
    if trim(request("TicketType")) = "" then
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Question"">Question</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Issue"">Issue</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Suggestion"">Suggestion</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Request"">Request</a> , "
        strTicketType = strTicketType & "All"
    elseif trim(request("TicketType")) = "Question" then
        strTicketType = strTicketType & "Question , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Issue"">Issue</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Suggestion"">Suggestion</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Request"">Request</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & """>All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketType")) = "Issue" then
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Question"">Question</a> , "
        strTicketType = strTicketType & "Issue , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Suggestion"">Suggestion</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Request"">Request</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & """>All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketType")) = "Suggestion" then
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Question"">Question</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Issue"">Issue</a> , "
        strTicketType = strTicketType & "Suggestion , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Request"">Request</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & """>All</a>"
        AnyFiltersSelected = true
    elseif trim(request("TicketType")) = "Request" then
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Question"">Question</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Issue"">Issue</a> , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & "&TicketType=Suggestion"">Suggestion</a> , "
        strTicketType = strTicketType & "Request , "
        strTicketType = strTicketType & "<a href=""Default.asp?TicketCategory=" & server.HTMLEncode(request("TicketCategory"))  & "&TicketOwner=" & server.HTMLEncode(request("TicketOwner"))  & "&TicketProject=" & server.HTMLEncode(request("TicketProject")) & "&TicketStatus=" & server.HTMLEncode(request("TicketStatus")) & """>All</a>"
        AnyFiltersSelected = true
    end if


%>
    <div id="tabTickets">
        <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
	    <tr>
		<td valign="top">
                    <table>
                        <tr>
                            <td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b><br></font><font size="1"><a href="Default.asp">Reset</a></font></td>
                        </tr>
                    </table>
		<td width="100%">
                        <table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td valign="top"><b>Owner:</b></td>
				<td valign="top"><%=strTicketOwner%></td>
            </tr>
			<tr>
				<td valign="top"><b>Project:</b></td>
				<td valign="top"><%=strTicketProject%></td>
            </tr>

                            <tr style="display: <%=strTicketCategoryDisplay%>">
				<td valign="top"><b>Category:</b></td>
				<td valign="top"><%=strTicketCategory%></td>
            </tr>
			<tr>
				<td valign="top"><b>Status:</b></td>
				<td valign="top"><%=strTicketStatus%></td>
            </tr>
			<tr>
				<td valign="top"><b>Type:</b></td>
				<td valign="top"><%=strTicketType%></td>
            </tr>
                            <tr style="display: ">
				<td valign="top"><b>Search:</b></td>
				<td valign="top">
                                    <input id="txtTicketSearch" type="text" onkeypress="return txtTicketSearch_onkeypress()" value="<%=server.htmlencode(request("TicketSearch"))%>" /><img height="16px" width="16px" src="../images/go.gif" onmouseover="return goFind_onmouseover(event)" onclick="return GoTicketSearch_onclick()" />
                    &nbsp;
                    <select style="display: none" id="cboSearchOptions" onchange="cboSearchOptions_onclick()">
                        <%if  AnyFiltersSelected then %>
                            <option value="1" selected>Search Filtered Tickets</option>
                            <option value="2">Search All Tickets</option>
                        <%else%>
                            <option value="1">Search Filtered Tickets</option>
                            <option value="2" selected>Search All Tickets</option>
                        <%end if%>
                    </select>
                    </td>
            </tr>

            </table>
        </td>
        </tr>
        </table>

        <font size="2" face="Verdana"><b>Tickets</b></font>
        <font size="1" face="verdana"><a href="javascript:AddTicket();"><br><br>Add New Ticket</a></font>&nbsp;|&nbsp;<a target="_blank" href="StatusReport.asp">Status&nbsp;Report</a><br>
        <br>

        <%

            strSQL = "  Select i.id, i.summary, p.name as project, e1.Name as Submitter, e2.name as owner, c.name as Category, sis.Name as Status, DATEDIFF(d,i.DateCreated,getdate()) as DaysOld, DATEDIFF(HH,i.DateCreated,getdate()) as hoursOld, DATEDIFF(d,i.DateCreated,DateClosed) as DaysToClose, DATEDIFF(HH,i.DateCreated,DateClosed) as HoursToClose, i.DateCreated, i.DateClosed " & _
                     "  from SupportIssues i, SupportIssueStatus sis, SupportProject p, Employee e1, employee e2, SupportCategory c  " & _
                     "  Where sis.id = i.statusid " & _
                     "  and i.Categoryid = c.id " & _
                     "  and i.Submitterid = e1.id " & _
                     "  and i.Ownerid = e2.id " & _
                     "  and i.projectid = p.id " 

        if trim(request("TicketSearch")) <> "" then 'if the search box contains a value then do not depend of the other filters (ignore other filter settings)
                strSQl = strSQl & " and ("
                if isinteger(request("TicketSearch"))  then                     
                    strSQl = strSQl & " i.id=" & request("TicketSearch") & " "
                    strSQl = strSQl & " )"
                else  'if not a bumber then serach summary field to avoid error                   
                    strSQl = strSQl & " i.summary like '%" & request("TicketSearch") & "%' "
                    strSQl = strSQl & " )"                    
                end if
            
        else
            if request("TicketOwner") = ""  then                     
                strSQl = strSQl & " and i.ownerid= " & Currentuserid
            end if
            if request("TicketProject") <> "" then                     
                strSQl = strSQl & " and i.Projectid= " & clng(request("TicketProject"))
            end if
            if request("TicketProject") <> "" and request("TicketCategory") <> "" then                     
                strSQl = strSQl & " and i.categoryid= " & clng(request("TicketCategory"))
            end if
            if request("TicketStatus") <> "All" then                     
                if request("TicketStatus")  = "Closed" then
                    strSQl = strSQl & " and i.StatusID = 2 " 
                elseif request("TicketStatus")  = "In Progress" then
                    strSQl = strSQl & " and i.StatusID = 3 " 
                elseif request("TicketStatus")  = "New" then
                    strSQl = strSQl & " and i.StatusID = 1 " 
                else
                    strSQl = strSQl & " and i.Statusid in (1,3) "
                end if
            end if
            if request("TicketType") <> "" then                     
                if trim(request("TicketType")) = "Question" then
                    strSQl = strSQl & " and i.Typeid=0 " 
                elseif trim(request("TicketType")) = "Issue" then                     
                    strSQl = strSQl & " and i.Typeid=1 " 
                elseif trim(request("TicketType")) = "Suggestion" then                     
                    strSQl = strSQl & " and i.Typeid=2 " 
                elseif trim(request("TicketType")) = "Request" then                     
                    strSQl = strSQl & " and i.Typeid=3 " 
                end if
            end if
        end if
            rs.open strSQL,cn
            if rs.eof and rs.bof then
                response.write "<table id=""TicketTable"" bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%""><tr><td>No Tickets match your selected criteria.</td></tr></table>"
            else
                response.write "<table id=""TicketTable"" bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%"">"
                response.write "<thead><tr bgcolor=cornsilk>"
                response.write "<td onclick=""SortTable( 'TicketTable', 0 ,1,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>ID</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 1 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Project</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 2 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Category</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 3 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Submitter</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 4 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Owner</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 5 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Status</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 7 ,1,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Age</b></td>"
                response.write "<td style=""display:none""><b>hours</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 8 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Summary</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 9 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Created Date</b></td>"
                response.write "<td onclick=""SortTable( 'TicketTable', 10 ,0,1);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();""><b>Closed Date</b></td>"
                response.write "</tr></thead>"
            end if
            do while not rs.EOF
                response.write "<tr onclick=""javascript:ShowTicket(" & rs("ID") & ");"" onmouseover=""javascript:row_onmouseover();"" onmouseout=""javascript:row_onmouseout();"">"
                response.write "<td onclick=""SortTable( 'TicketTable', 0 ,0,2);"">" & rs("ID") & "</td>"
                response.write "<td nowrap>" & rs("Project") & "</td>"
                response.write "<td nowrap>" & rs("Category") & "</td>"
                response.write "<td nowrap>" & rs("Submitter") & "</td>"
                response.write "<td nowrap>" & rs("Owner") & "</td>"
                response.write "<td nowrap>" & rs("Status") & "</td>"
                if lcase(rs("Status")) = "closed" then
                    if rs("hourstoclose") < 48 then
                        response.write "<td nowrap>" & rs("hourstoclose") & "h*</td>"
                    else
                        response.write "<td nowrap>" & rs("daystoclose") & "d*</td>"
                    end if
                else
                    if rs("HoursOld") < 48 then
                        response.write "<td nowrap>" & rs("hoursold") & "h&nbsp;</td>"
                    else
                        response.write "<td nowrap>" & rs("daysold") & "d&nbsp;</td>"
                    end if
                end if
                if lcase(rs("Status")) = "closed" then
                    response.write "<td style=""display:none"" nowrap>" & rs("hourstoclose") & "h*</td>"
                else
                    response.write "<td style=""display:none"" nowrap>" & rs("hoursold") & "h&nbsp;</td>"
                end if
                response.write "<td>" & rs("Summary") & "</td>"
                response.write "<td>" & rs("DateCreated") & "</td>"
                if IsNull(rs("DateClosed")) then                    'if value is NULL then field the cell with blank HTML value
                    response.write "<td> &nbsp; </td>"
                else
                    response.write "<td>" & rs("DateClosed") & "</td>"
                end if
                response.write "</tr>"
                rs.MoveNext
            loop
            if not(rs.eof and rs.bof) then
                response.write "</table>"
            end if

            rs.Close

        '    response.write "<BR><BR>" & strSQL
        
        %>
    </div>
    <%
        dim strArticleOwner
        dim strArticleProject
        dim strArticleCategory
        dim strArticleStatus
        dim strArticleCategoryDisplay

        if trim(request("ArticleOwner")) = "" then
            strArticleOwner = "Mine , <a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleCategory=" & server.HTMLEncode(request("ArticleCategory"))  & "&ArticleOwner=All&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & """>All</a>"
        else
            strArticleOwner = "<a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleCategory=" & server.HTMLEncode(request("ArticleCategory"))  & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & """>Mine</a> , All"
        end if

        strArticleProject = ""
        rs.open "spSupportProjectsListSelect",cn
        do while not rs.EOF
            if trim(rs("ID")) = trim(request("ArticleProject")) then
                strArticleProject = strArticleProject & " , " & replace(rs("Name")," ","&nbsp;")
            else
                strArticleProject = strArticleProject & " , <a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & rs("ID") & """>" & replace(rs("Name")," ","&nbsp;")  & "</a>"
            end if
            rs.MoveNext
        loop
        rs.Close
        if trim(request("ArticleProject")) = "" then
            strArticleProject = strArticleProject & " , All"
        else
            strArticleProject = strArticleProject & " , <a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject="">All</a>"
        end if
        if strArticleProject <> "" then
            strArticleProject = mid(strArticleProject,4) 
        end if

        if trim(request("ArticleProject")) = "" then
            strArticleCategoryDisplay = "none"
        else
            strArticleCategory = ""
            rs.open "spSupportCategoryListSelect " & clng(request("ArticleProject")),cn
            do while not rs.EOF
                if trim(rs("ID")) = trim(request("ArticleCategory")) then
                    strArticleCategory = strArticleCategory & " , " & replace(rs("Name")," ","&nbsp;")
                else
                    strArticleCategory = strArticleCategory & " , <a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & "&ArticleCategory=" & rs("ID") & """>" & replace(rs("Name")," ","&nbsp;")  & "</a>"
                end if
                rs.MoveNext
            loop
            rs.Close

            if  strArticleCategory = "" then
                strArticleCategoryDisplay = "none"
            else
                strArticleCategoryDisplay = ""
                if trim(request("ArticleCategory")) = "" then
                    strArticleCategory = strArticleCategory & " , All"
                else
                    strArticleCategory = strArticleCategory & " , <a href=""Default.asp?Tab=Articles&ArticleStatus=" & server.HTMLEncode(request("ArticleStatus")) & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & "&ArticleCategory="">All</a>"
                end if
            
                if strArticleCategory <> "" then
                    strArticleCategory = mid(strArticleCategory,4) 
                end if
            end if

        end if 

        strArticleStatus = ""
        rs.open "spSupportArticleStatusSelect",cn
        do while not rs.EOF
            if trim(rs("ID")) = trim(request("ArticleStatus")) or (trim(rs("ID")) = "1" and trim(request("ArticleStatus"))= "" ) then
                strArticleStatus = strArticleStatus & " , " & rs("Name")
            elseif trim(rs("ID")) = "1" then
                strArticleStatus = strArticleStatus & " , <a href=""Default.asp?Tab=Articles&ArticleCategory=" & server.HTMLEncode(request("ArticleCategory"))  & "&ArticleStatus=&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & """>" & rs("Name")  & "</a>"
            else
                strArticleStatus = strArticleStatus & " , <a href=""Default.asp?Tab=Articles&ArticleCategory=" & server.HTMLEncode(request("ArticleCategory"))  & "&ArticleStatus=" & rs("ID") & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & """>" & rs("Name")  & "</a>"
            end if
            rs.MoveNext
        loop
        rs.Close
        if trim(request("ArticleStatus")) = "All" then
            strArticleStatus = strArticleStatus & " , All"
        else
            strArticleStatus = strArticleStatus & " , <a href=""Default.asp?Tab=Articles&ArticleCategory=" & server.HTMLEncode(request("ArticleCategory")) & "&ArticleOwner=" & server.HTMLEncode(request("ArticleOwner")) & "&ArticleProject=" & server.HTMLEncode(request("ArticleProject")) & "&ArticleStatus=All"">All</a>"
        end if
        if strArticleStatus <> "" then
            strArticleStatus = mid(strArticleStatus,4) 
        end if


    %>

    <div id="tabArticles" style="display: none">
        <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
	    <tr>    
		<td valign="top">
                    <table>
                        <tr>
                            <td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;<br></b></font><font size="1"><a href="Default.asp?Tab=Articles">Reset</a></font></td>
                        </tr>
                    </table>
		<td width="100%">
                        <table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td valign="top"><b>Owner:</b></td>
				<td valign="top"><%=strArticleOwner%></td>
            </tr>
			<tr>
				<td valign="top"><b>Project:</b></td>
				<td valign="top"><%=strArticleProject%></td>
            </tr>

                            <tr style="display: <%=strArticleCategoryDisplay%>">
				<td valign="top"><b>Category:</b></td>
				<td valign="top"><%=strArticleCategory%></td>
            </tr>

			<tr>
				<td valign="top"><b>Status:</b></td>
				<td valign="top"><%=strArticleStatus%></td>
            </tr>

                            <tr style="display: none">
				<td valign="top"><b>Search:</b></td>
				<td valign="top">
                    <input id="txtArticleSearch" type="text" /><img height="16px" width="16px" src="../images/go.gif" /></td>
            </tr>

            </table>
        </td>
        </tr>
        </table>
        <font size="2" face="Verdana"><b>Articles</b><br><br></font>
        <font size="1" face="verdana"><a href="javascript:AddArticle();">Add New Article</a></font>
        <br>
        <br>
        <%
            strSQL = "  Select a.ID, a.Title,e.Name as Author, p.Name as project,c.Name as Category, sas.name as status " & _
                     "  from SupportArticle a, Employee e, SupportCategory c, SupportProject p, SupportArticleStatus sas" & _
                     "  Where sas.id = a.statusid" & _
                     "  and c.supportprojectid = p.id " & _
                     "  and a.AuthorID = e.ID" & _
                     "  and a.SupportCategoryID = c.id " 

            if request("ArticleOwner") = "" then                     
                strSQl = strSQl & " and a.authorid= " & Currentuserid
            end if
            if request("ArticleProject") <> "" then                     
                strSQl = strSQl & " and p.id= " & clng(request("ArticleProject"))
            end if
            if request("ArticleProject") <> "" and request("ArticleCategory") <> "" then                     
                strSQl = strSQl & " and a.supportcategoryid= " & clng(request("ArticleCategory"))
            end if
            if request("ArticleStatus") <> ""  and request("ArticleStatus") <> "All"  then                     
                 strSQl = strSQl & " and a.StatusID = " &  clng(request("ArticleStatus"))
            end if

            rs.open strSQL,cn
            if rs.eof and rs.bof then
                response.write "<table bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%""><tr><td>No Tickets match your selected criteria.</td></tr></table>"
            else
                response.write "<table bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%"">"
                response.write "<tr bgcolor=cornsilk><td><b>ID</b></td><td><b>Owner</b></td><td><b>Project</b></td><td><b>Category</b></td><td><b>Title</b></td></tr>"
            end if
            do while not rs.EOF
                response.write "<tr onclick=""javascript:ShowArticle(" & rs("ID") & ");"" onmouseover=""javascript:row_onmouseover();"" onmouseout=""javascript:row_onmouseout();"">"
                response.write "<td>" & rs("ID") & "</td>"
                response.write "<td nowrap>" & rs("Author") & "</td>"
                response.write "<td nowrap>" & rs("Project") & "</td>"
                response.write "<td nowrap>" & rs("category") & "</td>"
                response.write "<td>" & rs("Title") & "</td>"
                response.write "</tr>"
                rs.MoveNext
            loop
            if not(rs.eof and rs.bof) then
                response.write "</table>"
            end if

            rs.Close

             %>
    </div>

    <%
        dim strCategoryOwner
        dim strCategoryProject
        dim strCategoryStatus

        if trim(request("CategoryOwner")) = "" then
            strCategoryOwner= "<a href=""Default.asp?Tab=Categories&CategoryStatus=" & server.HTMLEncode(request("CategoryStatus")) & "&CategoryOwner=Mine&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>Mine</a> , All"
        else
            strCategoryOwner = "Mine , <a href=""Default.asp?Tab=Categories&CategoryStatus=" & server.HTMLEncode(request("CategoryStatus")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>All</a>"
        end if

        strCategoryProject = ""
        rs.open "spSupportProjectsListSelect",cn
        do while not rs.EOF
            if trim(rs("ID")) = trim(request("CategoryProject")) then
                strCategoryProject = strCategoryProject & " , " & replace(rs("Name")," ","&nbsp;")
            else
                strCategoryProject = strCategoryProject & " , <a href=""Default.asp?Tab=Categories&CategoryStatus=" & server.HTMLEncode(request("CategoryStatus")) & "&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & rs("ID") & """>" & replace(rs("Name")," ","&nbsp;")  & "</a>"
            end if
            rs.MoveNext
        loop
        rs.Close
        if trim(request("CategoryProject")) = "" then
            strCategoryProject = strCategoryProject & " , All"
        else
            strCategoryProject = strCategoryProject & " , <a href=""Default.asp?Tab=Categories&CategoryStatus=" & server.HTMLEncode(request("CategoryStatus")) & "&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject="">All</a>"
        end if
        if strCategoryProject <> "" then
            strCategoryProject = mid(strCategoryProject,4) 
        end if

        strCategoryStatus = ""
        if request("CategoryStatus") = "" then
            strCategoryStatus = strCategoryStatus & " Active , <a href=""Default.asp?Tab=Categories&CategoryStatus=Inactive&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>Inactive</a> , <a href=""Default.asp?Tab=Categories&CategoryStatus=All&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>All</a>"
        elseif request("CategoryStatus") = "Inactive" then
            strCategoryStatus = strCategoryStatus & "<a href=""Default.asp?Tab=Categories&CategoryStatus=&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>Active</a> , Inactive , <a href=""Default.asp?Tab=Categories&CategoryStatus=All&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>All</a>"
        else 'All
            strCategoryStatus = strCategoryStatus & "<a href=""Default.asp?Tab=Categories&CategoryStatus=Active&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>Active</a> , <a href=""Default.asp?Tab=Categories&CategoryStatus=Inactive&CategoryOwner=" & server.HTMLEncode(request("CategoryOwner")) & "&CategoryProject=" & server.HTMLEncode(request("CategoryProject")) & """>Inactive</a>, All"
        end if

    %>

    <div id="tabCategories" style="display: none">
        <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
	    <tr>
		<td valign="top">
                    <table>
                        <tr>
                            <td valign="top"><font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b><br></font><font size="1"><a href="Default.asp?Tab=Categories">Reset</a></font></td>
                        </tr>
                    </table>
                    <td width="100%" valign="top">
                        <table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td valign="top"><b>Project:</b></td>
				<td valign="top"><%=strCategoryProject%></td>
            </tr>
			<tr>
				<td valign="top"><b>Owner:</b></td>
				<td valign="top"><%=strCategoryOwner%></td>
            </tr>
			<tr>
				<td valign="top"><b>Status:</b></td>
				<td valign="top"><%=strCategoryStatus%></td>
            </tr>
            </table>
        </td>
        </tr>
        </table>
        <font size="2" face="Verdana"><b>Categories</b><br><br></font>
        <font size="1" face="verdana"><a href="javascript:AddCategory();">Add New Category</a></font>
        <br>
        <br>
        <%
            strSQL = "  Select p.ID as ProductID, c.ID as categoryid, p.Name as Project, c.Name as category, e.Name as ownerName, e.ID as OwnerID, c.NotificationList, c.TrackTickets " & _
                     "  from SupportProject p, SupportCategory c, Employee e " & _
                     "  where e.ID = c.OwnerID " & _
                     "  and c.SupportProjectID = p.id "
            if trim(request("CategoryStatus")) = "Active" or trim(request("CategoryStatus")) = "" then
                strSQL = strSQL &  " and c.Active=1 "
            elseif trim(request("CategoryStatus")) = "Inactive" then
                strSQL = strSQL &  " and c.Active=0 "
            end if

            if trim(request("CategoryProject")) <> "" then
                strSQL = strSQL & " and p.id = " & trim(clng(request("CategoryProject"))) & " "
            end if

            if trim(request("CategoryOwner")) <> "" then
                strSQL = strSQL & " and e.id = " & trim(clng(Currentuserid))  & " "
            end if

            strSQL = strSQL & "  order by p.DisplayOrder,p.Name, c.DisplayOrder,c.Name" 

            rs.open strSQL,cn
            if rs.eof and rs.bof then
                response.write "<table bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%""><tr><td>No Categories match your selected criteria.</td></tr></table>"
            else
                response.write "<table bgcolor=ivory border=1 bordercolor=tan style=""border-width:1px;"" cellpadding=2 cellspacing=0 width=""100%"">"
                response.write "<tr bgcolor=cornsilk><td><b>Project</b></td><td><b>Category</b></td><td><b>Owner</b></td><td><b>Notifications</b></td></tr>"
            end if
            do while not rs.EOF
                response.write "<tr onclick=""javascript:ShowCategory(" & rs("categoryid") & ");"" onmouseover=""javascript:row_onmouseover();"" onmouseout=""javascript:row_onmouseout();"">"
                response.write "<td>" & rs("project") & "</td>"
                response.write "<td nowrap>" & rs("category") & "</td>"
                response.write "<td nowrap>" & rs("ownername") & "</td>"
                response.write "<td nowrap>" & rs("notificationlist") & "&nbsp;</td>"
                response.write "</tr>"
                rs.MoveNext
            loop
            if not(rs.eof and rs.bof) then
                response.write "</table>"
            end if

            rs.Close

             %>
    </div>

<%
    set rs = nothing
    cn.Close
    set cn = nothing
%>


    <div id="mnuPopup" style="display: none; position: absolute; width: 2px; height: 2px; left: 0px; top: 0px; padding: 0px; background: white; border: 1px solid gainsboro; z-index: 100">
        <div style="border-right: black 1px solid; border-top: black 1px solid; left: 0px; border-left: black 1px solid; border-bottom: black 1px solid; position: relative; top: 0px">
            <div onmouseover="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
                <font face="Arial" size="2">
	    <SPAN onclick="javascript:EmailSubmitter();" >&nbsp;&nbsp;&nbsp;Email&nbsp;Submitter</SPAN></font>
            </div>

            <div>
                <hr width="95%">
            </div>

            <div onmouseover="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
                <font face="Arial" size="2">
        <SPAN onclick="javascript:StatusReport(0);" >&nbsp;&nbsp;&nbsp;Convert&nbsp;To&nbsp;Action&nbsp;Item&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></font>
            </div>
            <div id="CustomMenuOptions"><%=CustomStatusReports%></div>
        
            <div>
                <span>
                    <hr width="95%">
                </span>
            </div>

            <div onmouseover="this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
                <font face="Arial" size="2">
        <SPAN onclick="javascript:CustomStatusReport();" >&nbsp;&nbsp;&nbsp;Properties&nbsp;&nbsp;&nbsp;</SPAN></font>
            </div>

        </div>

</div>
    <%
    function IsInteger( strValue)
        Set re = new RegExp
        re.Pattern = "^\d+$"
        re.Global = true
        IsInteger = re.test(strValue)
    end function

    


        dim strQueryString
        strQueryString = request.querystring
        if request("Tab") = "" and strQueryString = "" then
            strQueryString = "Tab=Tickets"
        elseif request("Tab") = "" and strQueryString <> "" then
            strQueryString = "Tab=Tickets&" & strQueryString
        end if
    %>
    
    <input style="width: 100%; display: none" id="txtQueryString" type="text" value="<%=strQueryString%>" />
</body>
</html>





<%@ Language=VBScript %>

<%
	
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
    Dim AppRoot : AppRoot = Session("ApplicationRoot")	 
	  
%>
<HTML>
<HEAD>
<title>Ticket</title>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="includes/client/json2.js"></script>
<script type="text/javascript" src="includes/client/json_parse.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

<!--

//-->
</SCRIPT>

<STYLE>
h3
    {
        font-family: Verdana;
        font-size:x-small;
    }
td,textarea,input,select
{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
  fieldset
  {
      background-color:White;
      padding-left:5px;
      padding-right:5px;
      padding-top:5px;
      padding-bottom:5px;
      margin: 10px 5px 10px 5px;
  }

A:link
{
    COLOR: blue;
}
A:visited
{
    COLOR: blue;
}

A:hover
{
    COLOR: red;
}    
</STYLE>

<script language="javascript">

    function GetCategories(){
        var i;
        var OptionArray;
        frmMain.cboCategory.options.length = 0
        //frmMain.cboCategory.options[frmMain.cboCategory.options.length] = new Option('','');
        for(i=0;i<CategoryLookup.options.length;i++)
            if(CategoryLookup.options[i].value== frmMain.cboProject.options[frmMain.cboProject.selectedIndex].value)
                {
                OptionArray = CategoryLookup.options[i].text.split("|");
                frmMain.cboCategory.options[frmMain.cboCategory.options.length] = new Option(OptionArray[1],OptionArray[0]);
                }

                
    }

    function TitleMouseOver(){
        event.srcElement.style.color="red";
        event.srcElement.style.cursor="hand";
    }

    function TitleMouseOut(){
        event.srcElement.style.color="black";
        event.srcElement.style.cursor="default";
    }

    function TitleClick(ID){
        if (document.all("SearchRow" + ID).style.display=="" )
            document.all("SearchRow" + ID).style.display="none";
        else
            document.all("SearchRow" + ID).style.display=""
    }
    function TicketTitleClick(ID){
        if (document.all("TicketRow" + ID).style.display=="" )
            document.all("TicketRow" + ID).style.display="none";
        else
            document.all("TicketRow" + ID).style.display=""
    }

    function window_onload(){

        modalDialog.load();
    }

    function ShowArtileList(){

        var url = "ArticleList.asp";
        modalDialog.open({ dialogTitle: 'Supported Articles', dialogURL: '' + url + '', dialogHeight: 400, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
    }


    function ShowArtileList_return(strID)
    {
        var IDArray;
        var i;
        var strURL

        if (txtSubmitterPartnerID.value != "1")
            strURL = "https://<%=Application("Excalibur_ODM_ServerName")%>/excalibur/";
        else
            strURL = "http://<%=Application("Excalibur_ServerName")%>/Excalibur/";


        if (typeof (strID) != "undefined") {
            if (strID != "0") {
                IDArray = strID.split(",");
                frmMain.txtResolution.value = frmMain.txtResolution.value + "\r\r<u>Related Articles</u>\r"
                for (i = 0; i < IDArray.length; i++)
                    frmMain.txtResolution.value = frmMain.txtResolution.value + "<a href=\"" + strURL + "Support/Preview.asp?ID=" + IDArray[i] + "\">Article " + IDArray[i] + "</a>\r"

            }
        }
    }


    function UploadZip(ID) {
        //save ID for return function: ---
        globalVariable.save(ID, 'main_uploadzip_ID');

        var url = "<%=AppRoot %>/PMR/SoftpaqFrame.asp?Title=Upload Support Attachments&Page=../common/fileupload.aspx";
        modalDialog.open({ dialogTitle: 'Upload Support Attachments', dialogURL: '' + url + '', dialogHeight: 250, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
    }


    function UploadZip_return(strPath) {
        var ID = globalVariable.get('main_uploadzip_ID');
        if (typeof (strPath) != "undefined") {
            $("#UploadAddLinks" + ID).hide();
            $("#UploadRemoveLinks" + ID).show();
            $("#UploadPath" + ID).text(strPath.substr(strPath.lastIndexOf("\\") + 1, strPath.length));
            $("#txtAttachmentPath" + ID).val(strPath);
        }
    }

    function RemoveUpload(ID) {
        $("#UploadAddLinks" + ID).show();
        $("#UploadRemoveLinks" + ID).hide();
        $("#UploadPath" + ID).text("");
        $("#txtAttachmentPath" + ID).val("");
    }

    function cboCategory_onchange(){
        frmMain.CategoryChanged.value="1";
    }

    function cboStatus_onchange(){
        if(frmMain.cboStatus.options[frmMain.cboStatus.selectedIndex].text=="Closed")
            {
            RequireResolution.style.display="";
            NotifyRow.style.display = "";
           // ActionItemRow.style.display = "";
            }
        else
            {
            RequireResolution.style.display="none";
            NotifyRow.style.display = "none";
           // ActionItemRow.style.display = "none";
            }
    }


    function cmdAdd_onclick() {
	    var strResult;
	    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../Email/AddressBook.asp?AddressList=' + frmMain.txtNotify.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
	    globalVariable.save('txtNotify', 'email_field');
    }

</script>

</HEAD>


<BODY bgcolor="Ivory" onload="window_onload();">
<h3>Pulsar Support</h3>
<%  
	dim cn, rs, cm
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


    dim ProjectList
    dim CategoryLookupList
    dim CategoryList
    dim strSearchResults
    dim CurrentStep
    dim strDisplayTab1
    dim strDisplayTab2
    dim strDisplayTab3
    dim strDisplayTab4
    dim strRequirementsTemplate
    dim strRequirementsField



    CategoryLookupList=""
    rs.open "spSupportCategoryListSelect",cn
    CategoryLookupList = ""
    do while not rs.EOF
            CategoryLookupList = CategoryLookupList & "<option value=""" & rs("SupportProjectID") & """>" & rs("ID") & "|" & rs("Name") & "|" & rs("RequiredFields") & "</option>"
        rs.MoveNext
    loop
    rs.close

    strTicketNumber = clng(request("ID"))

    rs.open "spSupportTicketSelect " & clng(request("ID")),cn
    if rs.eof and rs.bof then
        response.write "Unable to find the selected ticket."
    else
        dim strSummary 
        dim strDetails
        dim strResolution
        dim strProjectID
        dim strCategoryID
        dim strSubmitterName
        dim strSubmitterEmail
        dim TypeList
        dim strTypeID
        dim StatusList
        dim strStatusID
        dim OwnerList
        dim strOwnerID
        dim strActionItemID
        dim strOwnerName
        dim strAttachment1
        dim strAttachment2
        dim strAttachment3
        dim strPartnerID
        dim strShowResponseRequired
        dim PathArray
        dim DateCreated
        dim DateClosed

        strSummary = rs("Summary") & ""
        strDetails = rs("Details") & ""
        strResolution = rs("Resolution") & ""
        strProjectID = rs("ProjectID") & ""
        strCategoryID = rs("CategoryID") & ""
        strSubmitterName = rs("SubmitterName") & ""
        strSubmitterEmail = rs("SubmitterEmail") & ""
        strActionItemID = trim(rs("ActionItemID") & "")
        strTypeID = rs("TypeID") & ""
        strStatusID = rs("StatusID") & ""
        strOwnerID = rs("OwnerID") & ""
        strOwnerName = rs("OwnerName") & ""
        strAttachment1 = rs("Attachment1") & ""
        strAttachment2 = rs("Attachment2") & ""
        strAttachment3 = rs("Attachment3") & ""
        strPartnerID = rs("SubmitterPartnerID") & ""
        DateCreated = rs("DateCreated") & ""
        DateClosed = rs("DateClosed") & ""

        rs.Close

        if trim(strAttachment1) <> "" and instr(strAttachment1,"\")> 0 then
            PathArray = split(strAttachment1,"\")
        
            strAttachment1 = "<a target=_blank href=""file://" & strAttachment1 & """>" & PathArray(ubound(PathArray)) & "</a>"
        end if
        if trim(strAttachment2) <> "" and instr(strAttachment2,"\")> 0 then
            PathArray = split(strAttachment2,"\")
        
            strAttachment2 = "<a target=_blank href=""file://" & strAttachment2 & """>" & PathArray(ubound(PathArray)) & "</a>"
        end if
        if trim(strAttachment3) <> "" and instr(strAttachment3,"\")> 0 then
            PathArray = split(strAttachment3,"\")
        
            strAttachment3 = "<a target=_blank href=""file://" & strAttachment3 & """>" & PathArray(ubound(PathArray)) & "</a>"
        end if

        if clng(strStatusID) = 2 then
            strShowResponseRequired = ""
        else
            strShowResponseRequired="none"
        end if 

        OwnerList = "<option selected value=" & strOwnerID & ">" & strOwnerName & "</option>"
        rs.open "spSupportAdminSelect",cn
        do while not rs.EOF
            if trim(strOwnerID) <> trim(rs("ID")) then
                OwnerList = OwnerList & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.Close

        TypeList=""
        rs.open "spSupportIssueTypeSelect", cn
        do while not rs.EOF
            if trim(strTypeID) = trim(rs("ID") & "") then
                TypeList = TypeList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            else
                TypeList = TypeList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.Close

        rs.open "spSupportIssueStatusSelect", cn
        StatusList = ""
        do while not rs.EOF
            if trim(strStatusID) = trim(rs("ID") & "") then
                StatusList = StatusList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            else
                StatusList = StatusList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.Close

        rs.open "spSupportProjectsListSelect",cn
        ProjectList = ""
        do while not rs.EOF
            if trim(strProjectID) = trim(rs("ID")) then
                ProjectList = ProjectList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            else
                ProjectList = ProjectList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.close


        CategoryList=""
        if trim(strProjectID) <> "" then
            rs.open "spSupportCategoryListSelect " & clng(strProjectID),cn
            do while not rs.EOF
                if trim(strCategoryID) = trim(rs("ID")) then
                    CategoryList = CategoryList & "<option selected value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
                else
                    CategoryList = CategoryList & "<option value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
                end if
                rs.MoveNext
            loop
            rs.close
        end if



%>
    <form id="frmMain" method="post" action="TicketSave.asp">
    <b>Ticket #:</b> <%=strTicketNumber%><br>
    <table cellpadding=2 border=1 cellspacing=0 bordercolor=tan bgcolor="cornsilk" style="border-width:1px;width:100%">
        <tr>
            <td><b>Question/Request:</b>&nbsp;<font color=red>*</font></td>
            <td colspan=3><input maxlength="500" id="txtSubject" name="txtSubject" type="text" style="width:100%" value="<%= server.htmlencode(strSummary)%>"></td>
        </tr>
        <tr>
            <td><b>Project:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboProject" name="cboProject" style="width:100%" onchange="javascript:GetCategories();">
                    <%=ProjectList%>
                </select></td>
            <td><b>Submitter:</b>&nbsp;</td>
            <td width="50%"><font size=1 face=verdana><a href="mailto:<%=strSubmitterEmail%>?Subject=Mobile Tools Support Ticket <%=strTicketNumber%>&Body=Ticket Summary: <%=server.htmlencode(replace(strSummary,"""","'"))%>"><%=strSubmitterName%></a></font>
            </td>
        </tr>
        <tr>
            <td><b>Category:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboCategory" name="cboCategory" style="width:100%">
                    <%=CategoryList%>
                </select>
            <td><b>Owner:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboOwner" name="cboOwner" style="width:100%">
                    <%=OwnerList%>
                </select>
            </td>
        </tr>
        <tr>
            <td><b>Request Type:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboType" name="cboType" style="width:100%">
                    <%=TypeList%>
                </select>
            <td><b>Status:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboStatus" name="cboStatus" style="width:100%" onchange="javascript:cboStatus_onchange();">
                    <%=StatusList%>
                </select>
            </td>
        </tr>
        <tr id=NotifyRow style="display:none">
            <td nowrap valign=top><b>Notify:</b>&nbsp;</td>
            <td colspan=3>
                        <TABLE width=100% cellpadding=0 cellspacing=0 border=0><TR><TD width=100%><input maxlength="500" id="txtNotify" name="txtNotify" type="text" style="width:100%" value="<%= server.htmlencode(strSubmitterEmail)%>"></TD><TD><INPUT type="button" value="Add" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()"></TD></TR></TABLE>
			            <INPUT <%=strCopyMe%> type="checkbox" checked id=chkCopyMe name=chkCopyMe value="1"> Copy Me&nbsp;&nbsp;<INPUT  <%=strCopyTeam%> type="checkbox" id=chkCopyTeam name=chkCopyTeam value="1"> Copy Development Team            
            </td>
        </tr>
        <tr>
            <td nowrap valign=top><b>Action Item:</b>&nbsp;</td>
            <td colspan=3>
                <%if strActionItemID = "" then %>
                    <INPUT type="checkbox" id=chkActionItem name=chkActionItem value="1"> Convert&nbsp;to&nbsp;Action&nbsp;Item            
                <%else%>
                    <INPUT style="display:none" type="checkbox" id=chkActionItem2 name=chkActionItem value="">Converted&nbsp;to&nbsp;Action&nbsp;Item:&nbsp;<a target=_blank href="../actions/action.asp?ID=<%=strActionItemID%>"><%=strActionItemID%></a>
                <%end if%>
            </td>
        </tr>
        <tr>
            <td valign=top><b>Response:</b><font id="RequireResolution" style="display:<%=strShowResponseRequired%>" color=red>&nbsp;*</font>
            <br><a href="javascript:ShowArtileList();">Add&nbsp;Article&nbsp;Links</a>
            </td>
            <td colspan=3 width="100%">
                <textarea id="txtResolution" style="width:100%;" name="txtResolution" rows=10><%=server.htmlEncode(strResolution)%></textarea>
            </td>
        </tr>
        <tr>
            <td valign=top><b>Working&nbsp;Notes:</b></td>
            <td colspan=3 width="100%">
                <textarea id="txtDetails" style="width:100%;" name="txtDetails" rows=12><%=server.htmlEncode(strDetails)%></textarea>
            </td>
        </tr>
        <%if trim(strAttachment1) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;1:</b></td>
            <td valign=top colspan="3"><%=strAttachment1%></td>
        </tr>
        <%if trim(strAttachment2) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;2:</b></td>
            <td valign=top colspan="3"><%=strAttachment2%></td>
        </tr>
        <%if trim(strAttachment3) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;3:</b></td>
            <td valign=top colspan="3"><%=strAttachment3%></td>
        </tr>
        <tr>
            <td><strong>Created Date:</strong></td>
            <td> <%=DateCreated%></td>
            <td><strong>Closed&nbsp;Date:</strong></td>
            <td> <%=DateClosed%></td>                                   
        </tr>
    </table>
    <input id="tagStatus" type="hidden" name="tagStatus" value="<%=trim(strStatusID)%>" />
    <input id="tagOwner" type="hidden" name="tagOwner" value="<%=trim(strOwnerID)%>" />
    <input id="txtID" type="hidden" name="txtID" value="<%=trim(clng(request("ID")))%>" />
    <input id="txtProjectName" type="hidden" name="txtProjectName" value="<%=trim(strProjectName)%>" />
    
    </form>
    <%end if%>
    <input id="txtSubmitterPartnerID" type="hidden" value="<%=strPartnerID%>" />
    <select id="CategoryLookup" style="display:none">
    <%=CategoryLookupList%>
    </select>
<%

    set rs = nothing
    cn.Close
    set cn = nothing
%>
</BODY>
</HTML>





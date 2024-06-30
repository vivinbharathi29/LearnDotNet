<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
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
a:link, a:visited
{
    color: blue;
}

a:hover
{
    color: red;
    text-decoration: none;
}
    
</STYLE>

<script language="javascript">

    function GetCategories(){
        var i;
        var OptionArray;
        frmMain.cboCategory.options.length = 0
        frmMain.cboCategory.options[frmMain.cboCategory.options.length] = new Option('','');
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



    function Preview(ID)
    {
        frmPreview.PreviewSummary.value = frmMain.txtTitle.value;
        frmPreview.PreviewDetails.value = frmMain.txtArticleText.value;
        frmPreview.PreviewURL.value = frmMain.txtURL.value;
        frmPreview.submit();
    }


function window_onload() {
	//frames.myEditor.document.body.contentEditable = "True";
	//frames.myEditor.document.body.innerHTML = "<font face=verdana size=2>" + frmMain.txtArticleText.value + "</font>";
	//frames.myEditor.focus();
}

function AddLink(){
    frmMain.txtArticleText.value = frmMain.txtArticleText.value + "\r\r<u>More Information</u>\r<a target=_blank href=\"Articles\\" + frmMain.cboProject.options[frmMain.cboProject.selectedIndex].text + "\\xxxxx.htm\">Link 1</a>";
}

function disableEnterKey(e)
{
     var key;

     if(window.event)
          key = window.event.keyCode;     //IE
     else
          key = e.which;     //firefox

     if(key == 13)
          return false;
     else
          return true;
}



</script>

</HEAD>


<BODY bgcolor="Ivory" onload="window_onload();">
<h3>Support Article</h3>
<%

    
	dim cn, rs, cm


    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    CategoryLookupList=""
    rs.open "spSupportCategoryListSelect",cn
    CategoryLookupList = ""
    do while not rs.EOF
            CategoryLookupList = CategoryLookupList & "<option value=""" & rs("SupportProjectID") & """>" & rs("ID") & "|" & rs("Name") & "|" & rs("RequiredFields") & "</option>"
        rs.MoveNext
    loop
    rs.close


    dim CategoryLookupList
    dim StatusList
    dim blnFound
    dim strTitle
    dim strArticleText
    dim strKeywords
    dim strURL
    dim CategoryList
    dim ProjectList
    dim strProjectID
    dim strCategoryID
    dim strProjectName
    dim strCategoryName
    dim strStatusID
    dim strOwnerID
    dim strOwnerName
    dim OwnerList

    blnFound = true
    if request("ID") = "" then
        strTitle=""
        strArticleText=""
        strStatusID="1"
    else
        rs.open "spSupportArticleSelect " & clng(request("ID")),cn
        if rs.eof and rs.bof then
            blnFound=false
        else
            strTitle = rs("Title") & ""
            strArticleText = rs("ArticleText") & ""
            strProjectID = rs("ProjectID") & ""
            strCategoryID = rs("CategoryID") & ""
            strProjectName = rs("ProjectName") & ""
            strCategoryName = rs("CategoryName") & ""
            strKeywords = rs("Keywords") & ""
            strURL = rs("ArticleURL") & ""
            strStatusID = rs("StatusID") & ""
            strOwnerID = rs("OwnerID") & ""
            strOwnerName = rs("OwnerName") & ""
        end if
        rs.Close
    end if
    if not blnFound then
        response.write "Unable to find the selected article."
    else


        rs.open "spSupportProjectsListSelect",cn
        ProjectList = ""
        blnFound=false
        do while not rs.EOF
            if trim(strProjectID) = trim(rs("ID")) then
                ProjectList = ProjectList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
                blnFound = true
            else
                ProjectList = ProjectList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.close
        if not blnFound then
            ProjectList = ProjectList & "<option selected value=""" & strProjectID & """>" & strProjectName & "</option>"
        end if

        CategoryList=""
        if trim(strProjectID) <> "" then
            rs.open "spSupportCategoryListSelect " & clng(strProjectID),cn
            blnFound = false
            do while not rs.EOF
                if trim(strCategoryID) = trim(rs("ID")) then
                    CategoryList = CategoryList & "<option selected value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
                    blnFound = true
                else
                    CategoryList = CategoryList & "<option value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
                end if
                rs.MoveNext
            loop
            rs.close
            if not blnFound then
                CategoryList = CategoryList & "<option selected value=""" & strCategoryID & """>" &  strCategoryName & "</option>"
            end if
        end if


        rs.open "spSupportArticleStatusSelect",cn
        StatusList = ""
        do while not rs.EOF
            if trim(strStatusID) = trim(rs("ID")) then
                StatusList = StatusList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            else
                StatusList = StatusList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.close

        OwnerList = "<option selected value=" & strOwnerID & ">" & strOwnerName & "</option>"
        rs.open "spSupportAdminSelect",cn
        do while not rs.EOF
            if trim(strOwnerID) <> trim(rs("ID")) then
                OwnerList = OwnerList & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.Close

        

%>
    <form id="frmMain" method="post" action="ArticleSave.asp">
    <table cellpadding=2 border=1 cellspacing=0 bordercolor=tan bgcolor="cornsilk" style="border-width:1px;width:100%">
        <tr>
            <td valign=top><b>Title:</b>&nbsp;<font color=red>*</font></td>
            <td colspan=3><textarea id="txtTitle" name="txtTitle" rows=2 style="width:100%"><%= server.htmlencode(strTitle)%></textarea></td>
        </tr>
        <tr>
            <td valign=top><b>Article&nbsp;Text:</b>&nbsp;<br><a href="javascript:Preview();">Preview</a><br><br><a href="javascript:AddLink();">Add Link</a></td>
            <td colspan=3><textarea style="width:100%" id="txtArticleText" name="txtArticleText" rows=23><%= strArticleText%></textarea></td>
        </tr>
        <tr>
            <td valign=top><b>URL:</b>&nbsp;<br><a href="javascript:Preview();">Preview</a></td>
            <td colspan=3><textarea id="txtURL" name="txtURL" rows=2 style="width:100%"><%= server.htmlencode(strURL)%></textarea></td>
        </tr>
        <tr>
            <td valign=top><b>Keywords:</b>&nbsp;</td>
            <td colspan=3><textarea id="txtKeywords" name="txtKeywords" rows=2 style="width:100%"><%= server.htmlencode(strKeywords)%></textarea></td>
        </tr>
        <tr>
            <td><b>Status:</b></td>
            <td width="50%">
                <select id="cboStatus" name="cboStatus" style="width:100%">
                    <%=StatusList %>
                </select></td>
            <td><b>Project:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboProject" name="cboProject" style="width:100%" onchange="javascript:GetCategories();">
                    <%=ProjectList%>
                </select>
            </td>
        </tr>
        <tr>
            <td><b>Owner:</b>&nbsp;<font color=red>*</font></td>
            <td>
                <select id="cboOwner" name="cboOwner" style="width:100%">
                    <%=OwnerList%>
                </select></td>
            <td><b>Category:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboCategory" name="cboCategory" style="width:100%">
                    <%=CategoryList%>
                </select>
        </tr>
        <tr style="display: none">
            <td><b>Access:</b></td>
            <td>
                <input id="chkHP" name="chkHP" type="checkbox" />&nbsp;HP&nbsp;&nbsp;&nbsp;&nbsp;<input id="chkODM" name="chkODM" type="checkbox" />&nbsp;ODM&nbsp;&nbsp;&nbsp;&nbsp;<input id="chkVendor" name="chkVendor" type="checkbox" />&nbsp;Vendor</td>
        </tr>
        <%if trim(request("ID")) <> "" then%>
            <tr>
        <%else%>
            <tr style="display:none">
        <%end if%>
            <td><b>HP&nbsp;Link</b></td>
            <td colspan=3>
                <input readonly style="background-color:ivory;width:100%" id="txtArticleLink" type="text" value="http://<%=Application("Excalibur_ServerName")%>/Excalibur/support/Preview.asp?ID=<%=trim(request("ID"))%>"/>
            </td>
        </tr>
        </tr>
        <%if trim(request("ID")) <> "" then%>
            <tr>
        <%else%>
            <tr style="display:none">
        <%end if%>
            <td><b>Partner&nbsp;Link</b></td>
            <td colspan=3>
                <input readonly style="background-color:ivory;width:100%" id="txtArticleLink" type="text" value="https://<%=Application("Excalibur_ODM_ServerName") %>/excalibur/support/Preview.asp?ID=<%=trim(request("ID"))%>"/>
            </td>
        </tr>
    </table>
    <input id="txtID" name="txtID" type="hidden" value="<%=request("ID")%>"/>
    </form>
<%

    end if





    set rs = nothing
    cn.Close
    set cn = nothing
%>
    <select id="CategoryLookup" style="display:none">
    <%=CategoryLookupList%>
    </select>

    <div style="display:none">
    <form id=frmPreview method=post target=_blank action="Preview.asp"> 
    <textarea id="PreviewSummary" Name="PreviewSummary" cols="20" rows="2"></textarea>
    <textarea id="PreviewDetails" name="PreviewDetails" cols="20" rows="2"></textarea>
    <textarea id="PreviewURL" name="PreviewURL" cols="20" rows="2"></textarea>
    </form>
    </div>
</BODY>
</HTML>





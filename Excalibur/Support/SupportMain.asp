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
    <TITLE>Mobile Tools Support</TITLE>
<STYLE>
h3
    {
        font-family: Verdana;
        font-size:x-small;
    }
td{
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

    
</STYLE>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../includes/client/json2.js"></script>
<script type="text/javascript" src="../includes/client/json_parse.js"></script>
<script language="javascript" type="text/javascript">

    function window_onload() {
        modalDialog.load();
    }


    $(function () {
    		var currentStep = $("#txtCurrentStep").val();
        if(currentStep == "1"){
            $("#txtSubject").focus();
        }
        else if(currentStep == "3") {
            $("#txtDetails").focus();
        }
            
        $("#cboProject").change(function(){
	        var i;
	        var selectedValue = $("#cboProject").val();
	        var OptionArray;
	        $("#cboCategory > option").remove();
	        $("#cboCategory").append(new Option('',''));
	        $("#CategoryLookup > option").each(function(){
						if($(this).val() == selectedValue)      {
							OptionArray = $(this).text().split("|");
							//$("#cboCategory").append(new Option(OptionArray[1],OptionArray[0]));  \\This was is not working in IE compatability mode so see workaround.
							$("#cboCategory").append("<option value=\"" + OptionArray[0] + "\">" + OptionArray[1] + "</option>");
						}
	        });
        });
        
        $("#cboCategory").change(function(){
    			$("#CategoryChanged").val(1);
				});

    });

    function TitleMouseOver(){
        event.srcElement.style.color="red";
        event.srcElement.style.cursor="hand";
    }

    function TitleMouseOut(){
        event.srcElement.style.color="black";
        event.srcElement.style.cursor="default";
    }

    function TitleClick(ID){
        $("#SearchRow" + ID).toggle();
    }

    function TicketTitleClick(ID){
        $("#TicketRow" + ID).toggle();
    }

	function UploadZip(ID){
	    //save ID for return function: ---
	    globalVariable.save(ID, 'main_uploadzip_ID');

	    var url = "<%=AppRoot %>/PMR/SoftpaqFrame.asp?Title=Upload Support Attachments&Page=<%=AppRoot %>/common/fileupload.aspx";
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

	function RemoveUpload(ID){
		$("#UploadAddLinks" + ID).show();
		$("#UploadRemoveLinks" + ID).hide();
		$("#UploadPath" + ID).text("");
		$("#txtAttachmentPath" + ID).val("");
	}
</script>

</HEAD>


<body bgcolor="Ivory" onload="window_onload();">
<h3>Mobile Tools Support</h3>
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
    dim strDetails


    rs.open "spSupportProjectsListSelect",cn
    ProjectList = "<option></option>"
    do while not rs.EOF
        if trim(request("cboProject")) = trim(rs("ID")) then
            ProjectList = ProjectList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
        else
            ProjectList = ProjectList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
        end if
        rs.MoveNext
    loop
    rs.close

    CategoryLookupList=""
    rs.open "spSupportCategoryListSelect",cn
    CategoryLookupList = ""
    do while not rs.EOF
            CategoryLookupList = CategoryLookupList & "<option value=""" & rs("SupportProjectID") & """>" & rs("ID") & "|" & rs("Name") & "|" & rs("RequiredFields") & "</option>"
        rs.MoveNext
    loop
    rs.close


    CategoryList="<option></option>"
    if request("cboProject") <> "" then
        rs.open "spSupportCategoryListSelect " & clng(request("cboProject")),cn
        do while not rs.EOF
            if trim(request("cboCategory")) = trim(rs("ID")) then
                CategoryList = CategoryList & "<option selected value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
            else
                CategoryList = CategoryList & "<option value=""" & rs("ID") & """>" &  rs("Name") & "</option>"
            end if
            rs.MoveNext
        loop
        rs.close
    end if

    strRadio1Checked = ""
    strRadio2Checked = ""
    strRadio3Checked = ""
    strRadio4Checked = ""
    if trim(request("optType")) = "" or trim(request("optType")) = "0" then
        strRadio1Checked = " checked "
    end if
    if trim(request("optType")) = "1" then
        strRadio2Checked = " checked "
    end if
    if trim(request("optType")) = "2" then
        strRadio3Checked = " checked "
    end if
    if trim(request("optType")) = "3" then
        strRadio4Checked = " checked "
    end if

    strSearchResults = ""
    strRequirementsTemplate = ""
    

    if trim(request("txtSubject")) <> "" then
        
       if trim(request("optType")) = "1" then

	        set cm = server.CreateObject("ADODB.Command")

            cm.ActiveConnection = cn
            cm.CommandText = "spSupportSearchOpenKnownIssuesSelect"
            cm.CommandType = &H0004
       
	        Set p = cm.CreateParameter("@SearchText", adVarChar, &H0001,2000)
            if trim(left(replace(replace(request("txtSubject") & "","problem"," ",1,-1,1),"issue"," "),2000)) <> "" then
    	        p.Value = trim(left(replace(replace(request("txtSubject") & "","problem"," ",1,-1,1),"issue"," "),2000))
	        else
       	        p.Value = left(request("txtSubject") & "",2000)
            end if
            cm.Parameters.Append p

            if request("cboProject") <> "" and isnumeric(request("cboProject"))then
                Set p = cm.CreateParameter("@ProjectID", adInteger , &H0001)
                p.Value = cint(request("cboProject"))
                cm.Parameters.Append p
            end if

            set rs = cm.Execute
            if not (rs.eof and rs.bof) then
                strSearchResults = strSearchResults & "<table style=""width=100%;border: 1px solid tan"" cellpadding=2 cellspacing=0 border=1 bordercolor=tan>"
                strSearchResults = strSearchResults & "<tr bgcolor=""cornsilk""><td><b>Related Open Tickets</b></td></tr>"
            end if
            do while not rs.eof
                strDetails = replace(rs("Details") & "",chr(13),"<br>")
                if strDetails = "" then
                    strDetails = "No further information available."
                end if
                strSearchResults = strSearchResults & "<tr><td><div onmouseover=""TitleMouseOver()"" onmouseout=""TitleMouseOut()"" onclick=""TicketTitleClick(" & rs("ID") & ")"">" & rs("ID") & " - " & rs("Summary") & "</div><div style=""display:"" id=""TicketRow" & trim(rs("ID")) & """><fieldset>" & strDetails & "</fieldset></div></td></tr>"
                rs.MoveNext
            loop
            if not (rs.eof and rs.bof) then
                strSearchResults = strSearchResults & "</table>"
            end if
            rs.Close
            set cm=nothing
        end if

	    set cm = server.CreateObject("ADODB.Command")

        cm.ActiveConnection = cn
        cm.CommandText = "spSupportSearchSelect"
        cm.CommandType = &H0004
       
	    Set p = cm.CreateParameter("@SearchText", adVarChar, &H0001,2000)
        if trim(left(replace(replace(request("txtSubject") & "","problem"," ",1,-1,1),"issue"," ",1,-1,1),2000)) <> "" then
    	    p.Value = trim(left(replace(replace(request("txtSubject") & "","problem"," ",1,-1,1),"issue"," ",1,-1,1),2000))
	    else
    	    p.Value = left(request("txtSubject") & "",2000)
        end if
        cm.Parameters.Append p

        if request("cboProject") <> "" and isnumeric(request("cboProject"))then
            Set p = cm.CreateParameter("@ProjectID", adInteger, &H0001)
            p.Value = cint(request("cboProject"))
            cm.Parameters.Append p
        end if

        set rs = cm.Execute
        if not (rs.eof and rs.bof) then
            if strSearchResults <> "" then
                strSearchResults = strSearchResults & "<BR>"
            end if
            strSearchResults = strSearchResults & "<table style=""width=100%;border: 1px solid tan"" cellpadding=2 cellspacing=0 border=1 bordercolor=tan>"
            strSearchResults = strSearchResults & "<tr bgcolor=""cornsilk""><td><b>Related Knowledgebase Articles</b></td></tr>"
        end if
        do while not rs.eof
            strDetails = replace(trim(rs("ArticleText") & ""),chr(13),"<br>")
            if strDetails = "" and trim(rs("ArticleURL")) = "" then
                strDetails = "No further information available."
            end if
            if strDetails = "" and trim(rs("ArticleURL")) <> "" then
                strDetails = "<a target=_blank href=""" & rs("ArticleURL") & """>More Information</a>"
            elseif trim(rs("ArticleURL")) <> "" then
                strDetails = strDetails & "<BR><BR><a target=_blank href=""" & rs("ArticleURL") & """>More Information</a>"
            end if

            strSearchResults = strSearchResults & "<tr><td><div  onmouseover=""TitleMouseOver()"" onmouseout=""TitleMouseOut()"" onclick=""TitleClick(" & rs("ID") & ")"">" & rs("Title") & "</div><div style=""display:"" id=""SearchRow" & trim(rs("ID")) & """><fieldset>" & replace(strDetails,chr(13),"<BR>") & "</fieldset></div></td></tr>"
            rs.MoveNext
        loop
        if not (rs.eof and rs.bof) then
            strSearchResults = strSearchResults & "</table>"
        end if
        rs.Close
        set cm=nothing

        rs.open "spSupportRequiredFieldsSelect " & clng(request("cboCategory")) ,cn
        if not (rs.EOF and rs.bof) then
            strRequirementsTemplate = trim(rs("RequiredFields") & "")
        end if
        rs.close

    end if

    if trim(strRequirementsTemplate) = "" then
        strDisplayRequired = "none"
    else
        strDisplayRequired = ""
    end if
    if trim(request("CategoryChanged")) = "1" then
        strRequirementsField = trim(strRequirementsTemplate)
    else
        strRequirementsField = trim(request("txtRequired")) 
    end if

    if trim(strSearchResults) <> "" then
        strResultsFound = "1"
    else
        strResultsFound = "0"
    end if


    if trim(request("txtCurrentStep")) = "2" and trim(strSearchResults) <> "" then
        CurrentStep = 2
        strDisplayTab1 = "none"
        strDisplayTab2 = ""
        strDisplayTab3 = "none"
        strDisplayTab4 = "none"
    elseif trim(request("txtCurrentStep")) = "3" or (trim(request("txtCurrentStep")) = "2" and trim(strSearchResults) = "") then
        CurrentStep = 3
        strDisplayTab1 = "none"
        strDisplayTab2 = "none"
        strDisplayTab3 = ""
        strDisplayTab4 = "none"
    elseif trim(request("txtCurrentStep")) = "4" then
        CurrentStep = 4
        strDisplayTab1 = "none"
        strDisplayTab2 = "none"
        strDisplayTab3 = "none"
        strDisplayTab4 = ""
    else
        CurrentStep = 1
        strDisplayTab1 = ""
        strDisplayTab2 = "none"
        strDisplayTab3 = "none"
        strDisplayTab4 = "none"
    end if



%>
    <form id="frmMain" method="post" action="SupportMain.asp">
    <div id="Tab1" style="display:<%=strDisplayTab1%>">
    Please enter your question or request and select a product.<br /><br />
    <table cellpadding=2 border=1 cellspacing=0 bordercolor=tan bgcolor="cornsilk" style="border-width:1px;width:100%">
        <tr>
            <td><b>Question/Request:</b>&nbsp;<font color=red>*</font></td>
            <td colspan=3><input maxlength="500" id="txtSubject" name="txtSubject" type="text" onkeypress="return event.keyCode!=13" style="width:100%" value="<%= server.htmlencode(request("txtSubject"))%>" /></td>
        </tr>
        <tr>
            <td><b>Project:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboProject" name="cboProject" style="width:100%">
                    <%=ProjectList%>
                </select></td>
            <td><b>Feature/Issue:</b>&nbsp;<font color=red>*</font></td>
            <td width="50%">
                <select id="cboCategory" name="cboCategory" style="width:100%">
                    <%=CategoryList%>
                </select>
            </td>
        </tr>
        <tr>
            <td valign="top"><b>Request&nbsp;Type:</b>&nbsp;<font color=red>*</font></td>
            <td colspan="3">
                <input id="Radio1" <%= strRadio1Checked%> name="optType" type="radio" value="0" />Ask a Question<br />
                <input id="Radio2" <%= strRadio2Checked%> name="optType" type="radio" value="1" />Report an Issue<br />
                <input id="Radio3" <%= strRadio3Checked%> name="optType" type="radio" value="2"   />Make a Suggestion<br />
                <input id="Radio4" <%= strRadio4Checked%> name="optType" type="radio"  value="3" />Request Admin Updates (Update deliverable workflows, etc.)</td>
        </tr>
    </table>
    </div>

    <div id="Tab2" style="display:<%=strDisplayTab2%>">
        <input id="tagSearchedOn" type="hidden" value="<%= server.htmlencode(request("txtSubject"))%>" />
        <input id="CategoryChanged" name="CategoryChanged" type="hidden" value="0" />
        Please review the related articles and tickets listed below before continuing<br /><br />
        <%=strSearchResults%>
    </div>


    <div id="Tab3" style="display:<%=strDisplayTab3%>">
        Enter additional details and add attachments if needed.<br><br>
        <table cellpadding=2 border=1 cellspacing=0 bordercolor="tan" bgcolor="cornsilk" style="border-width:1px;width:100%">
            <tr>
                <td valign=top><b>Details:</b></td>
                <td colspan=3 width="100%">
                    <textarea id="txtDetails" style="width:100%;font-family:verdana;font-size:xx-small" name="txtDetails" rows=6><%=server.htmlEncode(request("txtDetails"))%></textarea>
                </td>
            </tr>
            <tr style="display:<%=strDisplayRequired%>">
                <td valign=top><b>Required&nbsp;Info:</b>&nbsp;<font color=red>*</font></td>
                <td colspan=3 width="100%">
                    <textarea id="txtRequired" style="width:100%;font-family:verdana;font-size:xx-small" name="txtRequired" rows=6><%=server.htmlEncode(strRequirementsField)%></textarea>
                    <textarea id="txtRequiredTemplate" style="display:none;font-family:verdana;font-size:xx-small" name="txtRequiredTemplate" rows=6><%=server.htmlEncode(strRequirementsTemplate)%></textarea>
                </td>
            </tr>
                    <tr>
                        <td valign="top"><b>Attachment&nbsp;1:</b></td>
                        <td valign="top">
                            <%if request("txtAttachmentPath1") = "" then %>
                                <div id="UploadAddLinks1"><a href="javascript: UploadZip(1);">Upload</a></div>
                                <div id="UploadRemoveLinks1" style="display:none"><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath1></label></div>
                            <%else
                                AttachmentArray = split(request("txtAttachmentPath1"),"\")
                            %>
                                <div id="UploadAddLinks1" style="display:none"><a href="javascript: UploadZip(1);">Upload</a></div>
                                <div id="UploadRemoveLinks1"><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath1><%=AttachmentArray(ubound(AttachmentArray)) %></label></div>
                            <%end if%>
                            <input id="txtAttachmentPath1" name="txtAttachmentPath1" type="hidden" value="<%=server.htmlencode(request("txtAttachmentPath1"))%>" />
                        </td>
                    </tr>
                    <tr>
                        <td valign="top"><b>Attachment&nbsp;2:</b></td>
                        <td valign="top">
                            <%if request("txtAttachmentPath2") = "" then %>
                            <div id="UploadAddLinks2"><a href="javascript: UploadZip(2);">Upload</a></div>
                            <div id="UploadRemoveLinks2" style="display:none"><a href="javascript: UploadZip(2);">Change</a> | <a href="javascript: RemoveUpload(2);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath2></label></div>
                            <%else
                                AttachmentArray = split(request("txtAttachmentPath2"),"\")
                            %>
                            <div id="UploadAddLinks2" style="display:none"><a href="javascript: UploadZip(2);">Upload</a></div>
                            <div id="UploadRemoveLinks2"><a href="javascript: UploadZip(2);">Change</a> | <a href="javascript: RemoveUpload(2);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath2><%=AttachmentArray(ubound(AttachmentArray)) %></label></div>
                            <%end if%>
                            <input id="txtAttachmentPath2" name="txtAttachmentPath2" type="hidden" value="<%=server.htmlencode(request("txtAttachmentPath2"))%>" />
                        </td>
                    </tr>
                    <tr>
                        <td valign="top"><b>Attachment&nbsp;3:</b></td>
                        <td valign="top">
                            <%if request("txtAttachmentPath3") = "" then %>
                                <div id="UploadAddLinks3"><a href="javascript: UploadZip(3);">Upload</a></div>
                                <div id="UploadRemoveLinks3" style="display:none"><a href="javascript: UploadZip(3);">Change</a> | <a href="javascript: RemoveUpload(3);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath3></label></div>
                            <%else
                                AttachmentArray = split(request("txtAttachmentPath3"),"\")
                            %>
                                <div id="UploadAddLinks3" style="display:none"><a href="javascript: UploadZip(3);">Upload</a></div>
                                <div id="UploadRemoveLinks3"><a href="javascript: UploadZip(3);">Change</a> | <a href="javascript: RemoveUpload(3);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id=UploadPath3><%=AttachmentArray(ubound(AttachmentArray)) %></label></div>
                            <%end if%>
                            <input id="txtAttachmentPath3" name="txtAttachmentPath3" type="hidden" value="<%=server.htmlencode(request("txtAttachmentPath3"))%>" />
                        </td>
                    </tr>
        </table>
    </div>

    <div id="Tab4" style="display:<%=strDisplayTab4%>">
        Please review this information and click the Finish button to notify the support team.<br><br>
        <input id="chkCopyMe" name="chkCopyMe" type="checkbox"/> Copy me on this email.<br><br>
        <div id="EmailText" style="padding:3px 3px 3px 3px; background-color:white;border:solid 1px gray;width:100%;height:200px"></div>
    </div>
    <input id="txtCurrentStep" name="txtCurrentStep" type="hidden" value="<%=CurrentStep%>" />
    <input id="txtResultsFound" name="txtResultsFound" type="hidden" value="<%=strResultsFound%>" />
    <input id="txtProjectName" name="txtProjectName" type="hidden" value="" />
    <input id="txtCategoryName" name="txtCategoryName" type="hidden" value="" />
    

    </form>

    <select id="CategoryLookup" style="display:none">
    <%=CategoryLookupList%>
    </select>
<%

    set rs = nothing
    cn.Close
    set cn = nothing
%>
<div style="display:none">
    <form id="frmCancel" method="post" action="SupportCancel.asp">
        <input id="txtCancelSummary" name="txtCancelSummary" type="hidden" value="" />
        <input id="txtCancelStep" name="txtCancelStep" type="hidden" value="<%=CurrentStep%>" />
        <input id="txtCancelProject" name="txtCancelProject" type="hidden" value="" />
        <input id="txtCancelCategory" name="txtCancelCategory" type="hidden" value="" />
    </form>
</div>
</body>
</HTML>





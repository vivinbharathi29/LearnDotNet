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

    
</STYLE>

<script language="javascript">

</script>

</HEAD>


<BODY>
<h3>Mobile Tool Support</h3>
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
        dim strAttachment1
        dim strAttachment2
        dim strAttachment3
        dim strPartnerID
        dim strShowResponseRequired
        
        dim strProjectName
        dim strCategoryName
        dim strTypeName
        dim strStatusName
        dim strOwnerName

        strSummary = rs("Summary") & ""
        strDetails = rs("Details") & ""
        strResolution = rs("Resolution") & ""
        strProjectID = rs("ProjectID") & ""
        strProjectName = rs("ProjectName") & ""
        strCategoryID = rs("CategoryID") & ""
        strSubmitterName = rs("SubmitterName") & ""
        strSubmitterEmail = rs("SubmitterEmail") & ""
        strTypeID = rs("TypeID") & ""
        strStatusID = rs("StatusID") & ""
        strOwnerID = rs("OwnerID") & ""
        strOwnerName = rs("OwnerName") & ""
        strAttachment1 = rs("Attachment1") & ""
        strAttachment2 = rs("Attachment2") & ""
        strAttachment3 = rs("Attachment3") & ""
        strPartnerID = rs("SubmitterPartnerID") & ""
        rs.Close

        if clng(strStatusID) = 2 then
            strShowResponseRequired = ""
        else
            strShowResponseRequired="none"
        end if 


        strTypeName=""
        rs.open "spSupportIssueTypeSelect", cn
        do while not rs.EOF
            if trim(strTypeID) = trim(rs("ID") & "") then
                strTypeName = rs("Name") & "&nbsp;"
            end if
            rs.MoveNext
        loop
        rs.Close

        rs.open "spSupportIssueStatusSelect", cn
        strStatusName = ""
        do while not rs.EOF
            if trim(strStatusID) = trim(rs("ID") & "") then
                strStatusName =  rs("Name") & "&nbsp;"
            end if
            rs.MoveNext
        loop
        rs.Close

        strCategoryName=""
        if trim(strProjectID) <> "" then
            rs.open "spSupportCategoryListSelect " & clng(strProjectID),cn
            do while not rs.EOF
                if trim(strCategoryID) = trim(rs("ID")) then
                    strCategoryName= rs("Name") & "&nbsp;"
                    exit do
                end if
                rs.MoveNext
            loop
            rs.close
        end if



%>
    <b>Ticket #:</b> <%=strTicketNumber%><br>
    <table cellpadding=2 border=1 cellspacing=0 bordercolor=lightgray bgcolor="ivory" style="border-width:1px;width:100%">
        <tr>
            <td><b>Question/Request:</b>&nbsp;</td>
            <td colspan=3><%= server.htmlencode(strSummary)%></td>
        </tr>
        <tr>
            <td><b>Project:</b>&nbsp;</td>
            <td width="50%"><%=server.htmlencode(strProjectName)%></td>
            <td><b>Submitter:</b>&nbsp;</td>
            <td width="50%"><%=strSubmitterName%>
            </td>
        </tr>
        <tr>
            <td><b>Category:</b>&nbsp;</td>
            <td width="50%"><%=strCategoryName%></td>
            <td><b>Owner:</b>&nbsp;</td>
            <td width="50%"><%=strOwnerName%>&nbsp;</td>
        </tr>
        <tr>
            <td><b>Request Type:</b>&nbsp;</td>
            <td width="50%"><%=strTypeName%>&nbsp;</td>
            <td><b>Status:</b>&nbsp;</td>
            <td width="50%"><%=strStatusName%></td>
        </tr>
        <tr>
            <td valign=top><b>Response:</b></td>
            <td colspan=3 width="100%">
                <%=server.htmlEncode(strResolution)%>&nbsp;
            </td>
        </tr>
        <tr>
            <td valign=top><b>Working&nbsp;Notes:</b></td>
            <td colspan=3 width="100%">
                <%=server.htmlEncode(strDetails)%>&nbsp;
            </td>
        </tr>
        <%if trim(strAttachment1) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;1:</b></td>
            <td valign=top><%=strAttachment1%></td>
        </tr>
        <%if trim(strAttachment2) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;2:</b></td>
            <td valign=top><%=strAttachment2%></td>
        </tr>
        <%if trim(strAttachment3) <> "" then%>
        <tr>
        <%else %>
        <tr style="display:none">
        <%end if%>
            <td valign=top><b>Attachment&nbsp;3:</b></td>
            <td valign=top><%=strAttachment3%></td>
        </tr>
    </table>

    <%end if

    set rs = nothing
    cn.Close
    set cn = nothing
%>
</BODY>
</HTML>





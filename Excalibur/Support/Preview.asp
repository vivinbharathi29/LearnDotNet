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
    function window_onload(){
        if (txtType.value == "3")
            window.resizeTo(700, 400); 
    }
//-->
</SCRIPT>

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
</HEAD>



<BODY onload="window_onload();">
<%
	dim cn, rs, cm
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    dim blnFound
    dim strDetails
    dim strTitle
    dim strOwner

    blnFound = false

    if trim(request("ID")) <> "" then
        rs.open "spSupportArticleSelect " & clng(request("ID")),cn
        if (rs.eof and rs.bof) then
            response.write "Unable to find the requested article"
        else
            strTitle = rs("Title") & ""
            strDetails = formatdetails(rs("ArticleText") & "",rs("ArticleURL") & "")
            response.write "<font face=verdana size=2><b>" & rs("ProjectName") & " Support</b><BR><br></font>"
            response.write "<font face=verdana size=1 color=gray>Article ID: " & rs("ID") & " - Last Updated: " & formatdatetime(rs("LastUpdated"),vbshortdate) & "<BR><BR></font>"
            strOwner =  "<font face=verdana size=1 color=black>Article Owner: <a href=""mailto:" & rs("OwnerEmail") & """>" & rs("OwnerName") & "</a></font>"
        %>

            <table width="650px" bgcolor=ivory style="border: solid 1px tan">
                <tr>
                    <td>
                        <%=strTitle%>
                        <fieldset>
                        <%=strDetails%>
                        </fieldset>
                        <%=strOwner%>
                    </td>
                </tr>
            </table>




        <%
        end if
        rs.Close
    elseif request("List") <> "" then
        rs.open "spSupportArticleSelect",cn
        if (rs.eof and rs.bof) then
            response.write "Unable to find the requested article"
        else
            response.write "<font size=1 face=verdana>Select all related articles.<br><form id=""frmMain"" method=""post"" action=""ArticleListSave.asp""></font>"
            dim strLastProject
            dim strLastCategory
            do while not rs.EOF
                if trim(rs("ProjectName") & "") <> strLastProject then
                    if strLastProject <> "" then
                        response.write "</table></td></tr>"
                        response.write "</table>"
                        response.write "<BR><BR>"
                    end if
                    response.write "<b><font size=2 face=verdana>" & rs("ProjectName") & "</font></b><BR>"
                    response.write "<table cellpadding=2 cellspacing=0 style=""border: solid 1px gainsboro;width:100%"">"
                    strLastCategory = ""
                end if
                strLastProject = trim(rs("ProjectName") & "") 

                if trim(rs("CategoryName") & "") <> strLastCategory then
                    if strLastCategory <> "" then
                        response.write "</table></td></tr>"
                    end if
                    response.write "<tr bgcolor=beige><td><b>" & rs("CategoryName") & "</b></td></tr><tr><td><table cellpadding=2 cellspacing=0>"
                end if
                strLastCategory = trim(rs("CategoryName") & "") 


                response.write "<tr bgcolor=ivory><td valign=top><input id=""chkArticle"" name=""chkArticle"" type=""checkbox""  value=""" & rs("ID") & """ /></td><td>" & rs("Title") & "</td></tr>"
                rs.MoveNext
            loop
            response.write "</form>"
        end if
        rs.Close
    else
        strDetails = FormatDetails(request("PreviewDetails"),request("PreviewURL"))
        %>
            <table width="650px" bgcolor=ivory style="border: solid 1px tan">
                <tr>
                    <td>
                        <%=request("PreviewSummary")%>
                        <fieldset>
                        <%=strDetails%>
                        </fieldset>
                    </td>
                </tr>
            </table>
        <%
    end if


    function FormatDetails(strDetails, strURL)
        dim strOutput
        strOutput = strDetails
        if strOutput = "" and trim(strURL) = "" then
            strOutput = "No further information available."
        end if
        if strOutput = "" and trim(strURL) <> "" then
            strOutput = "<a target=_blank href=""" & strURL & """>More Information</a>"
        elseif trim(strURL) <> "" then
            strOutput = strOutput & "<BR><BR><a target=_blank href=""" & strURL & """>More Information</a>"
        end if
        strOutput = replace(strOutput,chr(13),"<BR>")

        FormatDetails = strOutput
    end function


    dim strType

    if request("ID") <> "" then
        strType = "1"
    elseif request("List") <> "" then
        strType = "2"
    else
        strType = "3"
    end if
%>

    <input id="txtType" type="hidden" value="<%=strType%>"/>

<%

    set rs = nothing
    cn.Close
    set cn = nothing
%>

</BODY>
</HTML>





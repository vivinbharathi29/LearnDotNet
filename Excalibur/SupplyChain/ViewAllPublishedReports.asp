<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file="../includes/lib_debug.inc" -->
<%
    Dim strError, strTabName
    Dim nProductBrandID, nProductVersionID
    Dim Security, m_UserFullName
    Dim cn, dw, cmd, rs, cnString 

    dim AppRoot
    AppRoot = Session("ApplicationRoot")

    Set Security = New ExcaliburSecurity
    m_UserFullName = Security.CurrentUserFullName()

    strError = ""

    if Request.QueryString("ProductBrandID") <> "" then
	    nProductBrandID = clng(Request.QueryString("ProductBrandID"))
    end if
    if Request.QueryString("PVID") <> "" then
	    nProductVersionID = clng(Request.QueryString("PVID"))
    end if

    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_ViewAllSCMPublishes")
    dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, nProductVersionID
    dw.CreateParameter cmd, "@p_intProductBrandID", adInteger, adParamInput, 8, nProductBrandID

    %>

    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
    <title>Published SCM Reports</title>    
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />        
    </head>
    
    <body> 
        <div style="font-size:12px;width:90%">
            <h2>SCM/Program Matrix Reports</h2>
        </div>
        <div style="font-size:10px;margin-top:5px;margin-bottom:15px;width:90%">
            SCM Reports are created in Excel format. Microsoft Office 2007 or higher is needed to run the reports.
        </div>      
    <%
    
	Set rs = dw.ExecuteCommandReturnRS(cmd)

    If rs.EOF Then
    %>
        <table id="tblNoReport" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <font face="Verdana" size="2">No reports found.</font>
                </td>
            </tr>
        </table>
    <%
    Else
    %>
       <div id="GridViewContainer" class="GridViewContainer" style="width: 100%;">
        <table id="tblPublishedReport" class="Table" width="100%">
            <col width="95" align="center">
            <col width="220" align="center">
            <col width="50" align="center">
            <col width="100" align="center">
            <col width="150" align="center">
            <thead>
                <tr class="FrozenHeader">
                    <th style="background-color: wheat">
                        Date Published
                    </th>
                    <th style="background-color: wheat">
                        SCM Name
                    </th>
                    <th style="background-color: wheat">
                       Revision
                    </th>
                    <th style="background-color: wheat">
                        Created By
                    </th>
                    <th style="background-color: wheat">
                        Reason
                    </th>
                </tr>
            </thead>        
    <% 
    End If
            
    Do until rs.EOF
    %>
        <tr>
            <% if (rs("StandardRelease") = "2") then %>
                <td ><%=rs("Created") %></td>
            <% else %>
                <td ><a href="/ipulsar/Reports/SCM/SCM_Report_DT.aspx?SCMID=<%=rs("SCMID")%>&PVID=<%=nProductVersionID%>&BID=<%=nProductBrandID %>"><%=rs("Created") %></a></td>
            <% end if %>
            <td style="text-align:left; padding-left:3px" > <%=rs("SCMName") %></td>
            <td > <%=rs("Revision") %></td>
            <td > <%=rs("CreatedBy") %></td>
            <td style="text-align:left;padding-left:3px"> <%= Server.HtmlEncode(Replace(rs("Reason"),"'","\'")) %></td>
        </tr>
    <%
       rs.MoveNext
    Loop       
    %></table></div></body></html><%
    Set rs= nothing
    Set cn= nothing
    %>
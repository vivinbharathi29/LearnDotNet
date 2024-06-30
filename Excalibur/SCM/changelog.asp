<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%

Dim rs, dw, cn, cmd
Dim IfeatureCategoryID, sFeatureCategory
Dim bFirstWrite
Dim bIsPc : bIsPc = False
Dim m_BrandID : m_BrandID = ""
Dim m_BrandName : m_BrandName = ""
on error resume next
Dim strTitleColor : strTitleColor = Request.Cookies("TitleColor")
if strTitleColor = "" then 
	strTitleColor = "#0000cd"
end if
on error goto 0
Dim bAdministrator : bAdministrator = false
Dim CurrentUser     : CurrentUser = lcase(Session("LoggedInUser"))
Dim bShowAll	    : bShowAll = Request.QueryString("ShowAll") : If bShowAll = "" Then bShowAll = False
Dim bShowSelected   : bShowSelected = Request.QueryString("ShowSelected") : If bShowSelected = "" Then bShowSelected = False
Dim bShowSCM        : bShowSCM = Request.QueryString("ShowSCM") : If bShowSCM = "" Then bShowSCM = True
Dim bShowPM         : bShowPM = Request.QueryString("ShowPM") : If bShowPM = "" Then bShowPM = False
Dim CurrentDomain
Dim CurrentUserPartner
Dim CurrentUserName
Dim CurrentUserID
Dim CurrentUserSysAdmin
Dim CurrentWorkgroupID
Dim bPreinstallPM
Dim bCommodityPM
Dim sFavs
Dim sFavCount
Dim sProductName
Dim sDisplayedProductName
Dim SEPMID
Dim PMID
Dim PCID
Dim RowClass
Dim sFilterQueryString
Dim sShowAllQueryString
Dim sSelectAll

sFilterQueryString = Request.QueryString
If InStr(sFilterQueryString, "ShowAll") Then
	sFilterQueryString = Replace(sFilterQueryString, "ShowAll=True", "ShowAll=False")
Else
	sFilterQueryString = sFilterQueryString & "&ShowAll=False"
End If
sShowAllQueryString = Replace(sFilterQueryString, "ShowAll=False", "ShowAll=True")

If instr(CurrentUser,"\") > 0 Then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If

Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_IsMarketingUser
Dim m_EditModeOn
Dim m_UserFullName
Dim m_ProductVersionID : m_ProductVersionID = Request("ID")

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	bIsPc = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Or m_IsMarketingUser Then
		bIsPc = True
	End If
	
	Set Security = Nothing
'##############################################################################	

'
' Setup the data connections
'
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommAndSP(cn, "spGetUserInfo")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, CurrentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
Set rs = dw.ExecuteCommAndReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then
	CurrentUserName = rs("Name") & ""
	CurrentUserID = rs("ID") & ""
	CurrentUserSysAdmin = rs("SystemAdmin")
	CurrentWorkgroupID = rs("WorkgroupID") & ""
	CurrentUserPartner = trim(rs("PartnerID") & "")
	bPreinstallPM = rs("PreinstallPM")
	bCommodityPM = rs("CommodityPM")
	
	sFavs = trim(rs("Favorites") & "")
	sFavCount = trim(rs("FavCount") & "")
End If
rs.Close

dim ShowItem
If CurrentUserPartner = "1" Then
	ShowItem = ""
Else
	ShowItem = "none"
End If

Set cmd = dw.CreateCommAndSP(cn, "spGetProductVersion")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("ID")
Set rs = dw.ExecuteCommAndReturnRS(cmd)

If (rs.EOF And rs.BOF) And Request("ID") <> "-1" Then
	Response.Write "Unable to find the selected program.<br /><font size=1>ID=" & request("ID") & "</font>"
	Response.Write "<br /><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & request("ID") & ")""><font face=verdana size=1>Remove From Favorites</font></a>"
	Response.Write "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
	Response.Write "<span id=EditLink style=""Display:none""></span><span id=StatusLink style=""Display:none""></span><span id=menubar style=""Display:none""></span><span ID=Wait style=""Display:none""></span>"
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""1"">"
Else
	Response.Write "<INPUT type=""hidden"" id=txtError name=txtError value=""0"">"
	sProductName = rs("Name") & " " & rs("Version") 
	sDisplayedProductName = rs("Name") & " " & rs("Version")
	SEPMID = rs("sepmid")
	PMID = rs("PMID")
End If				
rs.Close

Function PrepForWeb( value )

	If Trim( value ) = "" Or IsNull(value) Then
		PrepForWeb = "&nbsp;"
	Else
		PrepForWeb = Replace(Server.HTMLEncode( value ), vbCrLf, "<br />")
	End If

End Function

'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
If CurrentUserSysAdmin or SEPMID = CurrentUSerID or instr(trim(PMID),"_" & trim(CurrentUSerID) & "_") > 0  Then
	bAdministrator = true
End If

If Request.Form("hidIsPostback") = "true" And bIsPc Then
    Dim item
    Dim BitValue
    for each item in request.Form
        If Left(item, 3) = "cbx" then
            'response.Write item & "=" & request.Form(item) & " " & "hidShowOnScm" & mid(item, 4, len(item)-3) & "=" & Request.Form("hidShowOnScm" & mid(item, 4, len(item)-3)) & "<br />"
            If Request.Form(item) <> Request.Form("hidShowOnScm" & mid(item, 4, len(item)-3)) Then
                'response.Write "Newly Checked Item " & item & "<br />"
                Call UpdateShowOnScmBit(mid(item, 4, len(item)-3), 1)
            End If
        ElseIf Left(item, 3) = "cby" then
            If Request.Form(item) <> Request.Form("hidShowOnPM" & mid(item, 4, len(item)-3)) Then
                Call UpdateShowOnPMBit(mid(item, 4, len(item)-3), 1)
            End If
        End if
        If Left(item, 12) = "hidShowOnScm" then
            'response.Write item & "=" & request.Form(item) & " " & "cbx" & mid(item, 13, len(item)-12) & "=" & Request.Form("cbx" & mid(item, 13, len(item)-12)) & "<br />"
            If Request.Form(item) <> Request.Form("cbx" & mid(item, 13, len(item)-12)) Then
                'response.Write "Newly Unchecked Item " & item & "<br />"
                Call UpdateShowOnScmBit(mid(item, 13, len(item)-12), 0)
            End If
        ElseIf Left(item, 11) = "hidShowOnPM" then
            If Request.Form(item) <> Request.Form("cby" & mid(item, 12, len(item)-11)) Then
                Call UpdateShowOnPMBit(mid(item, 12, len(item)-11), 0)
            End If
        End if

    next

End If

Sub UpdateShowOnScmBit(HistoryID, ShowOnScm)
    
    Set cmd = dw.CreateCommandSP(cn, "usp_SetAvHistoryShowOnScmStatus")
	dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, HistoryID
	dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, ShowOnScm
	dw.Createparameter cmd, "@p_ShowOnPM", adBoolean, adParamInput, 1, NULL
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
	dw.ExecuteNonQuery(cmd)
    Set cmd = nothing
    
End Sub

Sub UpdateShowOnPMBit(HistoryID, ShowOnPM)
    
    Set cmd = dw.CreateCommandSP(cn, "usp_SetAvHistoryShowOnScmStatus")
	dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, HistoryID
	dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, NULL
	dw.Createparameter cmd, "@p_ShowOnPM", adBoolean, adParamInput, 1, ShowOnPM
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
	dw.ExecuteNonQuery(cmd)
    Set cmd = nothing
    
End Sub


%>
<html>
<head>
    <title>SCM Change Log</title>
    <link href="../style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="../style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="style.css" />
</head>

<script type="text/javascript">
function BrandLink_onClick(ProductBrandID)
{
	window.location.replace("changelog.asp?ID=<%=Request("ID")%>&Class=<%=Request("Class")%>&BID=" + ProductBrandID);
}

function Row_OnMouseOver()
{
	var node = window.event.srcElement;
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}
	
	node.style.color = "red";
	node.style.cursor = "hand";
}

function Row_OnMouseOut() {
	var node = window.event.srcElement;
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}

   	node.style.color = "black";
}

function Row_OnClick()
{
	var node = window.event.srcElement;
	
	if (node.type == "checkbox")
	    return;
	
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}
	
	var strID;
	strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=edit&PVID=" + node.pvid + "&CLID=" + node.clid,"","dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	document.location.reload();
}

function AddEntry( ProductVersionID, ProductBrandID )
{
	var strID;
	strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=add&PVID=" + ProductVersionID + "&PBID=" + ProductBrandID,"","dialogWidth:500px;dialogHeight:275px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	document.location.reload();
}
</script>

<body>
    <form id="frmMain" method="post">
        <font size="4"><strong>
            <%= sProductName%>
            SCM Change Log</strong></font><br />
        <br />
        <%
'
' Get the list of Brands for the product.
'
Set cmd = dw.CreateCommAndSP(cn, "spListBrands4Product")
dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 8, Request("ID")
dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
Set rs = dw.ExecuteCommAndReturnRS(cmd)
	
bFirstWrite = True

If Not rs.EOF Then
	'm_BrandID = rs("ProductBrandID")			
	'm_BrandName = rs("Name")
			
        %>
        <br />
        <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
            <tr>
                <td valign="top">
                    <table>
                        <tr>
                            <td valign="top">
                                <font color="navy" face="verdana" size="2"><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td>
                        </tr>
                    </table>
                    <td width="100%">
                        <table>
                            <tr>
                                <td>
                                    <b>Brand:</b></td>
                                <td width="100%">
                                    <%			
	Do Until rs.EOF
		'Response.Write "<td><a href=""javascript:void(0)"">" & server.HTMLEncode(rs("schedule_name")) & "</a></td>"
		If Not bFirstWrite Then
			Response.Write "&nbsp;|&nbsp;"
		End If
			
		If (Request("BID") = "" And m_BrandID = "") Or (CLng(rs("ProductBrandID")) = CLng(Request("BID"))) Then
			m_BrandID = rs("ProductBrandID")			
			m_BrandName = rs("Name")
			Response.Write server.HTMLEncode(m_BrandName)
		Else
			Response.Write "<a href=""javascript:BrandLink_onClick(" & rs("ProductBrandID") & ")"">" & server.HTMLEncode(rs("Name")) & "</a>"
		End If
		bFirstWrite = False
		rs.MoveNext
	Loop
	Response.Write "</td></tr></table></td></tr></table><br />"
Else
    Response.Write "<br /><h2><font color='red'>No Brand Information provided unable to create report</font></h2>"
    response.End

End If
rs.Close
set cmd = nothing

'##########################################################
'#
'# Draw Menu
'#
'##########################################################

                                    %>
                                    <button type="submit">
                                        Save Changes</button>
                                    <br />
                                    <br />
                                    <font size="1" face="verdana">
                                        <br />
                                        <%If bIsPc Then%>
                                        <a href="#" onclick="AddEntry(<%=Request("ID")%>, <%=m_BrandID%>);">Add Entry</a>
                                        |
                                        <%End If%>
                                        <%If bShowAll Then%>
                                        <a href='ChangeLog.asp?<%= sFilterQueryString%>'>Filter List</a>
                                        <%Else%>
                                        <a href='ChangeLog.asp?<%= sShowAllQueryString%>'>Show All</a>
                                        <%End If%>
                                    </font>
                                    <br />
                                    <br />
                                    <font size="2"><b>
                                        <%= m_BrandName%>
                                        SCM Change Log:</b></font><span style="font-size:8pt; font-family:Verdana; color:red">&nbsp;&nbsp;&nbsp;-&nbsp;(click on the change to view full details)</span>
                                    <%


'##########################################################
'#
'# Draw Main Display
'#
'##########################################################

If bShowAll Or bShowSelected Then
    sSelectAll = 1
Else
    sSelectAll = 0
End If

Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvHistory")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, m_BrandID
dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, ""
dw.CreateParameter cmd, "@p_ShowOnSCM", adBoolean, adParamInput, 8, NULL
dw.CreateParameter cmd, "@p_ShowOnPM", adBoolean, adParamInput, 8, NULL
dw.CreateParameter cmd, "@p_ShowAll", adBoolean, adParamInput, 8, sSelectAll
dw.CreateParameter cmd, "@p_ShowDays", adInteger, adParamInput, 8, 90

Set rs = dw.ExecuteCommAndReturnRS(cmd)

If rs.EOF Then
                                    %>
                                    <table id="TableScm" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>
                                                <font face="Verdana" size="2">No changes found for this brand.</font>
                                            </td>
                                        </tr>
                                    </table>
                                    <%
Else
                                    %>
                                    <table id="TableSchedule" cellspacing="1" cellpadding="1" width="100%" border="1"
                                        bordercolor="tan" bgcolor="ivory">
                                        <col align="center" />
                                        <col align="center" />
                                        <col />
                                        <col />
                                        <col />
                                        <col />
                                        <col />
                                        <col align="center" />
                                        <col />
                                        <col />
                                        <col />
                                        <tr>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Show<br />
                                                On SCM</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Show<br />
                                                On PM</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Change Date</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Changed By</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                AV No.</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                GPG Desc.</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Field</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Change<br />
                                                Type</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Change From</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Change To</th>
                                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                                Comment / Reason</th>
                                        </tr>
                                        <%
rs.Sort = "Last_Upd_Date desc"

	Do Until rs.EOF
		
		If bShowAll Or rs("ShowOnScm") Or rs("ShowOnPM") Then
			If RowClass = "tablerow1" Then
				RowClass = "tablerow2"
			Else
				RowClass = "tablerow1"
			End If
                                        %>
                                        <tr bgcolor="cornsilk" pvid="<%= Request("ID")%>" clid="<%= rs("ID")%>" onmouseover="return Row_OnMouseOver()"
                                            onmouseout="return Row_OnMouseOut()" onclick="return Row_OnClick()">
                                            <td class="cell" style="white-space: nowrap">
                                                <input name="cbx<%=rs("ID") %>" type="checkbox" <% if rs("showonscm") then response.write "checked" end if %> />
                                                <input name="hidShowOnScm<%=rs("ID") %>" type="checkbox" <% if rs("showonscm") then response.write "checked" end if %>
                                                    style="display: none" />
                                            <td class="cell" style="white-space: nowrap">
                                                <input name="cby<%=rs("ID") %>" type="checkbox" <% if rs("showonpm") then response.write "checked" end if %> />
                                                <input name="hidShowOnPM<%=rs("ID") %>" type="checkbox" <% if rs("showonpm") then response.write "checked" end if %>
                                                    style="display: none" />
                                            </td>
                                            <td class="cell" style="white-space: nowrap; text-align: right">
                                                <%=FormatDateTime(PrepForWeb(rs("Last_Upd_Date")), vbShortDate)%>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <%
                                                If InStr(rs("Last_Upd_User"), ",") Then
                                                    Response.Write PrepForWeb(Left(rs("Last_Upd_User"), InStr(rs("Last_Upd_User"), ",") + 2))
                                                Else
                                                    Response.Write PrepForWeb(rs("Last_Upd_User"))
                                                End If
                                                 %>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <%=PrepForWeb(rs("AvNo"))%>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <%=PrepForWeb(rs("GPGDescription"))%>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <%=PrepForWeb(rs("ColumnChanged"))%>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <%=PrepForWeb(rs("AvChangeTypeCd"))%>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <% 
                                                If Len(rs("OldValue")) > 35 Then
                                                    Response.Write Left(PrepForWeb(rs("OldValue")),30) & " ..."
                                                Else
                                                    Response.Write PrepForWeb(rs("OldValue"))
                                                End If
                                                %>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <% 
                                                If Len(rs("NewValue")) > 35 Then
                                                    Response.Write Left(PrepForWeb(rs("NewValue")),30) & " ..."
                                                Else
                                                    Response.Write PrepForWeb(rs("NewValue"))
                                                End If
                                                %>
                                            </td>
                                            <td class="cell" style="white-space: nowrap">
                                                <% 
                                                If Len(rs("Comments")) > 35 Then
                                                    Response.Write Left(PrepForWeb(rs("Comments")),30) & " ..."
                                                Else
                                                    Response.Write PrepForWeb(rs("Comments"))
                                                End If
                                                %>
                                            </td>
                                        </tr>
                                        <%
		End If
		rs.MoveNext
	Loop
                                        %>
                                    </table>
                                    <%
End If
                                    %>
                                    <div id="PopUpMenu" class="hidden">
                                        <ul id="menu">
                                            <li><strong><a href="#" onclick="parent.location.href='javascript:MenuProperties();'"
                                                target="parent">Properties</a></strong></li>
                                            <li id="spacer">
                                                <hr width="95%">
                                            </li>
                                            <li><a href="#">Link</a></li>
                                            <li><a href="#">Clone</a></li>
                                            <li id="spacer">
                                                <hr width="95%">
                                            </li>
                                            <li><a href="#">Obselete</a></li>
                                        </ul>
                                    </div>
                                    <input type="hidden" name="hidIsPostback" value="true" />
    </form>
</body>
</html>

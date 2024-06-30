<%@  language="VBScript" %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file="../includes/lib_debug.inc" -->
<%
Dim m_ProductVersionID : m_ProductVersionID = Request("PVID")
Dim m_UserID : m_ProductVersionID = Request("UserID")
Dim m_Mode : m_Mode = Request("Mode")
Dim m_IsMarketingUser
Dim m_IsPC : m_IsPC = Request("IsPC")

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If

    Dim isAdmin, isPulsarSystemAdmin
    isAdmin = Security.IsSysAdmin()
	isPulsarSystemAdmin = Security.IsPulsarSystemAdmin()
	Set Security = Nothing
'##############################################################################	
Dim CurrentUser : CurrentUser = lcase(Session("LoggedInUser"))
Dim CurrentDomain
If instr(CurrentUser,"\") > 0 Then
    CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
    CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If
Dim rs, dw, cn, cmd
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommAndSP(cn, "spGetUserInfo")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, CurrentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then
	'add the permission from the Users and Roles to the Pulsar products
    If Not m_IsMarketingUser Then
		MarketingProductCount = rs("MarketingProductCount")
        if MarketingProductCount > 0 then
            m_IsMarketingUser = True
        end if
	End If
End If
rs.Close



%>
<html>
<head>
    <title>SCM AV Details</title>
    <script type="text/javascript">
        function OpenAddExistingSharedAV(url, width, height) {
            window.parent.OpenAddExistingSharedAV(url, width, height);
        }
        function ResetSharedAvOption() {
            if (window.frames["UpperWindow"].frmMain != null)
                window.frames["UpperWindow"].frmMain.rdNonSharedAv.checked = true;
        }
        function ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform) {
            if (window.frames["UpperWindow"].frmMain != null)
                window.frames["UpperWindow"].ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform);
        }
        function LoadPage() {
            <% 
            Dim mktURL : mktURL = "avMarketingDetail.asp?" & Request.QueryString
            Dim scURL : scURL = "avDetail.asp?" & Request.QueryString & "&IsMarketingUser=" & m_IsMarketingUser

            If m_Mode = "add" then 'Add mode open the Supply Chain dialog %>
                document.getElementById('UpperWindow').src = "<%=scURL%>";    
            <% ElseIf m_Mode = "editdates" then %>
                document.getElementById('UpperWindow').src = "<%=mktURL%>";
            <% Else            
                If m_IsMarketingUser and not m_IsPC and not isPulsarSystemAdmin Then 'Marketing users only %>
			        document.getElementById('UpperWindow').src = "<%=mktURL%>";
                <% Else 
			        if isPulsarSystemAdmin and m_IsMarketingUser then 'Both a Pulsar Admin and in a Marketing role, ask Pulsar Admin which dialog they want %>
				        var r = confirm('Press OK for Supply Chain dialog or Cancel for Marketing dialog');
				        if (r == true) {
					        document.getElementById('UpperWindow').src = "<%=scURL%>";
				        } else {
					        document.getElementById('UpperWindow').src = "<%=mktURL%>";
				        }
                    <% else %>
				       document.getElementById('UpperWindow').src = "<%=scURL%>";
			        <% end if %>
                <% end if %>
            <% end if %>
        }
    </script>
</HEAD>
	<FRAMESET ROWS="*,62" id="TopWindow" onload="LoadPage();">
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="#">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="avButtons.asp?<%=Request.QueryString%>">
	</frameset>
    <input type="hidden" id="hidIsMarketingUser" name="hidIsMarketingUser" value="<%=m_IsMarketingUser%>" />
    <input type="hidden" id="hidIsPC" name="hidIsPC" value="<%=m_IsPC%>" />
</html>

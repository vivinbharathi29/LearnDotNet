<%@ Language=VBScript %>
<% Option Explicit
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"  
%>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/EmailWrapper.asp" -->
<!-- #include file = "clsSchedule.asp" -->

<html>
<head>
<title></title>
<script src="../includes/client/jquery.min.js" type="text/javascript"></script>
<script type="text/javascript">
<!--
    $(function () {
        var OutArray = new Array();

        if ($("#txtSuccess").val() == "1") {
            OutArray[0] = $("#txtScheduleDataID_TargetRTPMR").val();
            OutArray[1] = $("#txtScheduleDataID_EM").val();
            OutArray[2] = $("#txtTargetRTPMR_Proj").val();
            OutArray[3] = $("#txtEM_Proj").val();
            OutArray[4] = $("#txtTargetRTPMR_Proj_OLD").val();
            OutArray[5] = $("#txtEM_Proj_OLD").val();

            var iframeName = parent.window.name;
            if (iframeName != '') {
                parent.window.parent.ClosePropertiesDialog(OutArray);
            } else {
                window.returnValue = OutArray;
                window.parent.close();
            }
        }
        else {
            document.write(document.body.innerHTML + "<BR><BR>Unable to update schedule.  An unexpected error occurred.");
        }
    });
    //-->
</script>
</head>
<body>

<%
	dim cn
    dim cn2
	dim dw
	dim cmd
    dim strSuccess
'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim m_IsSysAdmin
	Dim m_IsProgramManager
	Dim m_IsSysEngProgramManager
	Dim m_IsSysTeamLead
    Dim m_IsPlatformPm
    Dim m_IsSEPMProductsEditor
	Dim m_EditModeOn
	Dim m_UserEmail
	Dim m_IsSupplyChain
	
	m_EditModeOn = False
	
	Dim Security
	Dim sUserFullName
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
    
    ' Debug Section
    '
    '	If Security.CurrentUserID = 1396 Then
    '		m_IsSysAdmin = False
    '		Security.CurrentUserID = 1288
    '		Response.Write Security.CurrentUserID
    '		Response.Write "<BR>"
    '		Response.Write Security.IsProgramManager(Request("PVID"))
    '		Response.Write "<BR>"
    '		Response.Write Security.IsSysEngProgramManager(Request("PVID"))
    '		Response.Write "<BR>"
    '		Response.Write Security.IsSystemTeamLead(Request("PVID"))
    '		Response.Write "<BR>"
    '		Response.Write Request.QueryString
    '		Response.Write "<BR>"
    '		Response.Write Request.Form
    '		Response.Write "<BR>"
    '		Response.End
    '	End If

    m_IsSupplyChain = Security.UserInRole(Request("PVID"), "SupplyChain")
	m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
	m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
    m_IsPlatformPm = Security.IsPlatformDevMgr(Request("PVID"))
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	sUserFullName = Security.CurrentUser()
	m_UserEmail = Security.CurrentUserEmail()
	
	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsPlatformPm Or m_IsSupplyChain Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

	'Create Database Connection
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cn2 = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans
    cn2.BeginTrans

    Dim obj
    Set obj = New Schedule

    'Update Target RTP/MR Date
    CALL obj.UpdateMilestone(cn, NULL, Request.Form("scheduledataid60"), Trim(sUserFullName), m_UserEmail, NULL, NULL, Request.Form("hdn60"), Request.Form("hdn60"), Request.Form("txt60"), Request.Form("txt60"), NULL, NULL, NULL, NULL, 60, NULL)
    cn.CommitTrans

    'Update End of Manufacturing (EM) Date
    CALL obj.UpdateMilestone(cn2, NULL, Request.Form("scheduledataid112"), Trim(sUserFullName), m_UserEmail, NULL, NULL, Request.Form("hdn112"), Request.Form("hdn112"), Request.Form("txt112"), Request.Form("txt112"), NULL, NULL, NULL, NULL, 112, NULL)
    cn2.CommitTrans
    
    strSuccess = "1"

	Set dw = nothing
	Set cmd = nothing	
	Set cn = nothing
    Set cn2 = nothing

%>
<input type="text" id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>" /><br />

<input type="text" id="txtScheduleDataID_TargetRTPMR" name="txtScheduleDataID_TargetRTPMR" value="<%=Request.Form("scheduledataid60")%>" />
<input type="text" id="txtScheduleDataID_EM" name="txtScheduleDataID_EM" value="<%=Request.Form("scheduledataid112")%>" />

<input type="text" id="txtTargetRTPMR_Proj" name="txtTargetRTPMR_Proj" value="<%=Request.Form("txt60")%>" />
<input type="text" id="txtEM_Proj" name="txtEM_Proj" value="<%=Request.Form("txt112")%>" />

<input type="text" id="txtTargetRTPMR_Proj_OLD" name="txtTargetRTPMR_Proj_OLD" value="<%=Request.Form("hdn60")%>" />
<input type="text" id="txtEM_Proj_OLD" name="txtEM_Proj_OLD" value="<%=Request.Form("hdn112")%>" />

</body>
</html>

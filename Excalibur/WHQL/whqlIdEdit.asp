<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->
<%

Dim rs, dw, cn, cmd
Dim BrandID, BrandDesc, BrandRadioButtons, BUID, BUDesc, BUCheckBoxes, CPUID, CPUDesc, CPUCheckBoxes

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sMode				: sMode = Request.QueryString("Mode")
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim sFeatureCat			: sFeatureCat = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sConfigRules		: sConfigRules = ""
Dim iBrandID			: iBrandID = ""
Dim sTblBody            : sTblBody = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_EditModeOn
Dim m_UserFullName


'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

'	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
'	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
'	If m_IsSysAdmin Then
		m_EditModeOn = True
'	End If
	
'	If Not m_EditModeOn Then
'		Response.Write "<H3>Insufficient User Privilidges</H3><H4>Access Denied</H4>"
'		Response.End
'	End If

	Set Security = Nothing
'##############################################################################	
'

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Sub Main()
'
' TODO: Select Base Units based on PVID
'

    Dim lastBrandID, firstDiv

	Set cmd = dw.CreateCommandSP(cn, "usp_SelectWHQLSubmissions")
	dw.CreateParameter cmd, "@p_ProductWhqlID", adInteger, adParamInput, 8, Trim(Request("WHQLID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)

    If rs.State > 0 Then
        Do Until rs.EOF
            sTblBody = sTblBody & "<TR><TD>" & rs("SubmissionID") & "</TD>" & _
                "<TD>" & rs("SubmissionDt") & "</TD>" & _
                "<TD nowrap>" & rs("OS") & "</TD>" & _
                "<TD nowrap>" & rs("BU") & "</TD>" & _
                "<TD nowrap>" & rs("CPU") & "</TD></TR>"

            rs.MoveNext
        Loop		
    End If

	rs.Close
		
End Sub

Sub Save()
On Error Goto 0
End Sub

Call Main()
%>
<html>
<head>
    <title>Excalibur :: WHQL Submission ID Detail</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css">

    <script language="JavaScript" src="../includes/client/Common.js"></script>

    <script type="text/javascript">
function Body_OnLoad()
{
}

</script>

</head>
<body onload="Body_OnLoad()">
    <form method="post" id="frmMain" action="whqlIdDetail.asp" />
        <input id="hidMode" name="hidMode" type="HIDDEN" value="<%= LCase(sMode)%>" />
        <input id="hidFunction" name="hidFunction" type="HIDDEN" value="<%= LCase(sFunction)%>" />
        <input id="PVID" name="PVID" type="HIDDEN" value="<%= Request("PVID") %>" />
        
        <table class="FormTable" cellpadding="1" cellspacing="0">
            <col width="10px" /><col width="10px" /><col width="35%" /><col width="35%" />
            <tr><th>Submission ID</th><th>Submission Dt</th><th>OS Family</th><th>Base Unit</th><th>CPU</th></tr>
            <%= sTblBody %>
            </table>
    </form>
</body>
</html>

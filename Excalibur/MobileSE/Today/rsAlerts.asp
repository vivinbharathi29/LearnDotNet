<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file = "../../includes/Security.asp" --> 
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/no-cache.asp" -->
<%Dim AppRoot
AppRoot = Session("ApplicationRoot")

'
' Setup the data connections
'
Dim rs, dw, cn, cmd
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim functionCalled
functionCalled = Request.Form("Function")

If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "CloseAlert"
            Dim alertId
            alertId = Request("AlertId")
            Response.Write(CloseAlert(alertId))
        Case Else
            Response.Write("No Function Called")
    End Select
End If

function CloseAlert(AlertId) 
	on error resume next 
    Dim rs, dw, cn, cmd, returnValue

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateAlertStatus")
    dw.CreateParameter cmd, "@p_AlertID", adInteger, adParamInput, 8, AlertId
    dw.CreateParameter cmd, "@p_Active", adBoolean, adParamInput, 8, false
    returnValue = dw.ExecuteNonQuery(cmd)

    If clng(returnValue) = 1 Then
        CloseAlert = AlertId
    Else
        CloseAlert = 0
    End If

    set cmd = nothing
    set cn = nothing
    set dw = nothing
    
end function

%>

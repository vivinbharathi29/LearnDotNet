<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/emailwrapper.asp" -->
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd, rs
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
Set rs = Server.CreateObject("ADODB.RecordSet")
Dim sMessage		        : sMessage = ""

Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "CreateSimpleAv"
            if Request.Form("Pulsar") = 1 then
                Set cmd = dw.CreateCommAndSP(cn, "usp_CreateSimpleAv_Pulsar")
                dw.CreateParameter cmd, "@p_AvCreateID", adInteger, adParamInput, 8, Request.Form("AvCreateID")
                dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
                dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request.Form("UserID")
                dw.ExecuteNonQuery(cmd)
    		    Set rs = dw.ExecuteCommandReturnRS(cmd)
                sMessage = rs("Message") 'returns feature names if GPG description is empty
                rs.Close
            else
                Set cmd = dw.CreateCommAndSP(cn, "usp_CreateSimpleAv")
                dw.CreateParameter cmd, "@p_AvCreateID", adInteger, adParamInput, 8, Request.Form("AvCreateID")
                dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
                dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request.Form("UserID")
                dw.ExecuteNonQuery(cmd)


            end if
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            'Response.Write("Added Record")
            Response.Write(sMessage)
        Case "ChangeCategory"            
            Set cmd = dw.CreateCommAndSP(cn, "usp_EditCategoryforAvCreate")
            dw.CreateParameter cmd, "@p_AvCreateID", adInteger, adParamInput, 8, Request.Form("AvCreateID")
            dw.CreateParameter cmd, "@p_CategoryID", adInteger, adParamInput, 8, Request.Form("CategoryID")
            dw.ExecuteNonQuery(cmd)            
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Updated Record")
        Case Else
            set cn = nothing
            Response.Write("No Function Called")
    End Select
End If

%>
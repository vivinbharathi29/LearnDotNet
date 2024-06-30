<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/emailwrapper.asp" -->
<!-- #include file="../../includes/Security.asp" --> 
<%
    Dim dw, cn, cmd
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Dim m_CurrentUserEmail
    Dim Security
    Set Security = New ExcaliburSecurity
    m_CurrentUserEmail = Security.CurrentUserEmail()

    Dim FunctionCalled : FunctionCalled = Request.Form("Function")
    Select Case FunctionCalled
        Case "UpdateFeature"
            Set cmd = dw.CreateCommAndSP(cn, "usp_Feature_UpdateOverrideRequestedFeature")
            dw.CreateParameter cmd, "@p_intFeatureID", adInteger, adParamInput, 8, Request.Form("FeatureIDs")
            dw.CreateParameter cmd, "@p_intActionType", adInteger, adParamInput, 8, Request.Form("ActionType")
            dw.CreateParameter cmd, "@p_chrUser", adVarChar, adParamInput, 250, Request.Form("CurrentUserName")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing            
        Case "SendEmail"
            Dim EmailID 
            Dim HTMLBody
            HTMLBody = ""
            EmailID = Request.Form("Email")
            if (Request.Form("ActionType") = 0) then
                HTMLBody = "<html><head><title></title></head><body><font size=3 face=calibri><b>Feature Naming Override Request</b></font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Feature ID:</b> " & Request.Form("FeatureIDs") & "</font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Feature Name:</b> " & Request.Form("FeatureName") & "</font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Completed By:</b> " & Request.Form("CurrentUserName") & "</font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Status:</b> Complete</font>"                
            else
                HTMLBody = "<html><head><title></title></head><body><font size=3 face=calibri><b>Feature Naming Override Request:</b></font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Comment:</b> " & Request.Form("Comment") & "</font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Completed By:</b> " & Request.Form("CurrentUserName") & "</font>"
                HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Status:</b> Not Actionable</font>"                
            end if
            If HTMLBody > "" Then
                Dim oMessage
                Set oMessage = New EmailWrapper	
	            oMessage.From = m_CurrentUserEmail
	            oMessage.To = EmailID
	            oMessage.Subject = "Feature Naming Override Request"
    		    oMessage.HTMLBody = HTMLBody
	            oMessage.Send 
	            Set oMessage = Nothing 	
            End If            
        Case else
            Response.Write("No Function Called")
     End Select

 %>
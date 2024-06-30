<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/emailwrapper.asp" -->
<!-- #include file="../../includes/Security.asp" --> 
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim m_CurrentUserEmail
Dim Security
Set Security = New ExcaliburSecurity
m_CurrentUserEmail = Security.CurrentUserEmail()

Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "UpdatePhWebAvData"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdatePhWebAvData")
            dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request.Form("AVDetailID")
            dw.CreateParameter cmd, "@p_Last_Upd_User", adVarChar, adParamInput, 50, Request.Form("CurrentUser")
            dw.CreateParameter cmd, "@p_CPLBlindDate", adDate, adParamInput, 10, Request.Form("CPLBlindDate")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Added Record")
        Case "UpdatePhWebEntryComplete"
                Set cmd = dw.CreateCommAndSP(cn, "usp_UpdatePhWebEntryComplete2")
                dw.CreateParameter cmd, "@p_AvActionItemID", adInteger, adParamInput, 8, Request.Form("AvActionItemID")
                dw.CreateParameter cmd, "@p_PhWebActionTypeID", adInteger, adParamInput, 8, Request.Form("PhWebActionTypeID")
                dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
                dw.ExecuteNonQuery(cmd)
                set cmd = nothing
                set cn = nothing
                set dw = nothing
                Response.Write("Added Record")  
        Case "UpdatePDMFeedback"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdatePDMFeedback2")
            dw.CreateParameter cmd, "@p_AvActionItemID", adInteger, adParamInput, 8, Request.Form("AvActionItemID")
            dw.CreateParameter cmd, "@p_Last_Upd_User", adVarChar, adParamInput, 50, Request.Form("CurrentUser")
            dw.CreateParameter cmd, "@p_PDMFeedback", adVarChar, adParamInput, 100, Request.Form("PDMFeedback")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write(Request.Form("AvActionItemID") & "," & Request.Form("CurrentUser") & "," & Request.Form("PDMFeedback"))
        Case "ProcessPDMFeedbackEmail"
            Dim i
            Dim ActionItemArray
            Dim Product
            Dim HTMLBody
            Dim Feedback
            HTMLBody = ""
            ActionItemArray  = split(Request.Form("AvActionItemIDs"),",")
            Dim Email
            Dim rs
            Set rs = Server.CreateObject("ADODB.RecordSet")
            
            For i = lbound(ActionItemArray) To ubound(ActionItemArray)
                Set cmd = dw.CreateCommAndSP(cn, "usp_SelectPhWebPDMFeedback2")
                dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request.Form("BID")
                dw.CreateParameter cmd, "@p_AvActionItemId", adInteger, adParamInput, 8, ActionItemArray(i)
                Set rs = dw.ExecuteCommandReturnRS(cmd)
                'rs.Open "usp_SelectPhWebPDMFeedback2 " & clng(Request.Form("BID")) & "," & clng(ActionItemArray(i)),cn,adOpenForwardOnly
                
                If rs.state = 1 Then 
                    If Not (rs.EOF And rs.BOF) Then
                      Product = rs("Product")
                      Feedback = Replace(rs("PDMFeedback"),"][","<br/>")
                      Feedback = Replace(Feedback,"[","")
                      Feedback = Replace(Feedback,"]","")
                      If HTMLBody = "" Then
                        HTMLBody = "<html><head><title></title></head><body><font size=3 face=calibri><b>PDM Team Member:</b><br/>" & Request.Form("CurrentUser") & "</font>"
                        HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Product:</b><br/>" & rs("Product") & "</font>"
                        HTMLBody = HTMLBody & "<br/><br/><font size=3 face=calibri><b>Impacted AVs:</b></font>"
                        HTMLBody = HTMLBody & "<br/><font size=3 face=calibri><a href=""http://houhpqexcal03.auth.hpicorp.net/SCM/PDMFeedbackFrame.asp?AvId=" & rs("AvDetailID") & """>" & rs("AvNo") & "</a>&nbsp;&nbsp;-&nbsp;&nbsp;" & rs("GPGDescription") & "&nbsp;&nbsp;-&nbsp;&nbsp;" & rs("ActionName") & "&nbsp;(" & rs("PCDate") & ")"  & "<br/>" & Feedback & "</font><br/>"
                      Else
                        HTMLBody = HTMLBody & "<br/><font size=3 face=calibri><a href=""http://houhpqexcal03.auth.hpicorp.net/SCM/PDMFeedbackFrame.asp?AvId=" & rs("AvDetailID") & """>" & rs("AvNo") & "</a>&nbsp;&nbsp;-&nbsp;&nbsp;" & rs("GPGDescription") & "&nbsp;&nbsp;-&nbsp;&nbsp;" & rs("ActionName") & "&nbsp;(" & rs("PCDate") & ")"  & "<br/>" & Feedback & "</font><br/>"
                      End If
                      Email = rs("Email")
                    End If
                    rs.Close
                End If
            Next
            If HTMLBody > "" Then
                Dim oMessage
                Set oMessage = New EmailWrapper	
	            oMessage.From = m_CurrentUserEmail
	            oMessage.To = Email
	            oMessage.Subject = "PDM Feedback Available - " & Product
    		    oMessage.HTMLBody = HTMLBody
	            oMessage.Send 
	            Set oMessage = Nothing 	
            End If
            set cmd = nothing
            set cn = nothing
            set dw = nothing
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>
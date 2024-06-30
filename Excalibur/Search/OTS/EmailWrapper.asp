<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Library" --> 
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<%

Class EmailWrapper

	Dim m_Message
	Dim m_TestTo
	Dim m_TestMode 
	Dim m_From

	Private Sub Class_Initialize()
	    Dim oConfig
	    
		Set oConfig = Server.CreateObject("CDO.Configuration")
'	    oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=1 'cdoSendUsingPickup
        oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
        oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp3.hp.com"
	     oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 
	     oConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=0
	    oConfig.Fields.Update
		Set m_Message = CreateObject("CDO.Message") 
		Set m_Message.Configuration = oConfig
	    Set oConfig = Nothing
	    m_TestMode = False
	End Sub
	
	Private Sub Class_Terminate()
	On Error Resume Next
  		Set m_Message = Nothing
  	End Sub
	
	Public Property Let TextBody(value)
		m_Message.TextBody = value
	End Property
	
	Public Property Let HtmlBody(value)
	    If InStr(LCase(value), "confidential") = 0 Then
	        value = value & "<br><br><strong>HP Restricted</strong>"
	    End If
		m_Message.HTMLBody = value
	End Property
	
	Public Property Let Subject(value)
		m_Message.Subject = value & " - HP Restricted"
	End Property
	
	Public Property Let Importance(value)
	    m_Message.Fields.Item(cdoImportance) = value
	    m_Message.Fields.Update
	End Property
	
	Public Property Let [To](value)
	    If Application("Mode") <> "Production" Then
	        m_Message.To = Application("PulsarSupportEmail")
            m_TestTo = value
	    Else
		    m_Message.To = value
		End If
	End Property
	
	Public Property Let From(value)
	    If LCase(Right(value, 6)) <> "hp.com" Then
	        m_Message.From = "pulsar.support@hp.com"
	        m_From = value
	    Else
		    m_Message.From = value
		End If
	End Property
	
	Public Property Let CC(value)
	    m_Message.CC = value
	End Property
	
	Public Property Let BCC(value)
		m_Message.BCC = value
	End Property

	Public Property Let ReplyTo(value)
		m_Message.ReplyTo = value
	End Property

    Public Property Let DSNOptions(value)
        m_Message.DSNOptions = value
    End Property
    
    Public Property Let TestMode(value)
        m_TestMode = value
    End Property
    
    Public Sub AddRelatedBodyPart(Url, Reference, ReferenceType)
        m_Message.AddRelatedBodyPart Url, Reference, ReferenceType
    End Sub

    Public Sub AddAttachment(Url)
        m_Message.AddAttachment Url, "auth\svcdclerkam", "Recordify$08"
    End Sub

	Public Sub Send()
	    If Len(Trim(m_From)) > 0 Then
	        m_Message.CC = m_Message.CC & m_From
	        m_Message.CC = Replace(m_Message.CC, ";;", ";")
	    End If
	    If Len(Trim(m_Message.HTMLBody)) > 0 Then
	        m_Message.HTMLBodyPart.ContentTransferEncoding = cdoBase64
	    End If
	    If Application("Mode") <> "Production" Then
	        m_Message.TextBody = m_TestTo & vbcrlf & vbcrlf & m_Message.TextBody
	        m_Message.HTMLBody = m_TestTo & "<br><br>" & m_message.HTMLBody
	        m_Message.Subject = "[TEST] " & m_Message.Subject
	    End If
		m_Message.Send
	End Sub			
End Class

%>
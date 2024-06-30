
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Purpose:	Insert a record into MessageQueuedEmail, to send email out.
'' Created By:	Herb, 10/17/2016
''
'' Modify: Herb, 10/17/2016, Cannot attach file. (will add in the future PBI)
''
''         Herb, 11/2/2016, be able to attach files.
''         Herb, 06/14/2017, TICKET#: 11776, skip the input of "FromName" to avoid the display of "pulsar.support@hp.com on behalf of Pulsar.Support@hp.com"
''         Malichi, Jason - 07/14/2017 - BUG 146111 - Pulsar Test is sending TEST emails to users when testing in Pulsar Test
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class EmailQueue

    Dim mq_FromName
    Dim mq_ToName
    Dim mq_Subject
    Dim mq_Importance
    Dim mq_DSNOptions
    Dim mq_HTMLBody
    Dim mq_To
    Dim mq_From
    Dim mq_CC
    Dim mq_BCC
    Dim mq_Attachment

    Private Sub Class_Initialize()
	  
        mq_FromName = "" 
        mq_ToName =""
        mq_Subject =""
        mq_Importance =""
        mq_DSNOptions =""
        mq_HTMLBody =""
        mq_To =""
        mq_From =""
        mq_CC =""
        mq_BCC =""
        mq_Attachment =""

    End Sub

	Public Property Let TextBody(value)
		mq_HTMLBody = value
	End Property
  
	Public Property Let HtmlBody(value)
	    If InStr(LCase(value), "confidential") = 0 Then
	        value = value & "<br><br><strong>HP Restricted</strong>"
	    End If
		mq_HTMLBody = value
	End Property
	
	Public Property Let Subject(value)
        If Application("Mode") <> "Production" Then
		    mq_Subject = value & " [ on " & Application("Repository") & " ]" 
        else
	    mq_Subject = value & " - HP Restricted"
        end if
	End Property

	Public Property Let Importance(value)
	    mq_Importance = value ''' 11/3/2016, Herb, just for Compatibility, not work, will add in the future PBI if need.
	End Property
	
    Public Property Let DSNOptions(value)
        mq_DSNOptions = value ''' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    End Property

    Public Property Let FromName(value)
	    mq_FromName = value
	    
	End Property

    Public Property Let ToName(value)
	    mq_ToName = value
	    
	End Property

    Public Property Let From(value)
	    mq_From = value
	    
	End Property

	Public Property Let [To](value)
        mq_To = value

	End Property
	
	Public Property Get [To]()
       [To]= mq_To
	End Property
	
	Public Property Let CC(value)
	    mq_CC = value
	End Property
	
	Public Property Let BCC(value)
		mq_BCC = value
	End Property

	Public Property Let ReplyTo(value)
		mq_Importance = value  'dummy'
	End Property

    Public Property Let TestMode(value)
        mq_Importance = value 'dummy'
    End Property

    Public Sub AddAttachment(filepath)
        If Trim(mq_Attachment) = "" Then
            mq_Attachment = filepath
        Else
            mq_Attachment = mq_Attachment + ";" + filepath
        End If
    End Sub

	Public Sub Send()

	    If Len(Trim(mq_From)) > 0 Then
            If Len(Trim(mq_CC)) > 5 Then 'Because 6 is the minimum email length
	            mq_CC = mq_CC & ";" & mq_From
            Else
                mq_CC = mq_From
            End If
            mq_CC = Replace(mq_CC, " ", "") 'Remove space
	        mq_CC = Replace(mq_CC, ";;", ";") 'Remove repeated ";"
	    End If

  	    dim cn
	    dim cm
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.CommandTimeout = 180
	    cn.Open
	    set cm = server.CreateObject("ADODB.Command")
		
	    cm.ActiveConnection = cn
	    cm.CommandType = &H0004
	    cm.CommandText = "usp_SendMessageQueuedEmail"
	
		Set p = cm.CreateParameter("@From",adVarChar, &H0001,500)
	    p.Value = left(mq_From,500)
	    cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@FromName",adVarChar, &H0001,500)
	    p.Value = ""
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@To",adVarChar, &H0001,5000)
	    p.Value = left(mq_To ,5000)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@ToName",adVarChar, &H0001,500)
	    p.Value = left(mq_ToName ,500)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@Cc",adVarChar, &H0001,2000)
	    p.Value = left(mq_CC ,2000)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@Bcc",adVarChar, &H0001,2000)
	    p.Value = left(mq_BCC ,2000)
	    cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Subject",adVarChar, &H0001,500)
	    p.Value = left(mq_Subject,500)
	    cm.Parameters.Append p
    
		Set p = cm.CreateParameter("@Body",adLongVarWChar, &H0001,-1)
	    p.Value = mq_HTMLBody
	    cm.Parameters.Append p

    	Set p = cm.CreateParameter("@AttachmentPaths",adVarChar, &H0001,2000)
	    p.Value = mq_Attachment
	    cm.Parameters.Append p

	    cm.Execute
	    Set cm = Nothing	
	    cn.Close
	    set cn=nothing

	End Sub			
    'add another function to send an email without adding sender to CC field
    Public Sub SendWithOutCopy()

	    If Len(Trim(mq_CC)) > 0 Then
            mq_CC = Replace(mq_CC, " ", "") 'Remove space
	        mq_CC = Replace(mq_CC, ";;", ";") 'Remove repeated ";"
	    End If

  	    dim cn
	    dim cm
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.CommandTimeout = 180
	    cn.Open
	    set cm = server.CreateObject("ADODB.Command")
		
	    cm.ActiveConnection = cn
	    cm.CommandType = &H0004
	    cm.CommandText = "usp_SendMessageQueuedEmail"
	
		Set p = cm.CreateParameter("@From",adVarChar, &H0001,500)
	    p.Value = left(mq_From,500)
	    cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@FromName",adVarChar, &H0001,500)
	    p.Value = ""
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@To",adVarChar, &H0001,5000)
	    p.Value = left(mq_To ,5000)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@ToName",adVarChar, &H0001,500)
	    p.Value = left(mq_ToName ,500)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@Cc",adVarChar, &H0001,2000)
	    p.Value = left(mq_CC ,2000)
	    cm.Parameters.Append p	

		Set p = cm.CreateParameter("@Bcc",adVarChar, &H0001,2000)
	    p.Value = left(mq_BCC ,2000)
	    cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Subject",adVarChar, &H0001,500)
	    p.Value = left(mq_Subject,500)
	    cm.Parameters.Append p
    
		Set p = cm.CreateParameter("@Body",adLongVarWChar, &H0001,-1)
	    p.Value = mq_HTMLBody
	    cm.Parameters.Append p

	    cm.Execute
	    Set cm = Nothing	
	    cn.Close
	    set cn=nothing

	End Sub				
End Class

%>
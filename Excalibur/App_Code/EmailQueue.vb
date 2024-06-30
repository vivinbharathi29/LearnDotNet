Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class EmailQueue

    Dim mq_FromName As String
    Dim mq_ToName As String
    Dim mq_Subject As String
    Dim mq_Importance As String '' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    Dim mq_DSNOptions As String '' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    Dim mq_HTMLBody As String
    Dim mq_To As String
    Dim mq_From As String
    Dim mq_CC As String
    Dim mq_BCC As String
    Dim mq_ReplyTo As String  '' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    Dim mq_Attachment As String

    Private Sub Class_Initialize()
        mq_FromName = ""
        mq_ToName = ""
        mq_Subject = ""
        mq_Importance = ""
        mq_DSNOptions = ""
        mq_HTMLBody = ""
        mq_To = ""
        mq_From = ""
        mq_CC = ""
        mq_BCC = ""
        mq_ReplyTo = ""
        mq_Attachment = ""
    End Sub

    Public Sub New()
        'MyBase.New()
        Class_Initialize()
    End Sub

    Property TextBody As String
        Get
            Return mq_HTMLBody
        End Get
        Set(ByVal value As String)
            mq_HTMLBody = value
        End Set
    End Property

    Property HtmlBody As String
        Get
            Return mq_HTMLBody
        End Get
        Set(ByVal value As String)
            If InStr(LCase(value), "confidential") = 0 Then
                mq_HTMLBody = value & "<br><br><strong>HP Restricted</strong>"
            Else
                mq_HTMLBody = value & "<br><br><strong>HP Confidential</strong>"
            End If

        End Set
    End Property

    Property Subject As String
        Get
            Return mq_Subject
        End Get
        Set(ByVal value As String)
            mq_Subject = value
        End Set
    End Property

    Property AddTo As String
        Get
            Return mq_To
        End Get
        Set(ByVal value As String)
            mq_To = value
        End Set
    End Property

    Property AddFrom As String
        Get
            Return mq_From
        End Get
        Set(ByVal value As String)
            mq_From = value
        End Set
    End Property

    Property AddCC As String
        Get
            Return mq_CC
        End Get
        Set(ByVal value As String)
            mq_CC = value
        End Set
    End Property

    Property AddBCC As String
        Get
            Return mq_BCC
        End Get
        Set(ByVal value As String)
            mq_BCC = value
        End Set
    End Property

    ''' <summary>
    ''' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    ''' </summary>
    ''' <returns></returns>
    Property ReplyTo As String
        Get
            Return mq_ReplyTo
        End Get
        Set(ByVal value As String)
            mq_ReplyTo = value
        End Set
    End Property

    ''' <summary>
    ''' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    ''' </summary>
    ''' <returns></returns>
    Property DSNOptions As String
        Get
            Return mq_DSNOptions
        End Get
        Set(ByVal value As String)
            mq_DSNOptions = value
        End Set
    End Property

    ''' <summary>
    ''' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    ''' </summary>
    ''' <returns></returns>
    Property Importance As String
        Get
            Return mq_Importance
        End Get
        Set(ByVal value As String)
            mq_Importance = value
        End Set
    End Property

    ''' <summary>
    ''' 29/8/2017, Herb, just for Compatibility, not work, add in the future PBI if need.
    ''' </summary>
    ''' <returns></returns>
    Property TestMode As String
        Get
            Return ""
        End Get
        Set(ByVal value As String)
            '_TestMode = value
        End Set
    End Property

    Public Sub AddAttachment(ByVal filepath As String)
        If Trim(mq_Attachment) = "" Then
            mq_Attachment = filepath
        Else
            mq_Attachment = mq_Attachment + ";" + filepath
        End If

    End Sub

    ''' <summary>
    ''' Without cc to the sender
    ''' </summary>
    Public Sub Send()
        If Len(Trim(mq_CC)) > 0 Then
            mq_CC = Replace(mq_CC, " ", "") 'Remove space
            mq_CC = Replace(mq_CC, ";;", ";") 'Remove repeated ";"
        End If

        Dim rowsAffected As Integer = 0

        Using cn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("Excalibur"))
            Dim cc As New SqlCommand()
            cc.Connection = cn
            cc.CommandType = CommandType.StoredProcedure
            cc.CommandText = "usp_SendMessageQueuedEmail"

            cc.Parameters.AddWithValue("@From", Left(mq_From, 500))
            cc.Parameters.AddWithValue("@FromName", Left(mq_FromName, 500))
            cc.Parameters.AddWithValue("@To", Left(mq_To, 5000))
            cc.Parameters.AddWithValue("@ToName", Left(mq_ToName, 500))
            cc.Parameters.AddWithValue("@Cc", Left(mq_CC, 2000))
            cc.Parameters.AddWithValue("@Bcc", Left(mq_BCC, 2000))
            cc.Parameters.AddWithValue("@Subject", Left(mq_Subject, 500))
            cc.Parameters.AddWithValue("@Body", mq_HTMLBody)
            cc.Parameters.AddWithValue("@AttachmentPaths", mq_Attachment)

            cn.Open()
            rowsAffected = cc.ExecuteNonQuery()

        End Using




    End Sub
End Class

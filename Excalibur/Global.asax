<%@ Application Language="VB" %>

<script runat="server">
    Function GetAttribute(AttrName)
        Dim AllAttrs
        Dim RealAttrName
        Dim Location
        Dim Result
        AllAttrs = Request.ServerVariables("ALL_HTTP")
        RealAttrName = AttrName
        Location = InStr(AllAttrs, RealAttrName & ":")
        If Location <= 0 Then
            GetAttribute = ""
            Exit Function
        End If
        Result = Mid(AllAttrs, Location + Len(RealAttrName) + 1)
        Location = InStr(Result, Chr(10))   'LF character
        If Location <= 0 Then
            Location = Len(Result) + 1
        End If

        GetAttribute = Left(Result, Location - 1)
    End Function

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application startup
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when an unhandled error occurs        
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        Dim ServerName As String = Request.ServerVariables("HTTP_HOST")
        If UCase(Left(ServerName, 3)) = "PRP" Then
            ServerName = ServerName & "/excalibur"
            Session("ApplicationRoot") = "/excalibur"
            Session("ServerName") = ServerName
        Else
            ServerName = Request.ServerVariables("SERVER_NAME")
            Session("ApplicationRoot") = ""
            Session("ServerName") = ServerName
        End If



        'Get user name from Site Minder
        ' This will look like "uid=first.last@odmdomain.com,ou=****,****"
        ' We need to get just "first.last@odmdomain.com"
        Dim User As String = String.Empty

        'Get user name from window authentication
        User = Request.ServerVariables("LOGON_USER")

        'Get user name from UID
        If Trim(User) = "" Then
            User = GetAttribute("HTTP_HPPF_AUTH_NTUSERDOMAINID")
            User = Replace(User, ":", "\")
        End If

        'Get email if username is not available
        If Trim(User) = "" Then
            User = GetAttribute("HTTP_HPPF_AUTH_UID")
        End If

        Session("LoggedInUser") = User
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
    End Sub
</script>
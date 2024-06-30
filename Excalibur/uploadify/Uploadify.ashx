<%@ WebHandler Language="VB" Class="Uploadify" %>

Imports System
Imports System.IO
Imports System.Web

Public Class Uploadify : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim oFile As HttpPostedFile = context.Request.Files("Filedata")
        If (Not oFile Is Nothing) Then
            Dim sDirectory As String = Path.Combine(context.Request.PhysicalApplicationPath, "temp")
            If (Not Directory.Exists(sDirectory)) Then
                Directory.CreateDirectory(sDirectory)
            End If
            
            oFile.SaveAs(Path.Combine(sDirectory, oFile.FileName))
            
            context.Response.Write("1")
            
        Else
            
            context.Response.Write("0")
            
        End If
        
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class

<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Private ReadOnly Property ProductBrandID() As String
        Get
            Return Request.QueryString("PBID")
        End Get
    End Property
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Cache.SetExpires(DateTime.Now())
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim sec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(User.Identity.Name)
        
        If sec.IsProgramCoordinator() Then
            dw.RollbackScmPublish(ProductBrandID)
        End If
        
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>SCM Publish Rollback</title>
</head>
<body onload="window.close();">
<p>&nbsp;</p>
</body>
</html>

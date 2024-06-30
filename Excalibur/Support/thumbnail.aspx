<%@Import Namespace="System.Drawing.Imaging" %>
<script language="VB" runat="server">
  Function ThumbnailCallback() as Boolean
    Return False
  End Function


  Sub Page_Load(sender as Object, e as EventArgs)
  
    'Read in the image filename to create a thumbnail of
    Dim imageUrl as String = Request.QueryString("img")
    
    'Read in the width and height
    Dim imageHeight as Integer = Request.QueryString("h")
    Dim imageWidth as Integer = Request.QueryString("w")

    'Make sure that the image URL doesn't contain any /'s or \'s
'    If imageUrl.IndexOf("/") >= 0 Or imageUrl.IndexOf("\") >= 0 then
'      'We found a / or \
'      Response.End()
'    End If
    
    'Add on the appropriate directory
 '   imageUrl = "/images/" & imageUrl
    
        Dim fullSizeImg As System.Drawing.Image
        fullSizeImg = System.Drawing.Image.FromFile(Server.MapPath(imageUrl))
    

        'Do we need to create a thumbnail?
        Response.ContentType = "image/gif"
        If imageHeight > 0 And imageWidth > 0 Then
            Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
            dummyCallBack = New  _
               System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)

            Dim thumbNailImg As System.Drawing.Image
            thumbNailImg = fullSizeImg.GetThumbnailImage(imageWidth, imageHeight, _
                                                         dummyCallBack, IntPtr.Zero)

            thumbNailImg.Save(Response.OutputStream, ImageFormat.Gif)

            'Clean up / Dispose...
            thumbNailImg.Dispose()
        Else
            fullSizeImg.Save(Response.OutputStream, ImageFormat.Gif)
        End If
    
        'Clean up / Dispose...
        fullSizeImg.Dispose()
  End Sub
</script>

 

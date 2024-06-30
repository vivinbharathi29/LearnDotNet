
Partial Class Service_selCalDate
    Inherits System.Web.UI.Page

    Protected Sub theCalendar_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles theCalendar.SelectionChanged
        Dim objCalendar As Calendar = DirectCast(sender, Calendar)
        Dim strScript As String = "<script language='javascript'>window.returnValue = '" & objCalendar.SelectedDate.ToShortDateString() & "';window.close();</script>"

        Response.Write(strScript)
    End Sub
End Class

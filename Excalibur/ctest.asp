<SCRIPT LANGUAGE=VBScript RUNAT=Server>

Option Explicit

 

Dim strConnectString 
strConnectString = Session("PDPIMS_ConnectionString")

Const strSQL = "SELECT * FROM employee with (NOLOCK) WHERE name='This is SQL Statement 1'"

 

Dim cnPubs ' As New ADODB.Connection

Dim i

 

For i = 1 To 1000

    Response.Write "Iteration " & i & "<br>"

    Set cnPubs = Server.CreateObject("ADODB.Connection")

    cnPubs.ConnectionString = strConnectString

    cnPubs.Open

 

    cnPubs.Execute strSQL

    Response.Write "&nbsp;"  ' our simple delay tactic

 

    cnPubs.Close

    Set cnPubs = Nothing

    Response.Flush

Next '

Response.Write "<p>Done"

</script>

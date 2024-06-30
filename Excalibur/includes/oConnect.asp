<%
Dim oConnect		'Our connection object.
Function PULSARDB()
	PULSARDB = Session("PDPIMS_ConnectionString")
End Function

Sub oConnectClose()
	oConnect.close
	Set oConnect = Nothing
End Sub

'*************************************************************************************
'* Purpose		: Open/Close database connections.
'* Inputs		: WhichDB - database connection string.
'*				: bOpen - boolean value.  (TRUE to open DB connection or FALSE to close it.)
'* Returns		: oConnect - database connection object.
'*************************************************************************************
Sub OpenDBConnection(WhichDB, bOpen)
	On Error Resume Next
		Select Case bOpen
		Case True	'Open connection
			Set oConnect = Server.CreateObject("ADODB.Connection")
			oConnect.Open(WhichDB)
		Case False	'Close connection.
			If Not (oConnect is Nothing) Then
				Call oConnectClose()
			End If
		Case Else	'Custom error handling.
			Response.Write("Error accessing database. Contact Pulsar Administrator.")
            Response.End()
		End Select		
End Sub
%>


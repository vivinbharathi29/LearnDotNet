<%
'*************************************************************************************
'* FileName		: oDataConnection.asp
'* Description	: Class for AMO Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/27/2016 - PBI 17487/ Task 21005
'*************************************************************************************
Class DBConnection
    '--DECLARE LOCAL VARIABLES------------------------------------------------------------
    Dim oConnect			'DB OBJECT   

    '*************************************************************************************
	'* Purpose		: Build database connection string.
	'* Inputs		: None
	'* Returns		: PULSARDB - database connection string.
	'*************************************************************************************
	Public Function PulsarConnectionString() 
        PulsarConnectionString = Application("PDPIMS_ConnectionString") 
	End Function
    
    '*************************************************************************************
	'* Purpose		: Method to initialize connection to database for queries
	'* Inputs		: None
	'* Returns		: Set DB Connection
	'*************************************************************************************
    Public Sub InitDBConnection(bOpenConnection)
        If bOpenConnection = True Then
            If IsObject(oConnect) = False then
                Set oConnect = Server.CreateObject("ADODB.Connection")
                
                'open database connection
                oConnect.ConnectionString = PulsarConnectionString() 
                oConnect.Open

                'set database connection to Session
                Set Session("oConnect") = oConnect
            End If
        End If
    End Sub

    '*************************************************************************************
	'* Purpose		: Method to destroy Connection object created.
	'* Inputs		: None
	'* Returns		: Close DB Connection
	'*************************************************************************************
    Public Sub CloseDBConnection(bCloseConnection)
        If bCloseConnection = True Then
            If IsObject(Session("oConnect")) = True Then 
                Session("oConnect").Close
                Set Session("oConnect") = nothing
            End If 
        End If
    End Sub

End Class
%>
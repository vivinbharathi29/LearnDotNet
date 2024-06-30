<%
    '*************************************************************************************
    '* Purpose		: Initialize DB Connection Session - oDataConnection.asp
    '*************************************************************************************
    Dim oDBSvr
    Set oDBSvr = New DBConnection 
    oDBSvr.InitDBConnection(True)    
%>

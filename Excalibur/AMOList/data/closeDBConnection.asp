<%
    '*************************************************************************************
    '* Purpose		: Close DB Connection Session - oDataConnection.asp
    '*************************************************************************************
    Set oDBSvr = New DBConnection 
    oDBSvr.CloseDBConnection(True)
%>

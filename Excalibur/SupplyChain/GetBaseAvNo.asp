<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim cn, cmd

Dim m_AvNo	: m_AvNo = Request("AvNo")


function ScrubSQL(strWords) 

    dim badChars 
	dim newChars 
	dim i

	'strWords=replace(strWords,"'","''")
		
	badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
	newChars = strWords 
		
	for i = 0 to uBound(badChars) 
		newChars = replace(newChars, badChars(i), "") 
	next 
		
	ScrubSQL = newChars 
	
 end function

'##############################################################################	

set cn = server.CreateObject("ADODB.connection")
set cmd = Server.CreateObject("ADODB.Command")
'                      
cn.ConnectionString = Session("PDPIMS_ConnectionString")
cn.Open

rs.Open "select avdetailid from avdetail where AvNo = " & m_AvNo, cn, adOpenForwardOnly
do while not rs.EOF
    response.Write rs("avdetailid")
    rs.MoveNext
loop                           
rs.Close     


set cmd=nothing
set cn=nothing


%>
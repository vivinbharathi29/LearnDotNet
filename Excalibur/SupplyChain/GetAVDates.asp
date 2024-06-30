<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim cn, cmd, rs
Dim m_AvDetailId	: m_AvDetailId = Request("AvDetailId")

Set rs = Server.CreateObject("ADODB.RecordSet")
set cn = server.CreateObject("ADODB.connection")
set cmd = Server.CreateObject("ADODB.Command")
                       
cn.ConnectionString = Session("PDPIMS_ConnectionString")
cn.Open

rs.Open "select isnull(CPLBlindDt,'') as CPLBlindDt, isnull(GeneralAvailDt,'') as GeneralAvailDt, isnull(RASDiscontinueDt,'') as RASDiscontinueDt, isnull(RTPDate,'') as RTPDate, isnull(PHWebDate,'') as PHWebDate from avdetail where avdetailid = " & m_AvDetailId, cn, adOpenForwardOnly
do while not rs.EOF
    response.Write "CPLBlindDt:" & rs("CPLBlindDt") & ";GeneralAvailDt:" & rs("GeneralAvailDt") & ";RASDiscontinueDt:" & rs("RASDiscontinueDt") & ";RTPDate:" & rs("RTPDate") & ";PHWebDate:" & rs("PHWebDate")
    rs.MoveNext
loop                           
rs.Close  
      

set cmd=nothing
set cn=nothing

%>
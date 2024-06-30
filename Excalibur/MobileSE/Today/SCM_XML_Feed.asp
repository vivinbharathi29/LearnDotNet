<?xml version="1.0" encoding="UTF-8"?>

  <%  
  Response.ContentType = "text/xml"

  Dim cn 
  Dim rs
  Dim CurrentUserID
  Dim PBID: PBID = request.QueryString("PBID")

  Set cn = server.CreateObject("ADODB.Connection")
  cn.ConnectionString = Session("PDPIMS_ConnectionString") 
  cn.Open
  
  Set rs = server.CreateObject("ADODB.recordset")

  rs.Open "spFusion_SCM_ListSCMFeedErrors " & request("PBID"),cn,adOpenForwardOnly

  if not(rs.EOF and rs.BOF) then
    Response.Write rs("OutgoingMessage")
  end if
  

  rs.Close
  cn.Close

  %>

<%@ Language=VBScript %>
<!-- #include file = "../includes/noaccess.inc" -->
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</HEAD>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</STYLE>

<BODY>
<%
    dim cn
    dim rs
    dim strSQl

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    strSQL = "Select distinct r.id " & _
             "from DeliverableVersion v with (NOLOCK), DeliverableRoot r with (NOLOCK) " & _
             "where (v.DeveloperID = 6283 or testerid = 6283 or devmanagerid = 6283) " & _
             "and r.ID = v.deliverablerootid"
    rs.open strSQL,cn,adOpenStatic
    response.write "Preview Mode<BR><BR>"
    do while not rs.EOF
        response.write rs("ID") & "<BR>"
      '  populateSI "1," & rs("ID") & ",0,'Deliverable Root Refresh - No Products Updates'"
        rs.MoveNext
    loop

    rs.Close
    set rs = nothing
    cn.Close
    set cn = nothing


	
    Sub PopulateSI (strMessage)
        Dim objQInfo
        Dim objQSend
        Dim objMessage
     
        'open the queue
        Set objQInfo = Server.CreateObject("MSMQ.MSMQQueueInfo")
        objQInfo.PathName = ".\private$\SIExcalSync" 
        Set objQSend = objQInfo.Open(2, 0)
  
        'build/send the message
        Set objMessage = Server.CreateObject("MSMQ.MSMQMessage")
        objMessage.Body = "<?xml version=""1.0""?><string>" & strMessage & "</string>"
        objMessage.Send objQSend
        objQSend.Close

        'clean up
        Set objQInfo = Nothing
        Set objQSend = Nothing
        Set objMessage = Nothing
    end sub

%>
</BODY>
</HTML>





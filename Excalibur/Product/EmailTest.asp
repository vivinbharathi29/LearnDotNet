<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file="../includes/EmailQueue.asp" -->

<HTML>
<HEAD>
</HEAD>

<BODY>
<%
    	Set oMessage = New EmailQueue
	
        omessage.from = "Pulsar.Support@hp.com"
        oMessage.To= "Pulsar.Support@hp.com"
	    oMessage.Subject = "Test Email"
    	oMessage.Importance = cdoNormal
	    oMessage.HTMLBody = "See Attachment"
        oMessage.AddAttachment("E:\PulsarTest\Excalibur\temp\Dunes_3.0_ProBook_Lev_Rev-67-0-0-1.SCMX")
        oMessage.AddAttachment("\\HOUPulsarJob02.TEX.RD.HPICORP.NET\ExcaliburFileStore$\omf2yvsz.sqc\Dunes_3 0_ProBook_Lev_Rev-67-0-0-1.SCMX")
        
	    oMessage.Send 
	    Set oMessage = Nothing 			
	
%>
Message Sent


</BODY>
</HTML>

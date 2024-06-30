<!-- WHQLMenu.asp  -->
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function WHQLSignature(){
	var strResult;
	strResult = window.showModalDialog("../WHQL/WHQLSignature.asp","","dialogWidth:950px;dialogHeight:520px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No"); 
	if (typeof(strResult) != "undefined")
		{
			window.location.reload(true);
		}
}
//-->
</script>

<%

strServername = LCase(Request.ServerVariables("SERVER_NAME")) & "/mobilese"
strURL = LCase(Request.ServerVariables("URL"))	
if left(strurl ,1) = "/" then
	strurl = mid(strurl,2)
end if
Select Case strServername
	Case Application("Excalibur_ServerName")
		strServer2 = Application("Excalibur_ServerName")
		strServername = strServername 
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
	Case "localhost"
		strServer2 = "localhost"
		strServername = strServername
		strURL = mid(strURL, 16)
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
	Case Else
		strServer2 = strServername
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
end select


%>

<p><table cellPadding="1" border="0" cellSpacing="1" width="100%"> 
  <tr>
    <td nowrap width="240" style="VERTICAL-ALIGN: top" id="side">
    </H1>	
	<h1>
	<BR>
	<font face=verdana size=4 color=yellow>Links</font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a href="javascript:WHQLSignature();">Test Signature Request</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a target=_blank href="file://\\mitfilestorage.cca.cpqcorp.net\whql$\WHQL\Projects\Active Projects\Web\News\news.htm">News and Announcements</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a target=_blank href="file://\\mitfilestorage.cca.cpqcorp.net\whql$\WHQL\Projects\Active Projects\Web\Metrics\metrics.htm">Weekly Matrix</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a target=_blank href="file://\\mitfilestorage.cca.cpqcorp.net\whql$\WHQL\Projects\Active Projects\Web\Status\status.htm">Weekly Status Report</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a target=_blank href="file://\\mitfilestorage.cca.cpqcorp.net\whql$\WHQL\Projects\Active Projects\Web\Submission History\submissions.htm">Submission History</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a HREF="mailto:HouPSGNotebookWHQL@hp.com?subject=WHQL Suggestions">WHQL Suggestions</a></font><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><font color=white><a HREF="mailto:HouPSGNotebookWHQL@hp.com?subject=WHQL Questions">WHQL Questions</a></font><br>      
     
      <br> 

      </h1>      
</td>

    <td style="VERTICAL-ALIGN: top">      

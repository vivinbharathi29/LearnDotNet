<%@  language="VBScript" %>

<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<head>
    <title>Release Notes</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script language="javascript">
        $(document).ready(function () {
            var pulsarplusDivId = document.getElementById("PulsarPlusDivId").value;
            if (pulsarplusDivId != undefined && pulsarplusDivId != "")
                document.getElementById("tblClose").style.display = "block";
        });
        function cmdClose_onclick(pulsarplusDivId) {
            var pulsarplusDivId = document.getElementById("PulsarPlusDivId").value;
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                // For Closing current popup
                parent.window.parent.ClosePulsarPlusPopup();
            }
        }
    </script>
</head>
<style>
    TD {
        font-size: x-small;
    }

    body {
        background: #fcfdfd;
    }
</style>
<body>


    <%
    
    
    getReleaseNotes(Request("ID"))
    
    function getReleaseNotes(ID)
    dim strChanges
    dim strOTSList
	dim strSelectedOTS
	
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.ConnectionTimeout = 60
	cn.IsolationLevel=256
	cn.commandtimeout = 180
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	if ID = "" then 'no deliverable version is specified
		Response.Write "No deliverable version is specified"		
	else
		rs.Open "spGetVersionProperties4Web " & clng(ID),cn,adOpenForwardOnly
			strChanges = rs("Changes") & ""
        rs.Close

		strOTSList = ""
		strSelectedOTS = ""
		strOTSList = mid(strOTSList,2)
    	rs.Open "spGetOTSByDelVersion "  & clng(ID),cn,adOpenForwardOnly
			if cn.Errors.count = 0 then		
        		do while not rs.EOF
					strOTSList = strOTSList &  rs("OTSNumber") & ","
					strSelectedOTS = strSelectedOTS & rs("OTSNumber") & " - " &  rs("shortdescription") 
					rs.MoveNext
				loop
    		end if
        rs.Close

        response.Write "<table>"
	    response.Write "<TR><TD colspan=3><TABLE width=""100%"">"
	    response.Write "<td><font face=verdana size=1><b>Observations fixed in this release:</b>&nbsp;&nbsp;&nbsp;"
	    response.Write 	strSelectedOTS
	    response.Write "<br><br></font></td>"
	    response.Write "</TABLE></td></tr>"

	    response.Write "<TR><TD colspan=3><TABLE width=""100%"">"
	    response.Write "<td><font face=verdana size=1><b>Modifications, Enhancements, or Reason for Release:</b>&nbsp;&nbsp; "
	    response.Write "	<BR>"
        response.Write replace(strChanges,vbcrlf,"<BR>")
	    response.Write "</font></td>"
	    response.Write "</TABLE></td></tr>"
        response.Write "</table>"

    
    end if
    
    set rs = nothing
	cn.close
	set cn = nothing
End Function
    %>
    <table style="display:none" id="tblClose" width="100%" border="0">
        <tr>
            <td align="right">
                <input type="button" value="Close" id="cmdClose" name="cmdClose" language="javascript" onclick="return cmdClose_onclick()">
            </td>
        </tr>
    </table>
<input type="hidden" id="PulsarPlusDivId" name="PulsarPlusDivId" value="<%= Request("pulsarplusDivId")%>">
</body>
</html>

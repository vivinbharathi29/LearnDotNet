<%@ Language=VBScript %>

	<%
	
  'Response.Buffer = True
  'Response.ExpiresAbsolute = Now() - 1
  'Response.Expires = 0
  'Response.CacheControl = "no-cache"
	%>
<HTML>
<STYLE>
A:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}


</STYLE>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>
<meta http-equiv="refresh" content="5" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (txtRestart.value != '')
            window.location.href = "MFTStatus.asp?ID=" + txtID.value;

        //Update the image path on the wizard window if it completes successfully.
        if (window.parent.AddVersion.txtLocation.value.substring(0, 19).toLowerCase() == "softwarecomponents/" && txtMFTStatusConstant.value == "HPMGR_REQUEST_COMPLETED" && txtIRSPath.value != "")
            window.parent.AddVersion.txtLocation.value = txtIRSPath.value;

    }

    function restart() {
        //Do the validation here

        //restart
        window.location.href = "MFTStatus.asp?ID=" + txtID.value + "&Restart=1&NewPath=" + encodeURIComponent(window.parent.AddVersion.txtLocation.value);
    }

//-->
</SCRIPT>
<style>
    body
    {
        font-family: Verdana;
        font-size: x-small;
    }

</style>
</HEAD>
<body bgcolor="cornsilk" language="javascript" onload="return window_onload()">
<b>MFTStatus:</b>
<%
	dim cn
	dim rs
    dim cm
    dim strIRSPath
    dim strMFTStatusConstant
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    
    if request("ID")  = "" then
        response.write "No ID sepecified."
    else
        if request("Restart") = "1" then
            response.write "Restarting..."

            'Call the restart procedure
            set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
	        cm.CommandType = 4
	        cm.CommandText = "spFusion_COMPONET_RestartMFTTransfer"
	

	        Set p = cm.CreateParameter("@VersionID", 3, &H0001)
	        p.Value = clng(request("ID"))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@ImagePath", 200, &H0001, 256)
	        p.Value = request("NewPath")
	        cm.Parameters.Append p

		    cm.Execute rowschanged
					
	        set cm=nothing

        else
            rs.open "spGetDeliverableMFTStatus " & clng(request("ID")),cn
            if rs.eof and rs.bof then
                response.write "Unknown"
            elseif trim(ucase(rs("MFTStatusConstant"))) =  "HPMGR_REQUEST_FAILED" then
                if trim(rs("Comments")) <> "" then
                    response.write rs("Status") & "&nbsp;&nbsp;<a href=""javascript: alert('" & rs("Comments") & "')"">Show Error</a>&nbsp;|&nbsp;<a href=""javascript:restart();"">Restart</a>"
 '                   response.write rs("Status") & "&nbsp;&nbsp;<a href=""javascript: alert('" & rs("Comments") & "')"">Show Error</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""MFTStatus.asp?ID=" & clng(request("ID")) & "&Restart=1"">Restart</a>"
                else
                    response.write rs("Status") & "&nbsp;&nbsp;" & "<a href=""javascript:restart();"">Restart</a>"
'                    response.write rs("Status") & "&nbsp;&nbsp;" & "<a href=""MFTStatus.asp?ID=" & clng(request("ID")) & "&Restart=1"">Restart</a>"
                end if
                response.write "&nbsp;|&nbsp;<a target=_blank href=""https://mftp.b2b.americas.hp.com/portal-seefx"">Log-in to MFT</a>"
            else
                response.write rs("Status") & ""
            end if
            strIRSPath = rs("IRSPath") & ""
            strMFTStatusConstant = rs("MFTStatusConstant") & ""
            rs.close
       end if
    end if

    set rs = nothing
    cn.close
    set cn=nothing
%>
    <input id="txtRestart" type="hidden" value="<%=request("Restart")%>" />
    <input id="txtID" type="hidden" value="<%=request("ID")%>" />
    <input id="txtIRSPath" type="hidden" value="<%=server.htmlencode(strIRSPath)%>" />
    <input id="txtMFTStatusConstant" type="hidden" value="<%=server.htmlencode(strMFTStatusConstant)%>" />
    
</body>
</HTML>
<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<%if request("Status") = 1 then%>
<TITLE>Preinstall Prep Complete</TITLE>
<%else%>
<TITLE>Preinstall Perp Cancelled</TITLE>
<%end if%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value=="1")
			{
			window.returnValue = 1;
			window.close();
			}
		}
	else
		{
		document.write ("Unable to update this deliverable.  An unexpected error occurred.");
		}


}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<BR>
<table width=100%><TR><TD align=center>
<font face=verdana size =2>Updating Version.  Please wait...</font>
</td></tr></table>
<%

	if (request("VersionList") = "" or request("Status") = "") then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
	else


	    dim cn
	    dim cm
    	
    	set cn = server.CreateObject("ADODB.Connection")
    	
    	cn.ConnectionString = Session("PDPIMS_ConnectionString")
    	cn.Open
    
    
		dim IDArray
		IDArray = split(request("VersionList"),",")
		
        cn.BeginTrans
		
        for i = 0 to ubound(IDArray)
			set cm = server.CreateObject("ADODB.Command")

			cm.ActiveConnection = cn
			cm.CommandText = "spUpdatePreinstallPrepStatus"
			cm.CommandType = &H0004
			       
			Set p = cm.CreateParameter("@VersionID",adInteger, &H0001)
			p.Value = clng(IDarray(i))
			cm.Parameters.Append p
			
			Set p = cm.CreateParameter("@Status",adInteger, &H0001)
			p.Value = clng(request("Status"))
			cm.Parameters.Append p

			cm.Execute RowsUpdated
				
            Set cm = Nothing
		
			if cn.Errors.count > 0 or RowsUpdated <> 1 then
				strSuccess = "0"
				exit for
			else
				strSuccess = "1"
			end if
		next 

		if strSuccess="0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if


	    set cn = nothing
	end if

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

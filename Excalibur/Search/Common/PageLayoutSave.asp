<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
            if(navigator.appName != "Microsoft Internet Explorer" && navigator.appName != "Internet Explorer" && navigator.appName != "IE")
               if (typeof( window.parent.opener) != "undefined")
                window.parent.opener.location.reload();

			window.returnValue = "1";
			window.parent.close();
			}
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
		dim cn
		dim i
		dim ItemArray
		dim blnFailed
		dim RowsUpdated
        dim strSuccess

        blnFailed = false

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open

		cn.BeginTrans	

		set cm = server.CreateObject("ADODB.Command")
					            
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateEmployeeUserSetting2"
		cm.CommandType = &H0004
		                
		Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
		p.Value = clng(request("txtUserID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@UserSettingID", 3, &H0001)
		p.Value = 8
		cm.Parameters.Append p
		                    
		Set p = cm.CreateParameter("@Value", 200, &H0001,8000)
		p.Value = trim(left(request("txtLayout"),8000))
		cm.Parameters.Append p
		                    
		cm.Execute rowschanged

		if rowschanged <> 1 then
            blnFailed = false
		end if
	
		set cm = nothing


		set cm = server.CreateObject("ADODB.Command")
					            
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateEmployeeUserSetting2"
		cm.CommandType = &H0004
		                
		Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
		p.Value = clng(request("txtUserID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@UserSettingID", 3, &H0001)
		p.Value = 9
		cm.Parameters.Append p
		                    
		Set p = cm.CreateParameter("@Value", 200, &H0001,8000)
		p.Value = trim(left(request("txtFieldFilters"),8000))
		cm.Parameters.Append p
		                    
		cm.Execute rowschanged

		if rowschanged <> 1 then
            blnFailed = false
		end if
	
		set cm = nothing

		if not blnFailed then
			strSuccess = "1"
			cn.committrans
		else
			strSuccess = ""
			cn.rollbacktrans
		end if

		cn.Close
		set cn=nothing

        
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>


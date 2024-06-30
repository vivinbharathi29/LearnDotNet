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
    <script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
		{

			//window.returnValue = txtColumns.value;
		    //window.parent.close();
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.ReOrderColumnsResult(txtColumns.value);
		        ClosePulsarPlusPopup();
		    }
		    else 
		    {
		        window.returnValue = txtColumns.value;
		        window.parent.close();
		    }
		    
//		else
//			document.write ("<BR><font size=2 face=verdana>Unable to update order.</font>");
		}
//	else
//		document.write ("<BR><font size=2 face=verdana>Unable to update order.</font>");

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

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open

		cn.BeginTrans	

		set cm = server.CreateObject("ADODB.Command")
					            
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateDefaultProductFilter"
		cm.CommandType = &H0004
		                
		Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
		p.Value = clng(request("txtEmployeeID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@UserSettingID", 3, &H0001)
		p.Value = 7
		cm.Parameters.Append p
		                    
		Set p = cm.CreateParameter("@Value", 200, &H0001,8000)
		p.Value = trim(left(request("txtNewOrder"),8000))
		cm.Parameters.Append p
		                    
		cm.Execute rowschanged

		if rowschanged = 1 then
			strSuccess = "1"
			cn.committrans
		else
			strSuccess = ""
			cn.rollbacktrans
		end if
	
		set cm = nothing
		cn.Close
		set cn=nothing
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=txtColumns name=txtColumns value="<%=request("txtNewOrder")%>">
</BODY>
</HTML>


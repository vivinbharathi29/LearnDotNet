<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined"){
	    if (txtSuccess.value != "0") {
	        if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.popupCallBack(txtSuccess.value);
	            ClosePulsarPlusPopup();
	        }
	        else {
	            if (parent.window.parent.document.getElementById('modal_dialog')) {
	                parent.window.parent.CommodityResults(txtSuccess.value);
	                parent.window.parent.modalDialog.cancel();
	            } else {
	                window.returnValue = txtSuccess.value;
	                window.parent.close();
	            }
	        }	        
		}
	}
}



//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	Response.Write "ID: " & clng(request("txtID")) & "<BR>"
	Response.Write "PartNumber: " & left(request("txtPartNumber"),20) & "<BR>"
	Response.Write "<INPUT style=""Width:100%"" type=""text"" id=text1 name=text1 value=""" & Request.Form & """>"
	
	dim cn
	dim cm
	dim strSuccess
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	cn.BeginTrans
	strSuccess = request("txtPartNumber")


	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
			
	cm.CommandText = "spUpdateVersionPartNumber"	
	
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = clng(request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PartNumber", 200,  &H0001,20)
	p.Value = left(request("txtPartNumber"),20)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@AffectRows", 3,  &H0002)
	cm.Parameters.Append p

	cm.Execute 
	rowschanged = cm.Parameters("@AffectRows").Value
	set cm=nothing

	if rowschanged <> 1 then
		strSuccess = "0"
	end if	

	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	cn.Close
	set cn = nothing


%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

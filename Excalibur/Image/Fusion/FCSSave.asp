<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	    if (txtSuccess.value == "1") {
	        if ('<%=Request("isFromPulsarPlus")%>' != '') {
	            parent.window.parent.location.reload();;
	            parent.window.parent.closeExternalPopup();
	        }
	        else {
	            window.returnValue = 1;
	            window.parent.close();
	        }
	    }
	    else
	        document.write("<BR><font size=2 face=verdana>Unable to update FCS dates.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update FCS dates.</font>");
}



//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim TAGArray
	dim IDArray
	dim DateArray
	dim ActualArray
	dim ActualTagArray
	dim FAISKUArray
	dim FAISKUTagArray
	dim CommentArray
	dim CommentTagArray
	
	dim i
	dim cn
	dim cm
	dim strSuccess
	dim rowschanged
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	IDArray = split(request("txtIDList"),",")
	TAGArray = split(request("txtDateTag"),",")
	DateArray = split(request("txtFCS"),",")
	ActualArray = split(request("txtActual"),",")
	ActualTagArray = split(request("txtActualTag"),",")
	CommentArray = split(request("txtComments"),",")
	CommentTagArray = split(request("txtCommentTag"),",")
	FAISKUArray = split(request("txtFAISKU"),",")
	FAISKUTagArray = split(request("txtFAISKUTag"),",")
	cn.BeginTrans
	strSuccess = "1"
	for i = lbound(IDArray) to ubound(IDArray)
		if ucase(trim(DateArray(i))) <> ucase(trim(TagArray(i))) or ucase(trim(ActualArray(i))) <> ucase(trim(ActualTagArray(i))) or ucase(trim(CommentArray(i))) <> ucase(trim(CommentTagArray(i))) or ucase(trim(FAISKUArray(i))) <> ucase(trim(FAISKUTagArray(i))) then
			
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
		
			cm.CommandText = "spUpdateImageRampPlan"	

			Set p = cm.CreateParameter("@ID", 3,  &H0001)
			p.Value = IDArray(i)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@FCSDate", 135,  &H0001)
			if trim(DateArray(i)) = "" then
				p.Value = null
			else
				p.Value = cdate(DateArray(i))
			end if
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@FCSActual", 135,  &H0001)
			if trim(ActualArray(i)) = "" then
				p.Value = null
			else
				p.Value = cdate(ActualArray(i))
			end if
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@FAISKU", 200,  &H0001,20)
			p.Value = left(FAISKUArray(i),20)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Comments", 200,  &H0001,256)
			p.Value = left(CommentArray(i),256)
			cm.Parameters.Append p


			cm.Execute rowschanged

			set cm=nothing

			if rowschanged <> 1 then
				strSuccess = "0"
				exit for
			end if	
		end if
	next


	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	cn.Close
	set cn = nothing
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>

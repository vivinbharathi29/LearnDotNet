<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cmdCancel_onclick() {
        if (parent.window.parent.loadDatatodiv != undefined) {
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.close();
            }
        }
    }

function cmdOK_onclick() {
	if (frmInput.chkNotify.checked && frmInput.txtReason.value=="")
		{
		alert("Please enter the reason why this deliverable is being removed."); 
		frmInput.txtReason.focus();
		}
	else
		{
		cmdOK.disabled = true;
		cmdCancel.disabled = true;
		frmInput.submit();
		}
}


function chkNotify_onclick() {
	if (frmInput.chkNotify.checked)
		{
		divReason.style.display = "";
		frmInput.txtReason.focus();
		}
	else
		divReason.style.display = "none";
}

function window_onload() {
	frmInput.txtReason.focus();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgColor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%
	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strName

	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")

	if request("DeliverableID") = "" or request("ProductID") = "" then
		Response.Write "<font size=2 face=verdana>Not enough information provided to perform this function.</font>"
	else
		rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strProductName = ""
			Response.Write "<BR><font size=2 face=verdana>Unable to find the selected product.</font>"
		else
			strProductName = rs("Name") & ""
		end if
		rs.Close	
		rs.Open "spGetDeliverableRootName " & clng(request("DeliverableID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strName = ""
			Response.Write "<BR><font size=2 face=verdana>Unable to find the selected deliverable.</font>"
		else
			strName = rs("Name") & ""
		end if
		rs.Close	
		
		if strName <> "" and strProductName <> "" then
			Response.Write "<font face=verdana size=3><b>Remove Root Deliverable</b><BR><BR></font>"  		
			Response.write "<Font face=verdana size=2>Are you sure you want to remove " & strName & " from " & strProductName & "?</font><BR><BR>"
			Response.Write "<form id=frmInput action=""RemoveRootSave.asp"" method=post>"
			Response.Write "<table width=""100%"" border=0 cellspacing=0 cellpadding=2><tr><td><INPUT type=""checkbox"" checked id=chkNotify name=chkNotify LANGUAGE=javascript onclick=""return chkNotify_onclick()""></td><td><font face=verdana size=2>Notify Deliverable Owner of this Action.</font></td></tr>"
			Response.Write "<TR><TD>&nbsp;</TD><TD width=""100%""><div id=divReason><font size=2 face=verdana><b>Reason For Removal <font color=red>*</font><BR></B></font>"
			Response.Write "<INPUT type=""Text"" style=""width:100%"" id=txtReason name=txtReason value=""""></div></TD></TR></TABLE>"
			Response.Write "<INPUT type=""hidden"" id=txtDeliverableID name=txtDeliverableID value=""" & request("DeliverableID") & """>"
			Response.Write "<INPUT type=""hidden"" id=txtProductID name=txtProductID value=""" & request("ProductID") & """>"
			Response.Write "<INPUT type=""hidden"" id=txtProductName name=txtProductName value=""" & strProductName & """>"
			Response.Write "</form>"
			Response.Write "<HR>"
			Response.Write "</font>"
			Response.Write "<table width=100% ><TR>"
			Response.Write "<TD align=right><INPUT type=""button"" value=""Yes"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cmdOK_onclick()"">&nbsp;<INPUT type=""button"" value="" No "" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick=""return cmdCancel_onclick()""></TD>"
			Response.Write "</TR></table>"
		end if
	end if
	

	set rs = nothing
	cn.Close
	set cn = nothing
%>
</BODY>
</HTML>

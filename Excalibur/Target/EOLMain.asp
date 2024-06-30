<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
	if (frmInput.optChoose[1].checked && frmInput.cboRoot.selectedIndex==0)
		alert("You must select one option to continue.");
	else
		frmInput.submit();
}

function optRemove_onclick() {
	ChooseRoot.style.display="none";
}

function optReplace_onclick() {
	ChooseRoot.style.display="";
}

//-->
</SCRIPT>
</HEAD>
<BODY bgColor=Ivory>

<%
	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strExceptions
	dim strOOC
	dim strName

	strExceptions = ""
	strOOC = ""
	
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
		rs.Open "spGetDeliverableRootName " & clng(request("DeliverableID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<font size=2 face=verdana>Unable to find the selected deliverable.</font>"
			rs.Close
		else
			strName = rs("Name") & ""
			rs.Close
			Response.Write "<font face=verdana size=3><b>" & strName & " is inactive.</b><BR><BR></font>"  		
			Response.Write "<u><b>Options:</b></u><BR><BR><font size=2 face=verdana>"
			Response.Write "<form id=frmInput action=""EOLSave.asp"" method=post>"
			Response.Write "<INPUT type=""radio"" value=1 id=optChoose name=optChoose LANGUAGE=javascript onclick=""return optRemove_onclick()"">&nbsp;Remove deliverable from this product.<BR>"
			Response.Write "<INPUT type=""radio"" value=2 id=optChoose name=optChoose LANGUAGE=javascript onclick=""return optReplace_onclick()"">&nbsp;Replace with a different deliverable.<BR>"
			Response.Write "<div style=""display:none;"" ID=ChooseRoot>&nbsp;&nbsp;&nbsp;&nbsp;<SELECT style=""width:600"" id=cboRoot name=cboRoot><OPTION></OPTION>"
			rs.Open "spGetDelRoot",cn,adOpenForwardOnly
			do while not rs.EOF
				if trim(rs("ID")) <> trim(request("DeliverableID")) then
					Response.Write  "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
				end if
				rs.MoveNext
			loop
			rs.Close
			Response.Write "</SELECT><BR><BR></div>"
			'Response.Write "<INPUT type=""hidden"" id=txtDeliverableID name=txtDeliverableID value=""" & request("DeliverableID") & """>"
			'Response.Write "<INPUT type=""hidden"" id=txtProductID name=txtProductID value=""" & request("ProductID") & """>"
			Response.Write "</form>Note: Contact the developer if you would like to request to have this deliverable reactivated.<BR><BR>"
			Response.Write "<HR>"
			Response.Write "</font>"
			Response.Write "<table width=100% ><TR>"
			Response.Write "<TD align=right><INPUT type=""button"" value=""OK"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cmdOK_onclick()"">&nbsp;<INPUT type=""button"" value=""Cancel"" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick=""return cmdCancel_onclick()""></TD>"
			Response.Write "</TR></table>"
		end if
	end if
	

	set rs = nothing
	cn.Close
	set cn = nothing
%>

</BODY>
</HTML>

<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function instr(MyString,Find){
	return MyString.substr(0,MyString.indexOf(Find,0));
	
}

function instrAt(MyString,Find){
	return MyString.indexOf(Find,0);
	
}

function midstr(MyString,Find){
	return MyString.substr(MyString.indexOf(Find,0)+1);
	
}


function window_onload() {
	if (typeof(Main.txtCode) != "undefined")
		Main.txtCode.focus();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font size=4 face=verdana><b>Enter New Supplier Code</b></font><font size=1><BR></font>


<form id="Main" method="post" action="SupplierCodeSave.asp">
<%

	dim strSQL
	dim rs
	dim cn
	dim strCategory
	dim strVendor
	
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open 
	
	rs.Open "spGetDeliverableCategoryName " & clng(request("CategoryID")),cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		strCategory = rs("name") & ""
	end if
	rs.Close
	rs.Open "spGetVendorName " & clng(request("VendorID")),cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		strVendor = rs("name") & ""
	end if
	rs.Close
		
	if strCategory = "" or strVendor = "" then
		Response.Write "<BR>Not enough information supplied."
	else
%>		
		<table  WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<TR>
			<TD><b>Category:</b></TD>
			<TD><font size=2 face=verdana><%=strCategory%></font></TD>
		</TR>
		<TR>
			<TD><b>Vendor:</b></TD>
			<TD><font size=2 face=verdana><%=strVendor%></font></TD>
		</TR>
		<TR>
			<TD><b>Supplier&nbsp;Code:</b></TD>
			<TD><INPUT type="text" id=txtCode name=txtCode maxlength=50></TD>
		</TR>
		</table>		
				
<%		
	end if
	
	set rs = nothing
	cn.Close
	set cn = nothing
%>
<INPUT type="hidden" id=txtCategoryID name=txtCategoryID value="<%=request("CategoryID")%>">
<INPUT type="hidden" id=txtVendorID name=txtVendorID value="<%=request("VendorID")%>">
</form>

</BODY>
</HTML>

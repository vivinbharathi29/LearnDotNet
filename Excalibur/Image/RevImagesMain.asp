<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdGO_onclick() {
	var i;

	if (typeof (frmMain.chkImage.length) != "undefined") {
	    for (i = 0; i < frmMain.chkImage.length; i++) {
	        if (!frmMain.chkImage[i].checked)
	            frmMain.txtSKU[i].value = ""; //Clear out ones that are not checked
	        else
	            if (frmMain.txtSKU[i].value.indexOf("\t") == -1)
	                frmMain.txtSKU[i].value = frmMain.chkImage[i].value + "\t" + frmMain.txtSKU[i].value
	        }
	    }
	    else {
	        if (!frmMain.chkImage.checked)
	            frmMain.txtSKU.value = ""; //Clear out ones that are not checked
	        else
	            if (frmMain.txtSKU.value.indexOf("\t") == -1)
	                frmMain.txtSKU.value = frmMain.chkImage.value + "\t" + frmMain.txtSKU.value
	        }	
	frmMain.submit();
}

function chkAll_onclick() {
    if (typeof (frmMain.chkImage.length) != "undefined")
        for (i = 0; i < frmMain.chkImage.length; i++)
            frmMain.chkImage[i].checked = frmMain.chkAll.checked;
    else
        frmMain.chkImage.checked = frmMain.chkAll.checked;
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
TD{
	FONT-SIZE:xx-small;
	FONT-FAMILY: Verdana;
}
</STYLE>
<BODY>
<font size=3 face=verdana><b>Rev Images</b></font><BR>

<form ID=frmMain action="RevImagesSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>

<INPUT type="button" value="Clone Selected Images" id=cmdGO name=cmdGO LANGUAGE=javascript onclick="return cmdGO_onclick()">

<%

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")



	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	
	if (rs.EOF and rs.BOF) then
		set rs = nothing
       	set cn=nothing
       	Response.Redirect "../NoAccess.asp?Level=1"
	else
		CurrentUserID = rs("ID") & ""
	end if
	rs.Close




	dim ProductID
	ProductID = clng(request("ProductID"))'253
	rs.Open "spListImageDefinitionsByProduct " & ProductID & "",cn,adOpenStatic
	Response.Write "<Table bgcolor=ivory width=100% cellspacing=0 cellpadding=2 border=1>"
	Response.Write "<TR><TD bgcolor=beige><INPUT type=""checkbox"" id=chkAll name=chkAll  LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD bgcolor=beige><b>ID</b></TD><TD bgcolor=beige><b>Old SKU</b></TD><TD bgcolor=beige><b>New SKU</b></TD><TD bgcolor=beige><b>Brand</b></TD><TD bgcolor=beige><b>OS</b></TD><TD bgcolor=beige><b>SW</b></TD><TD bgcolor=beige><b>ImageType</b></TD><TD bgcolor=beige><b>Status</b></TD><TD bgcolor=beige><b>Comments&nbsp;</b></TD></TR>"
	do while not rs.EOF
		NewSKU = rs("Skunumber")
		if NewSKU <> "" then
			NewSKU = left(NewSKU, len(NewSKU)-1) & Right(NewSKU,1) + 1
		end if
		Response.Write "<TR>"
		Response.Write "<TD><INPUT type=""checkbox"" id=chkImage name=chkImage value=""" & rs("ID") & """></TD>"
		Response.Write "<TD>" & rs("ID") & "</TD>"
		Response.Write "<TD>" & rs("Skunumber") & "&nbsp;</TD>"
		Response.Write "<TD><INPUT style=""WIDTH:100"" ID=""SKU" & rs("ID") & """ type=""text"" maxlength=20 id=txtSKU name=txtSKU value=""" & NewSKU & """></TD>"
		Response.Write "<TD>" & rs("Brand") & "</TD>"
		Response.Write "<TD>" & rs("OS") & "</TD>"
		Response.Write "<TD>" & rs("SWType") & "</TD>"
		Response.Write "<TD>" & rs("ImageType") & "</TD>"
		Response.Write "<TD>" & rs("Status") & "</TD>"
		Response.Write "<TD>" & rs("Comments") & "&nbsp;</TD>"
		Response.Write "</TR>"
		rs.MoveNext
	loop
	Response.Write "</Table>"
	rs.Close
		
	set rs = nothing
	cn.Close
	set cn = nothing



%>
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=ProductID%>">
</form>
</BODY>
</HTML>

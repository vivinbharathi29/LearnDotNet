<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function DisplayRoot(strID) {
	var rc;
	rc=window.showModalDialog("../root.asp?ID=" + strID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No"); 
	if (typeof(rc) != "undefined")
		window.location.reload();
}

//-->
</SCRIPT>
</HEAD>
<BODY ><LINK rel="stylesheet" type="text/css" href="../style/general.css">

<%
	dim cn
	dim rs
	dim strLastID
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	strLastID = ""
	rs.open "spListDeliverableTitles",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(strLastID) <> trim(rs("RootID")& "") then
			if strLastID <> "" then
				Response.Write "</table></td></tr><tr><td class=WhiteRow colspan=3>&nbsp;</td></tr>"
			else
				Response.Write "<table class=HighlightRow width=""100%"" >"
			end if
			Response.Write "<TR><TD><b>Name:</b></td><td>"
			Response.Write "<a href=""javascript:DisplayRoot(" & rs("RootID") & ")"">" & rs("Name") & "</a></td></tr>"
			Response.Write "<TR><TD><b>Description:</b></td><td>" & rs("Description") & "</td></tr>"
			Response.Write "<TR><TD><b>Manager:</b></td><td>" & rs("Devmanager") & "</td></tr>"
			Response.Write "<TR><TD><b>Translations:</b></td><td><TABLE width=""100%""><THEAD><TH>Language</TH><TH>Short Name</TH><TH>Short Description</TH></THEAD>"
			strLastID = rs("RootID")
		end if
		Response.Write "<TR>"		
		Response.Write "<TD>" & rs("Abbreviation") & "</TD>"		
		Response.Write "<TD>" & rs("Title") & "</TD>"		
		Response.Write "<TD>" & rs("ShortDescription") & "</TD>"		
		Response.Write "</TR>"				
		rs.movenext
	loop
	rs.Close
	Response.Write "</table>"


	set rs = nothing
	set cn = nothing

%>

</BODY>
</HTML>

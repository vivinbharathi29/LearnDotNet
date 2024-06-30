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

    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }

//-->
</SCRIPT>
</HEAD>
<LINK rel="stylesheet" type="text/css" href="../../style/programoffice.css">
<BODY bgcolor=Ivory>

<%
	dim cn
	dim rs
	dim blnFound
	dim strID
	dim strDeliverable
	dim strPartNumber

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	if request("VersionID") = "" then
		Response.Write "Not enough information supplied to process your request."
	else
		blnFound=false
		rs.Open "spGetVersionPartNumber " & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strPartNumber = ""
		else
			blnFound = true
			strPartNumber = rs("PartNumber") & ""
		end if
		strDeliverable = rs("DeliverableName") & "<BR><b>HW:</b> " & rs("Version") & "<BR><b>FW:</b> " & rs("Revision") & "<BR><b>Vendor:</b> " & rs("Vendor") 
		
	
		rs.Close
	end if

	if blnFound then
		Response.Write "<form ID=frmMain action=""PartNumberSave.asp"" method=post>"	
		Response.Write "<b>ID: </b>" & clng(request("VersionID")) & "<BR>"
		Response.Write "<b>Deliverable: </b>" & strDeliverable & "<BR>"

		Response.Write "<b>Part Number:&nbsp;</b><INPUT type=""text"" id=txtPartNumber name=txtPartNumber value=""" & strPartNumber & """>" 
	
		Response.Write "<INPUT type=""hidden"" id=txtID name=txtID value=""" & clng(request("VersionID")) & """>"
	
		Response.write "<hr>"
		Response.Write "<table border=0 width=""100%""><tr><td align=right><INPUT type=""submit"" value=""OK"" id=cmdOK name=cmdOK>&nbsp;<INPUT type=""Button"" value=""Cancel"" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick=""return cmdCancel_onclick()""></td></tr></table>"
		Response.Write "</form>"
	end if


	cn.Close
	set rs = nothing
	set cn=nothing
%>
</BODY>
</HTML>

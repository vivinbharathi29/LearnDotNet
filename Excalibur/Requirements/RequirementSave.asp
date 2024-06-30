<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->
<DOCTYPE html>
<HTML>
<HEAD>
<title>Requirement</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../includes/client/jquery.min.js"></script>
<script type="text/javascript" src="../includes/client/json2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array();
	if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
		    OutArray[0] = txtSpec.value;
		    OutArray[1] = txtDel.value;
		    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
		        // For Reload PulsarPlusPmView Tab
		        parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

		        // For Closing current popup
		        parent.window.parent.closeExternalPopup();
		    }
		    else {
		        var iframeName = parent.window.name;
		        if (iframeName != '') {
		            parent.window.parent.CloseRequirementsDialog(OutArray);
		        } else {
		            if (parent.window.parent.document.getElementById('modal_dialog')) {
		                //save array value and return to parent page: ---
		                parent.window.parent.modalDialog.passArgument(JSON.stringify(OutArray), 'requirement_save_array');
		                parent.window.parent.requirementrowsResults();
		                parent.window.parent.modalDialog.cancel();
		            } else {
		                window.parent.returnValue = OutArray;
		                window.parent.close();
		            }
		        }
		    }
        } else {
		    document.write("<BR><BR>Unable to update requirement.  An unexpected error occurred.");
		}
	} else {
        document.write ("<BR><BR>Unable to update requirement.  An unexpected error occurred.");
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%


Function StripHTMLTag(ByVal sText)
   StripHTMLTag = ""
   fFound = False
   Response.Write sText & "<BR>" & vbcrlf
   Do While InStr(sText, "<")
      fFound = True
      StripHTMLTag = StripHTMLTag & " " & Left(sText, InStr(sText, "<")-1)
      strTag = lcase(trim(mid(sText,InStr(sText, "<"),InStr(sText, ">") - InStr(sText, "<")+1)))

'	  if strTag = "<b>" or strTag = "</b>" or strTag = "<i>" or strTag = "</i>" or strTag = "<u>" or strTag = "</u>" then
		if left(replace(ucase(strTag)," ",""),5) <> "<" & trim("FONT") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("FONT") and left(replace(ucase(strTag)," ",""),5) <> "<" & trim("SPAN") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("SPAN") and left(replace(ucase(strTag)," ",""),4) <> "<" & trim("DIV") and left(replace(ucase(strTag)," ",""),5) <> "</" & trim("DIV") and left(replace(ucase(strTag)," ",""),2) <> "<" & trim("P") and left(replace(ucase(strTag)," ",""),3) <> "</" & trim("P") then
			StripHTMLTag = StripHTMLTag & strTag
      end if

	  
      sText = MID(sText, InStr(sText, ">") + 1)

      
   Loop
   StripHTMLTag = StripHTMLTag & sText
   If Not fFound Then StripHTMLTag = sText
End Function

	dim cn
	dim cm
	dim p
	dim rowschanged
	dim strSuccess
	dim FoundErrors
	dim strSpecification
	dim strDelList
	dim strDeliverablesAdded
	dim strDeliverablesRemoved
	dim rs
	set rs = server.CreateObject("ADODB.Recordset")
	
	strSpecification = request("txtSpecification")
'	if lcase(left(strSpecification,26)) = "<font face=verdana size=1>" then
'		strSpecification = mid(strSpecification,27)
'	end if
'	if lcase(right(strSpecification,7)) = "</font>" then
'		strSpecification = left(strSpecification,len(strSpecification)-7)
'	end if
	strSpecification = stripHTMLTag(strSpecification)

	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	cn.BeginTrans

	FoundErrors = false	

	if request("tagSpecification") <>  request("txtSpecification") then
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn	
		cm.CommandText = "spUpdateRequirementDefinitionWeb"	
	
		Set p = cm.CreateParameter("@ID", 3,  &H0001)
		p.Value = clng(request("txtDisplayedID"))
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@ID", 3,  &H0001)
		p.Value = clng(request("txtProdID"))
		cm.Parameters.Append p


		Set p = cm.CreateParameter("@Definition", 201, &H0001, 2147483647)
		p.value = strSpecification
		cm.Parameters.Append p
	
		cm.Execute rowschanged

		if rowschanged <> 1 then
			FoundErrors = true
		end if
		
		set cm = nothing
	end if
	
	if FoundErrors then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.committrans
		strSuccess = "1"	
	end if
	
	strDelList = "<table width=100% border=1 cellspacing=0 cellpadding=2>"
	rs.Open "spListDeliverableRoots4Requirements " & clng(request("txtPRID")),cn,adOpenForwardOnly
	do while not rs.eof
		if not isnull(rs("reqid")) then
'			strDelList = strDelList & "-" & rs("Name") & "<BR>"
            strDelList = strDelList & "<TR><TD><font size=1>" & rs("Name") &  "</font></TD></TR>"
		end if
		rs.movenext
	loop	
	strDelList = strDelList & "</table>"

	rs.close
	set rs = nothing
	if strDelList = "" then
		strDelList = "&nbsp;"
	end if

	
	set cn = nothing


	if strSpecification = "" then
		strSpecification = "&nbsp;"
	end if
%>
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=txtID name=txtID value="<%=request("txtDisplayedID")%>">
<TEXTAREA rows=2 cols=20 id=txtSpec name=txtSpec style="Display:none"><%=strSpecification%></TEXTAREA>
<TEXTAREA rows=2 cols=20 id=txtDel name=txtDel style="Display:none"><%=strDelList%></TEXTAREA>

</BODY>
</HTML>

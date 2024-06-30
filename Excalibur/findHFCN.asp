<%@ Language=VBScript %>
<HTML>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function IsNumeric(sText)
{
   var ValidChars = "0123456789.";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }


function window_onload() {
	SearchBar.innerHTML = txtSearchBar.value;
	txtHFCNID.focus();
}

function Goto(DelID){
	window.location.href = "dmview.asp?ID=" + DelID;
}

function goID_onmouseover() {
		window.event.srcElement.style.cursor = "hand"
}

function txtHFCNID_onkeypress() {
	if (window.event.keyCode == 13)
		goID_onclick();
}
function goID_onclick() {
	if (txtHFCNID.value != "")
		{
		if (! IsNumeric(txtHFCNID.value))
			{
			alert("You may only enter numbers into the HFCN ID field.");
			txtHFCNID.focus();
			}
		else
			{
			strID = window.showModalDialog("WizardFrames.asp?Type=1&ID=" + txtHFCNID.value,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
			if (typeof(strID) != "undefined")
				{
				txtHFCNID.value = "";
				txtHFCNID.focus();
				}
			}
		}
	else
		{
		window.alert("Please enter search criteria first.");
		txtHFCNID.focus();
		}
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<FONT face=verdana>
<H3>Find Deliverable</H3> 
</FONT>
<H4>Search:  </H4>
<TABLE><TR><TD><font size=2 face=verdana>HFCN&nbsp;ID:&nbsp;</font></TD><TD><INPUT type="text" id=txtHFCNID name=txtHFCNID LANGUAGE=javascript onkeypress="return txtHFCNID_onkeypress()">&nbsp;<img id="goID" border="0" src="images\go.gif" WIDTH="23" HEIGHT="20" LANGUAGE=javascript onclick="return goID_onclick()" onmouseover="return goID_onmouseover()"></a></TD></TR></TABLE>


<HR>
<H4>Browse:  </H4>
<P>
<Label ID=SearchBar></Label><BR>
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%
	dim txtSearch
	dim lastLetter
	txtSearch = "<TABLE borderColor=tan cellSpacing=1 cellPadding=1 border=1><TR bgColor=ivory>"
  
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetRootDeliverables",cn,adOpenForwardOnly
	
	lastLetter = ""
	do while not rs.EOF
		if rs("RootFilename") = "HFCN" then
			if lcase(lastletter) <> lcase(replace(left(rs("Name"),1),"""","&quot;")) then
				lastLetter = replace(left(rs("Name"),1),"""","&quot;")
				Response.Write "<TR><a name=" & lastLetter & "><TD colspan=2 nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>" & lastletter & "</FONT></STRONG></TD></TR>"
				txtsearch = txtsearch & "<TD width=5 align=middle><font size=2 face=verdana><a HREF=#" & lastletter & ">" & lastLetter & "</font></TD>"
			end if
			Response.Write "<TR><TD><FONT face=verdana size=1><a href=""javascript:Goto(" & rs("ID") & ")"">" & rs("Name") & "</a></FONT></TD></TR>"
		end if
		rs.MoveNext
	loop
	Response.Write "</Table>"
%>	  
    
    <% txtSearch = txtSearch & "</TR></TABLE>"
    %>
  <INPUT style="Display:none" type="text" id=txtSearchBar name=txtSearchBar value ="<%=txtSearch%>">
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
</font>
</BODY>
</HTML>

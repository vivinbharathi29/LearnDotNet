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
<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Sustaining Product Test</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdTargetDate_onclick() {
	var strID;
	strID = window.showModalDialog("../../mobilese/today/calDraw1.asp",frmUpdate.txtTargetDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			frmUpdate.txtTargetDate.value = strID;
		}
}

//-->
</SCRIPT>
</HEAD>
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY bgcolor=Ivory>
<font face=verdana size=><b>
<label ID="lblTitle">

Update test deliverable status

<%
	dim cn
	dim rs
	dim strTargetDate
	dim strComments

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	strTargetDate="TBD"
%>

<form id="frmUpdate" method="post" action="TestRequestSave.asp?RootID=<%=Request("RootID")%>&VersionID=<%=Request("VersionID")%>">

<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr><td valign=top width=110 nowrap><b>&nbsp;Root:</b></td>
	<td>
	<%
		rs.Open "spGetDeliverableRootName " & request("RootID") ,cn,adOpenForwardOnly		if not rs.EOF then			Response.Write "<OPTION value=""" &request("RootID") & """>" & rs("Name") & "</OPTION>"		else			Response.Write "<OPTION value=0>No root has been selected</OPTION>"
		end if		rs.Close	%>	</td>	</tr>

	<tr><td valign=top width=110 nowrap><b>&nbsp;Version:</b></td>
	<td>
	<%
		rs.open "spGetDeliverableVersionProperties " & request("VersionID") ,cn,adOpenForwardOnly
		if not rs.EOF then
			Response.Write "<Option selected value=""" & request("VersionID") & """>" & rs("Version") & ", " & rs("Revision") & ", " & rs("Pass") & "</Option>"
		end if
		rs.Close
	%>
	</td>
	</tr>

	<tr><td valign=top width=110 nowrap><b>&nbsp;Product:</b></td>
	<td>
	<TABLE ID=TableProduct Name=TableProduct width=100%>
	<THEAD bgcolor=LightSteelBlue  ><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Support&nbsp;</TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Product&nbsp;</TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Street Name&nbsp;</TD></THEAD>
	<%
		rs.open "spGetProductList4Root " & request("RootID") ,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then 
			Response.Write "<TR><TD nowrap>No product has picked up this deliverable</TD></TR>"
	    else 
			ProductIDs = ""
			dim strStreetNames
			dim SeriesArray
			dim strSeriesText
	
			strStreetNames = ""
			
			do while not rs.EOF
				ProductIDs = ProductIDs & rs("ID") & "*"
				if trim(rs("PostRTMStatus")) = "0" then
					Response.Write "<TR><TD nowrap>No</TD>"
				elseif trim(rs("PostRTMStatus")) = "1" then
					Response.Write "<TR><TD nowrap>Yes</TD>"
				end if	
			
				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spListBrands4Product " & rs("ID"),cn,adOpenForwardOnly
				do while not rs2.EOF
					if trim(rs2("SeriesSummary") & "") <> ""  then
						SeriesArray = split(rs2("SeriesSummary"),",")
						for each strSeriesText in SeriesArray
							strStreetNames = strStreetNames & "<BR>" & rs2("StreetName2")  & " " & strSeriesText 
						next
					end if	
					rs2.MoveNext
				loop		
				rs2.Close
				set rs2 = nothing
				if trim(strStreetNames) = "" then
					strStreetNames = "TBD"
				end if
				
				Response.Write "<TD nowrap>" & rs("Name") & " " & rs("Version") & "</TD>"
				Response.Write "<TD>XXX" & strStreetNames & "</TD></TR>"
				rs.MoveNext
			loop
		end if
		rs.Close	
	%>
	</TABLE>
	</td>
	</tr>	

	<tr id=StatusRow name=StatusRow><td valign=top width=110 nowrap><b>&nbsp;Status:</b></td>
	<td>
	<div id=StatusParent name=StatusParent>
	<SELECT style="width=100%" id=cboVersionStatus name=cboVersionStatus>
	<%
		rs.open "spGetDeliverableVersionProperties " & request("VersionID") ,cn,adOpenForwardOnly
		if not rs.EOF then
			if rs("PostRTMStatus") = "2" then
				Response.Write "<OPTION selected value=2>In Test</OPTION>"
				Response.Write "<OPTION value=3>Passed</OPTION>"
				Response.Write "<OPTION value=4>Failed</OPTION>"
				Response.Write "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			else
				Response.Write "<OPTION value=2>In Test</OPTION>"
				Response.Write "<OPTION value=3>Passed</OPTION>"
				Response.Write "<OPTION value=4>Failed</OPTION>"
				Response.Write "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			end if
		end if
		rs.Close
	%>
	</SELECT>
	</div>
	</td>
	</tr>
	
	<%
	if request("VersionID") <> "" and request("VersionID") <> "0" then
		rs.open "spGetDeliverableVersionProperties " & request("VersionID") ,cn,adOpenForwardOnly
		if not rs.EOF then
			strTargetDate=rs("PostRTMTargetDate") & ""
		end if
		rs.Close
	end if
	%>

	<tr><td valign=top width=110 nowrap><b>&nbsp;Target Date:</b></td>
	<td>
		<INPUT type="text" style="width:170" id=txtTargetDate name=txtTargetDate value="<%=strTargetDate%>">
		<a href="javascript: cmdTargetDate_onclick()"><img ID="picTarget" SRC="../../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
	</td>
	</tr>
		<%
	if request("VersionID") <> "" and request("VersionID") <> "0" then
		rs.open "spGetDeliverableVersionProperties " & request("VersionID") ,cn,adOpenForwardOnly
		if not rs.EOF then
			strComments=rs("PostRTMComments") & ""
		end if
		rs.Close
	end if
	%>

	<tr><td valign=top width=110 nowrap><b>&nbsp;Comments:</b></td>
	<td>
		<INPUT type="text" style="width=100%" id=txtComments name=txtComments disabled value="<%=strComments%>">
	</td>
	</tr>
			
</table>
</form>

<%
	set rs=nothing
	cn.Close
	set cn=nothing
%>

</BODY>
</HTML>

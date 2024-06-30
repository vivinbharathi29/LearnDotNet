<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdGenerate_onclick() {
	frmMain.submit();
}

function cmdStartOver_onclick() {
	window.location.reload ("EmailListGenerator.asp");
}

function cmdCopy_onclick() {
	var MyField=eval("frmMain.txtEmailList")
	MyField.focus();
	MyField.select();
//	MyField.createTextRange().execCommand("Copy");
	window.clipboardData.setData('text',frmMain.txtEmailList.value);
	
	
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
Body{
	Font-Size: x-small;
	Font-Family: Verdana;
}
TD{
	Font-Size: x-small;
	Font-Family: Verdana;
}
</STYLE>

<BODY bgcolor=Ivory>
<!--<%
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

	if CurrentUser = "auth\dwhorton" or CurrentUser = "auth\wajohnson" or CurrentUser = "auth\lyoung" or CurrentUser = "auth\malichi" or CurrentUser = "auth\carrolld" or CurrentUser = "auth\aa6iy" or CurrentUser = "auth\emartinez"  or CurrentUser = "auth\thtran" then
%>-->
<font size=3 face=verdan><b>Email List Generator</b><BR><BR></font>
<form id=frmMain action=EmailListGenerator.asp method=post>
<%if Request.Form.Count = 0 then%>

	<b>Select Teams to Notify:<BR></b>
	<TABLE width=100%>
		<TR>
			<TD>Primary System Team:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerSystemTeam name=chkConsumerSystemTeam> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialSystemTeam name=chkCommercialSystemTeam> Commercial</TD>	
		</TR>
		<TR>
			<TD>Extended System Team:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerExtendedSystemTeam name=chkConsumerExtendedSystemTeam> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialExtendedSystemTeam name=chkCommercialExtendedSystemTeam> Commercial</TD>	
		</TR>
		<TR>
			<TD>Super-Users:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerSuperusers name=chkConsumerSuperusers> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialSuperusers name=chkCommercialSuperusers> Commercial</TD>	
		</TR>
		<TR>
			<TD>SE PMs:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerSEPMs name=chkConsumerSEPMs> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialSEPMs name=chkCommercialSEPMs> Commercial</TD>	
		</TR>
		<TR>
			<TD>SW Development Team:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerSWDevelopers name=chkConsumerSWDevelopers> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialSWDevelopers name=chkCommercialSWDevelopers> Commercial</TD>	
		</TR>
		<TR>
			<TD>HW Development Team:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkConsumerHWDevelopers name=chkConsumerHWDevelopers> Consumer</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkCommercialHWDevelopers name=chkCommercialHWDevelopers> Commercial</TD>	
		</TR>
		<TR>
			<TD>Active HP Users:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersAmericas name=chkUsersAmericas> Americas</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersAsiapacific name=chkUsersAsiapacific> Asiapacific</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersEMEA name=chkUsersEMEA> EMEA</TD>	
		</TR>
		<TR>
			<TD>ODM Users:</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersInventec name=chkUsersInventec> Inventec</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersCompal name=chkUsersCompal> Compal</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersQuanta name=chkUsersQuanta> Quanta</TD>	
		</TR>
		<TR>
			<TD>&nbsp;</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersWistron name=chkUsersWistron> Wistron</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersFoxconn name=chkUsersFoxconn> Foxconn</TD>	
			<TD><INPUT value=1 type="checkbox" id=chkUsersModus name=chkUsersModus> Modus</TD>	
		</TR>
		<TR>
			<TD>&nbsp;</TD>	
			<TD><INPUT value=1 type="checkbox" id=Checkbox1 name=chkUsersFlextronics> Flextronics</TD>	
		</TR>

	</Table>

	<HR>

	<TABLE width=100%>
	<TR>
		<TD align=right><INPUT type="button" value="Generate List" id=cmdGenerate name=cmdGenerate LANGUAGE=javascript onclick="return cmdGenerate_onclick()"></TD>
	</TR>
	</TABLE>

<%else

	dim strSQL
	dim strSelected
	dim strSystemTeamQuery
	dim strSuperuserQuery
	dim strSEPMQuery
	dim strExtendedQuery

	dim strSWDevelopmentTeamQuery
	dim strHWDevelopmentTeamQuery
		
	strSystemTeamQuery = ""
	strSuperuserQuery = ""
	strSEPMQuery = ""
	strExtendedQuery = ""
	strSWDevelopmentTeamQuery = ""
	strHWDevelopmentTeamQuery = ""
	
	strSelected  =""
	if request("chkConsumerSystemTeam") = 1 and request("chkCommercialSystemTeam") = 1 then
		strSelected = strSelected & ", All System Teams"  
		strSystemTeamQuery = " (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id or p.ComMarketingid=e.id or p.SMBMarketingid=e.id or p.ConsMarketingid=e.id or p.SupplyChainid=e.id or p.platformDevelopmentid=e.id or p.Serviceid=e.id or p.PDEid=e.id or p.Financeid=e.id) "
	elseif request("chkConsumerSystemTeam") =1  then
		strSelected = strSelected & ", Consumer System Team"  
		strSystemTeamQuery = " ( (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id or p.ComMarketingid=e.id or p.SMBMarketingid=e.id or p.ConsMarketingid=e.id or p.SupplyChainid=e.id or p.platformDevelopmentid=e.id or p.Serviceid=e.id or p.PDEid=e.id or p.Financeid=e.id) "
		strSystemTeamQuery = strSystemTeamQuery & " and p.devcenter=2 ) "
	elseif request("chkCommercialSystemTeam") =1 then
		strSelected = strSelected & ", Commercial System Team"  
		strSystemTeamQuery = " ( (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id or p.ComMarketingid=e.id or p.SMBMarketingid=e.id or p.ConsMarketingid=e.id or p.SupplyChainid=e.id or p.platformDevelopmentid=e.id or p.Serviceid=e.id or p.PDEid=e.id or p.Financeid=e.id) "
		strSystemTeamQuery = strSystemTeamQuery & " and p.devcenter<>2 ) "
	end if

	if request("chkConsumerExtendedSystemTeam") = 1 and request("chkCommercialExtendedSystemTeam") = 1 then
		strSelected = strSelected & ", All Extended System Teams"  
		strSystemTeamQuery = " (p.AccessoryPMid=e.id or p.PCid=e.id or p.PINPM=e.id or p.BIOSLeadID=e.id or p.ProcessorPMid=e.id or p.MarketingOPSid=e.id or p.SEPE=e.id or p.SETestLead=e.id or p.CommHWPMid=e.id or p.VideoMemoryPMid=e.id or p.GraphicsControllerPMid=e.id ) "
	elseif request("chkConsumerExtendedSystemTeam") =1  then
		strSelected = strSelected & ", Consumer Extended System Team"  
		strSystemTeamQuery = " ( (p.AccessoryPMid=e.id or p.PCid=e.id or p.PINPM=e.id or p.BIOSLeadID=e.id or p.ProcessorPMid=e.id or p.MarketingOPSid=e.id or p.SEPE=e.id or p.SETestLead=e.id or p.CommHWPMid=e.id or p.VideoMemoryPMid=e.id or p.GraphicsControllerPMid=e.id) "
		strSystemTeamQuery = strSystemTeamQuery & " and p.devcenter=2 ) "
	elseif request("chkCommercialExtendedSystemTeam") =1 then
		strSelected = strSelected & ", Commercial Extended System Team"  
		strSystemTeamQuery = " ( (p.AccessoryPMid=e.id or p.PCid=e.id or p.PINPM=e.id or p.BIOSLeadID=e.id or p.ProcessorPMid=e.id or p.MarketingOPSid=e.id or p.SEPE=e.id or p.SETestLead=e.id or p.CommHWPMid=e.id or p.VideoMemoryPMid=e.id or p.GraphicsControllerPMid=e.id) "
		strSystemTeamQuery = strSystemTeamQuery & " and p.devcenter<>2 ) "
	end if

	if request("chkConsumerSuperusers") = 1 and request("chkCommercialSuperusers") = 1 then
		strSelected = strSelected & ", All Super-users"  
		strSuperuserQuery = " (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id ) "
	elseif request("chkConsumerSuperusers") =1  then
		strSelected = strSelected & ", Consumer Super-users"  
		strSuperuserQuery = " ( (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id ) "
		strSuperuserQuery = strSuperuserQuery & " and p.devcenter=2 ) "
	elseif request("chkCommercialSuperusers") =1 then
		strSelected = strSelected & ", Commercial Super-users"  
		strSuperuserQuery = " ( (p.smid=e.id or p.pmid=e.id or p.sepmid=e.id or p.tdccmid=e.id ) "
		strSuperuserQuery = strSuperuserQuery & " and p.devcenter<>2 ) "
	end if

	if request("chkConsumerSEPMs") = 1 and request("chkCommercialSEPMs") = 1 then
		strSelected = strSelected & ", All SEPMs"  
		strSEPMQuery = " (p.sepmid=e.id ) "
	elseif request("chkConsumerSEPMs") =1  then
		strSelected = strSelected & ", Consumer SEPMs"  
		strSEPMQuery = " ( (p.sepmid=e.id ) "
		strSEPMQuery = strSEPMQuery & " and p.devcenter=2 ) "
	elseif request("chkCommercialSEPMs") =1 then
		strSelected = strSelected & ", Commercial SEPMs"  
		strSEPMQuery = " ( (p.sepmid=e.id ) "
		strSEPMQuery = strSEPMQuery & " and p.devcenter<>2 ) "
	end if


	strSWDevelopmentTeamQuery = ""
	if request("chkConsumerSWDevelopers") = 1 and request("chkCommercialSWDevelopers") = 1 then
		strSelected = strSelected & ", All SW Development Teams"  
		strSWDevelopmentTeamQuery = " ( r.typeid <> 1 )"
	elseif request("chkConsumerSWDevelopers") =1  then
		strSelected = strSelected & ", Consumer SW Development Teams"  
		strSWDevelopmentTeamQuery = " ( r.typeid <> 1 and p.devcenter = 2 ) "
	elseif request("chkCommercialSWDevelopers") =1 then
		strSelected = strSelected & ", Commercial SW Development Teams"  
		strSWDevelopmentTeamQuery = " ( r.typeid <> 1 and p.devcenter <> 2 ) "
	end if

	strHWDevelopmentTeamQuery = ""
	if request("chkConsumerHWDevelopers") = 1 and request("chkCommercialHWDevelopers") = 1 then
		strSelected = strSelected & ", All HW Development Teams"  
		strHWDevelopmentTeamQuery = " ( r.typeid = 1 )"
	elseif request("chkConsumerHWDevelopers") =1  then
		strSelected = strSelected & ", Consumer HW Development Teams"  
		strHWDevelopmentTeamQuery = " ( r.typeid = 1 and p.devcenter = 2 ) "
	elseif request("chkCommercialHWDevelopers") =1 then
		strSelected = strSelected & ", Commercial HW Development Teams"  
		strHWDevelopmentTeamQuery = " ( r.typeid = 1 and p.devcenter <> 2 ) "
	end if
	
	
	dim strUsers 
	dim strPartnerIDs
	dim strDomains
	strUsers = ""
	strPartnerIDs = ""
	strDomains = ""
	dim blnODMUsers 
	blnODMUsers = false
	
	if request("chkUsersAmericas") = 1 then
		strUsers  = strUsers & ", Americas"  
		strDomains =strDomains & ",'Americas'"
	end if
	if request("chkUsersAsiapacific") = 1 then
		strUsers  = strUsers & ", Asiapacific"  
		strDomains =strDomains & ",'Asiapacific'"
	end if
	if request("chkUsersEMEA") = 1  then
		strUsers  = strUsers & ", EMEA"  
		strDomains =strDomains & ",'emea'"
	end if
	if request("chkUsersInventec") = 1  then
		strUsers  = strUsers & ", Inventec"  
		strPartnerIDs = strPartnerIDs & ",2" 
		blnODMUsers = true
	end if
	if request("chkUsersCompal") = 1  then
		strUsers  = strUsers & ", Compal"  
		strPartnerIDs = strPartnerIDs & ",3" 
		blnODMUsers = true
	end if
	if request("chkUsersFlextronics") = 1  then
		strUsers  = strUsers & ", Flextronics"  
		strPartnerIDs = strPartnerIDs & ",16" 
		blnODMUsers = true
	end if
	if request("chkUsersQuanta") = 1  then
		strUsers  = strUsers & ", Quanta"  
		strPartnerIDs = strPartnerIDs & ",4" 
		blnODMUsers = true
	end if
	if request("chkUsersWistron") = 1  then
		strUsers  = strUsers & ", Wistron"  
		strPartnerIDs = strPartnerIDs & ",7" 
		blnODMUsers = true
	end if
	if request("chkUsersFoxconn") = 1  then
		strUsers  = strUsers & ", Foxconn"  
		strPartnerIDs = strPartnerIDs & ",10" 
		blnODMUsers = true
	end if
	if request("chkUsersModus") = 1  then
		strUsers  = strUsers & ", Modus"  
		strPartnerIDs = strPartnerIDs & ",9" 
		blnODMUsers = true
	end if
	
	if strUsers <> "" then
		strSelected = strSelected  & ", User Groups: " & mid(strUsers,2)
	end if
	
	if strSelected <> "" then
		strSelected = mid(strSelected,3)
	end if
	
	if blnODMUsers then
		Response.Write "<font color=red><b>WARNING:  This list includes people outside HP.</b></font><BR><BR>"
	end if
	
	Response.Write "<b>Groups Selected:</b> " & strSelected
	
	if strPartnerIDs <> "" then
		strSQl = "Select distinct email " & _
				 "from employee with (NOLOCK) " & _
				 "where Partnerid in (" & mid(strPartnerIDs,2) & ") " 
	end if
	if strDomains <> "" then
		if strSQL <> "" then
			strSQL = strSQl & " UNION " 
		end if
		strSQL = strSQl & " Select distinct email " & _
				 "from employee " & _
				 "where Domain in (" & mid(strDomains,2) & ") " & _
				 "and active=1 "& _
				 "and partnerid in (0,1) "
	end if
	
	if strSWDevelopmentTeamQuery <> "" or strHWDevelopmentTeamQuery <> "" then
		if strSQL <> "" then
			strSQL = strSQl & " UNION " 
		end if
		strSQL = strSQl & " Select distinct email " & _ 
						  " from productversion p, product_deliverable pd, deliverableversion v,deliverableroot r, employee e " & _
						  " where p.id = pd.productversionid " & _
						  " and pd.deliverableversionid = v.id " & _
						  " and r.id = v.deliverablerootid " & _
						  " and ( e.id = v.developerid or e.id = r.devmanagerid ) " & _
						  " and p.productstatusid < 4 " & _
						  " and p.typeid = 1 " & _
						  " and r.active =1 " & _
						  " and v.active =1 " & _
						  " and e.active =1 " & _
						  " and e.partnerid in (0,1) " & _
						  " and ( " & strSWDevelopmentTeamQuery & " "
		if strSWDevelopmentTeamQuery <> "" and strHWDevelopmentTeamQuery <> "" then
			strSQl = strSQL & " or "
		end if
		strSQl = strSQL & strHWDevelopmentTeamQuery & " ) "						  
						  
	end if
	
	if strSystemTeamQuery <> "" or strSuperuserQuery <> "" or strSEPMQuery <> "" or strExtendedQuery <> ""  then
		if strSQL <> "" then
			strSQL = strSQl & " UNION " 
		end if
		strSQL = strSQl & " Select distinct email " & _
				 "from productversion p, employee e " & _
				 "where p.productstatusid < 4 and p.typeid = 1 and e.active=1 and e.partnerid in (0,1) and (" 
		strProductQuery = ""
		if strSystemTeamQuery <> "" then
			if strProductQuery = "" then
				strProductQuery = strProductQuery & strSystemTeamQuery 
			else
				strProductQuery = strProductQuery & " or " & strSystemTeamQuery 
			end if
		end if

		if strSuperuserQuery <> "" then
			if strProductQuery = "" then
				strProductQuery = strProductQuery & strSuperuserQuery
			else
				strProductQuery = strProductQuery & " or " & strSuperuserQuery
			end if
		end if

		if strSEPMQuery <> "" then
			if strProductQuery = "" then
				strProductQuery = strProductQuery & strSEPMQuery
			else
				strProductQuery = strProductQuery & " or " & strSEPMQuery
			end if
		end if

		if strExtendedQuery <> "" then
			if strProductQuery = "" then
				strProductQuery = strProductQuery & strExtendedQuery
			else
				strProductQuery = strProductQuery & " or " & strExtendedQuery
			end if
		end if
		
		strSQL = strSQL & strProductQuery & ") "
	end if
	
	

%>
	<%
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	rs.Open strSQL,cn,adOpenStatic
	i = 0
	strEmailList = ""
	do while not rs.EOF
		if i= 0 then
			strEmailList = rs("email") & ""
		else
			strEmailList = strEmailList & "; " & rs("email")
		end if
		i=i+1
		rs.MoveNext
	loop
	rs.Close
	
	set rs = nothing
	cn.Close
	set cn = nothing	
	%>
	<TEXTAREA style="Width: 100%" rows=12 id=txtEmailList name=txtEmailList><%=strEmailList%></TEXTAREA>
	
	<table width=100%><TR><TD nowrap>Email&nbsp;Addresses&nbsp;Found:<%=i%></TD><TD align=right width=100%><INPUT type="button" value="Copy" id=cmdCopy name=cmdCopy LANGUAGE=javascript onclick="return cmdCopy_onclick()"><INPUT type="button" value="Start Over" id=cmdStartOver name=cmdStartOver LANGUAGE=javascript onclick="return cmdStartOver_onclick()"></td></tr></table>
<%end if%>
</form>

<!--<%
else
	Response.write "You are not authorized to view this page."
end if%>-->
</BODY>
</HTML>

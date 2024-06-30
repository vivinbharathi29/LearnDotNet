<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<TITLE>Deliverable Comparison</TITLE>
</HEAD>
<STYLE>
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}
</STYLE>
<BODY>
<FONT face=verdana size=4><b>Compare Versions</b></FONT><BR>
<%
	Dim cn
	dim cm
	dim p
	Dim rs
	dim rs2
	dim strDeliverables
	dim rowcount
	dim ReportName
	dim lastassembly
	dim ID
	dim strChanges
	dim strPart
	dim strOTS
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn  
	set rs2 = server.CreateObject("ADODB.recordset")
	rs2.ActiveConnection = cn  

	ID = request("ID")
	if not isnumeric(trim(ID)) then
		ID = 0
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetDeliverableRootName"
		

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = ID
	cm.Parameters.Append p
	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	if rs.eof and rs.bof then
		Response.Write "Deliverable Not Found"
	else
		reportName = rs("Name")
		rs.close
%>
	

<BR><b><FONT face=Verdana size=3><%=ReportName%></FONT></b><FONT size=1>
<HR color=#006697>

<P><FONT face=Verdana>
<!--<TABLE borderColor=tan cellSpacing=1 cellPadding=1 width="100%"  bgcolor=Ivory border=1>
  
 <TR bgcolor=cornsilk>
    <TD><FONT size=1><b>Version</b></FONT></TD>
    <TD><FONT size=1><b>Developer</b></FONT></TD>
    <TD><FONT size=1><b>Languages</b></FONT></TD>
    <TD><FONT size=1><b>Problem Reports</b></FONT></TD>
    <TD><FONT size=1><b>Changes<b></FONT></TD></TR>
-->
    
    <%
		
	dim strTray
	dim strLanguages
	dim strPanel
	dim strDesktop
    dim strTile
	dim strTaskbarIcon
	dim strCertification
	dim strPackage
	dim strTemp
	dim strNew
	dim strTargeted
	dim strOther
	dim blnLangDiff
	dim strLastLang
	dim blnDTIconDiff
	dim strLastDTIcon
	dim blnTrayIconDiff
	dim strLastTrayIcon
	dim blnCPIconDiff
	dim strLastCPIcon
	dim blnPFWDiff
	dim strLastPFW
	dim blnCertDiff
	dim strLastCert
	dim intCount
    dim blnTileIconDiff
	dim strLastTileIcon
	dim blnTaskbarIconDiff
	dim strLastTaskbarIcon
	dim ReleaseID

    ReleaseID = 0        
    if Request.QueryString("ReleaseID") <> "" then
        ReleaseID = clng(Request.QueryString("ReleaseID"))
    end if

	strTemp = ""
	strOther = ""
	strTargeted = ""
	strNew = ""
	
    'for setting the bg color to mark difference
	blnLangDiff = false
	blnDTIconDiff = false
	blnCPIconDiff = false
	blnPFWDiff = false
	blnCertDiff = false
    blnTileIconDiff = false
	blnTaskbarIconDiff = false

	strLastLang = ""
	strLastDTIcon = ""
	strLastCPIcon = ""
	strLastPFW = ""
	strLastCert = ""
    strLastTileIcon = ""
	strLastTaskbarIcon = ""
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetVersionHistory4Product"
		
	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	p.Value = ID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProdID", 3, &H0001)
	p.Value = request("ProdID")
	cm.Parameters.Append p

    if ReleaseID > 0 then
        Set p = cm.CreateParameter("@ReleaseID",adInteger, &H0001)
	    p.Value = ReleaseID
	    cm.Parameters.Append p
    end if

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	intCount = 0
	do while not rs.EOF

		if isnull (rs("Packageforweb")) then
			strPackage = "No"
		elseif rs("Packageforweb") = 0 then
			strPackage = "No"
		else
			strPackage = "Yes"
		end if
		
		if rs("CertRequired") & "" <> "1" then
			strCertification = "Not Required"
		elseif rs("CertificationStatus") & "" = "0" or rs("CertificationStatus") & "" = "" then
			strCertification = "Required"
		elseif rs("CertificationStatus") = 1 then
			strCertification = "Submitted"
		elseif rs("CertificationStatus") = 2 then
			strCertification = "Approved"
		elseif rs("CertificationStatus") = 3 then
			strCertification = "Failed"
		elseif rs("CertificationStatus") = 4 then
			strCertification = "Waiver"
		else
			strCertification = "&nbsp;"
		end if
		
		if intCount <> 0 and strLastLang <> rs("Languages") and (rs("PMAlert") or rs("Targeted")) then
			blnLangDiff = true
		end if 

		if intCount <> 0 and strLastDTIcon <> rs("IconDesktop") and (rs("PMAlert") or rs("Targeted")) then
			blnDTIconDiff = true
		end if 

		if intCount <> 0 and strLastTrayIcon <> rs("IconTray") and (rs("PMAlert") or rs("Targeted")) then
			blnTrayIconDiff = true
		end if 
		
		if intCount <> 0 and strLastCPIcon <> rs("IconPanel") and (rs("PMAlert") or rs("Targeted")) then
			blnCPIconDiff = true
		end if 

		if intCount <> 0 and strLastPFW <> rs("PackageForWeb") and (rs("PMAlert") or rs("Targeted")) then
			blnPFWDiff = true
		end if 

		if intCount <> 0 and strLastCert <> strCertification and (rs("PMAlert") or rs("Targeted")) then
			blnCertDiff = true
		end if 
		
        if intCount <> 0 and strLastTileIcon <> rs("IconTile") and (rs("PMAlert") or rs("Targeted")) then
			blnTileIconDiff = true
		end if 
		if intCount <> 0 and strLastTaskbarIcon <> rs("IconTaskbarIcon") and (rs("PMAlert") or rs("Targeted")) then
			blnTaskbarIconDiff = true
		end if 

		if rs("PMAlert") or rs("Targeted") then
			strLastDTIcon = rs("IconDesktop")
			strLastCPIcon = rs("Iconpanel")
			strLastTrayIcon = rs("IconTray")
			strLastPFW = rs("PackageForWeb")
			strLastCert = strCertification
			strLastLang = rs("Languages")
            strLastTileIcon = rs("IconTile")
			strLastTaskbarIcon = rs("IconTaskbarIcon")

			intCount = intCount + 1
		end if 
					
		if isnull (rs("Icontray")) then
			strTray = "No"
		elseif rs("IconTray") = 0 then
			strTray = "No"
		else
			strTray = "Yes"
		end if

		if isnull (rs("Iconpanel")) then
			strPanel = "No"
		elseif rs("IconPanel") = 0 then
			strPanel = "No"
		else
			strPanel = "Yes"
		end if

		if isnull (rs("IconDesktop")) then
			strDesktop = "No"
		elseif rs("IconDesktop") = 0 then
			strDesktop = "No"
		else
			strDesktop = "Yes"
		end if
			
        strTile = "Yes"
		if rs("IconTile") = 0 then
			strTile = "No"
		end if
		strTaskbarIcon = "Yes"
		if rs("IconTaskbarIcon") = 0 then
			strTaskbarIcon = "No"
		end if
	
		strChanges = rs("Changes") & ""
		if trim(strchanges) = "" then
			strchanges = "&nbsp;"
		end if
		strpart = rs("partNumber") & ""
		if trim(strpart) = "" then
			strpart = "&nbsp;" 
		else
			strpart = " (" & strpart & ")"
		end if

		strots = ""
		on error resume next
        rs2.Open "spGetOTSByDelVersion "  & rs("ID"), cn, adOpenForwardOnly
		on error goto 0

		if cn.Errors.count > 0 then
			strOTS= "<BR><font size=2 face=verdana color=red><b>OTS is unavailable.</b></font>"
			cn.Errors.clear
		else
			do while not rs2.EOF
				strots = strots & rs2("OTSNumber") & " - " & rs2("shortdescription") & " (Priority: " & rs2("Priority") & ")" &  "<br>"
				rs2.MoveNext		
			loop
			if trim(strots)  ="" then
				strots = "&nbsp;"
			end if
    		rs2.Close
		end if

		strVersion = rs("Version") & ""
		if rs("Revision") & "" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass") & "" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if
		
		strTemp = ""
		strTemp = strTemp & "<TABLE borderColor=tan cellSpacing=1 cellPadding=1 width=""100%""  bgcolor=Ivory border=1>"
		strTemp = strTemp & "<TR><TD><TABLE border=0><TR><TD><b>ID:</b></TD><TD>" & rs("ID") & "</TD></TR><TR><TD><b>Version:</b></TD><TD>" & strVersion & "</TD></TR><TR><TD><b>Vendor Version:</b></TD><td>" & rs("VendorVersion" ) & "</td></tr></Table></td>"
		strTemp = strTemp & "<TD><TABLE border=0><TR><TD><b>Developer:</b></TD><TD>" & rs("Developer") & "</TD></TR><TR><TD><b>Certification:</b></TD><TD bgcolor=""--XXxxXCertXxxXX--"">" & strCertification & "</TD></TR><TR><TD><b>Package For Web:</b></TD><td  bgcolor=""--XXxxXPFWXxxXX--"">" & strpackage & "</td></tr></Table></td>"
        strTemp = strTemp & "<TD><TABLE border=0><TR><TD><b>Icon - Desktop:</b></TD><TD  bgcolor=""--XXxxXDTIconXxxXX--"">" & strDesktop & "</TD></TR><TR><TD><b>Icon - System Tray:</b></TD><TD bgcolor=""--XXxxXTrayXxxXX--"">" & strTray & "</TD></TR><TR><TD><b>Icon - Control Panel:</b></TD><td bgcolor=""--XXxxXCPIconXxxXX--"">" & strPanel & "</td></tr></Table></td><TD><TABLE border=0><TR><TD><b>Icon - Start Menu Tile:</b></TD><TD bgcolor=""--XXxxXDTIconXxxXX--"">" & strTile & "</TD></TR><TR><TD><b>Icon - TaskbarIcon:</b></TD><TD bgcolor=""--XXxxXDTIconXxxXX--"">" & strTaskbarIcon & "</TD></TR></TABLE></TD></tr>"
		strTemp = strTemp & "<TR><TD colspan=5 bgcolor=""--XXxxXLanguagesXxxXX--""><TABLE border=0><TR><TD valign=top width=90><b>Languages:</b></td><td>" & rs("Languages") & "</td></tr></table></td></tr>"
		strTemp = strTemp & "<TR><TD colspan=5><TABLE border=0><TR><TD valign=top width=90><b>Changes:</b></td><td>" & replace(rs("Changes") & "",vbcrlf,"<BR>") & "</td></tr></table></td></tr>"
		strTemp = strTemp & "<TR><TD colspan=5><TABLE border=0><TR><TD valign=top width=90><b>OTS Fixed:</b></td><td>" & replace(strOTS,vbcrlf,"<BR>") & "</td></tr></table></td></tr>"
		strTemp = strTemp & "<TR><TD colspan=5><TABLE border=0><TR><TD valign=top width=90><b>Comments:</b></td><td>" & replace(rs("Comments") & "",vbcrlf,"<BR>") & "</td></tr></table></td></tr>"		
		strTemp = strTemp & "</table><BR>"

	
		if rs("PMAlert") then
			strNew = strTemp & strNew  
		elseif rs("Targeted") then
			strTargeted = strTemp & strTargeted 
		else
			strOther = strTemp & strOther  
		end if


		rs.movenext
	loop
	
    rs.close
    cn.Close
    
    
    if strNew <> "" then
		if blnLangDiff then
			strNew = replace(strNew,"--XXxxXLanguagesXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXLanguagesXxxXX--","ivory")
		end if

		if blnDTIconDiff then
			strNew = replace(strNew,"--XXxxXDTIconXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXDTIconXxxXX--","ivory")
		end if

		if blnCPIconDiff then
			strNew = replace(strNew,"--XXxxXCPIconXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXCPIconXxxXX--","ivory")
		end if

		if blnTrayIconDiff then
			strNew = replace(strNew,"--XXxxXTrayXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXTrayXxxXX--","ivory")
		end if

		if blnPFWDiff then
			strNew = replace(strNew,"--XXxxXPFWXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXPFWXxxXX--","ivory")
		end if

		if blnCertDiff then
			strNew = replace(strNew,"--XXxxXCertXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXCertXxxXX--","ivory")
		end if

        if blnTileIconDiff then
			strNew = replace(strNew,"--XXxxXCertXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXCertXxxXX--","ivory")
		end if

		if blnTaskbarIconDiff then
			strNew = replace(strNew,"--XXxxXCertXxxXX--","mistyrose")
		else
			strNew = replace(strNew,"--XXxxXCertXxxXX--","ivory")
		end if
		
		Response.Write "<font size=2 face=verdana><b>New Releases:</b></font><BR>" & strNew

    end if
    if strtargeted <> "" then
		if blnLangDiff then
			strtargeted = replace(strtargeted,"--XXxxXLanguagesXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXLanguagesXxxXX--","ivory")
		end if

		if blnDTIconDiff then
			strtargeted = replace(strtargeted,"--XXxxXDTIconXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXDTIconXxxXX--","ivory")
		end if

		if blnCPIconDiff then
			strtargeted = replace(strtargeted,"--XXxxXCPIconXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXCPIconXxxXX--","ivory")
		end if

		if blnTrayIconDiff then
			strtargeted = replace(strtargeted,"--XXxxXTrayXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXTrayXxxXX--","ivory")
		end if
    
		if blnPFWDiff then
			strtargeted = replace(strtargeted,"--XXxxXPFWXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXPFWXxxXX--","ivory")
		end if

		if blnCertDiff then
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","ivory")
		end if

        if blnTileIconDiff then
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","ivory")
		end if

		if blnTaskbarIconDiff then
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","mistyrose")
		else
			strtargeted = replace(strtargeted,"--XXxxXCertXxxXX--","ivory")
		end if		
    
		Response.Write "<font size=2 face=verdana><b><BR>Targeted Versions:</b></font><BR>" & strtargeted

    end if
    if StrOther <> "" then
		Response.Write "<font size=2 face=verdana><b><BR>Other Versions:<BR></b></font>" & replace(replace(replace(replace(replace(replace(strOther,"--XXxxXLanguagesXxxXX--","ivory"),"--XXxxXDTIconXxxXX--","ivory"),"--XXxxXTrayXxxXX--","ivory"),"--XXxxXCPIconXxxXX--","ivory"),"--XXxxXPFWXxxXX--","ivory"),"--XXxxXCertXxxXX--","ivory")
    end if
    
    
    %>

    <!--</TABLE>-->
</FONT></P>
<%
	end if
%>
</BODY>
</HTML>
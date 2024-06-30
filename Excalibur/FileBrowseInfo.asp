<%@ Language=VBScript %>

<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="style/general.css">
<SCRIPT  id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

    function window_onload() {
        if (typeof (lblBrowse) == "undefined") {
            return;
        }
        else {
            lblBrowse.style.display = "none";
        }
    }

    function ShowPath2(strPath2Location) {
        window.open("file://" + strPath2Location);
    }

//-->
</SCRIPT>
</HEAD>
<BODY onload="return window_onload()">

<FONT SIZE="2">

<%
dim ProgramOfficeServer
dim PrinstallServer
dim SwtechServer

ProgramOfficeServer = InStr(1, lcase(request("DeliverablePath")), "ccmptple01", 1)
PreinstallServer = InStr(1, lcase(request("DeliverablePath")), "houbnbpindev01", 1)
SwtechServer = InStr(1, lcase(request("DeliverablePath")), "mobileswtech", 1)

'Get User
dim CurrentDomain
dim Currentuser
CurrentUser = lcase(Session("LoggedInUser"))

if instr(currentuser,"\") > 0 then
	CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
	Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
end if

'if the current user is not TDC or ODM in Asia, the default display screen shows the Houston path
if CurrentDomain <> "asiapacific" and CurrentDomain <> "excaliburweb" then
'Response.Write "Current user Domain: " & currentDomain
if request("Path2Location") = "" then
	if request("TDCImagePath") = "" then
		if request("DisplayError") <> "" then
			Response.Write request("DisplayError")
		elseif request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files for... </div><b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b> "&request("DeliverablePath")&"<BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files for... </div>" & request("DeliverablePath") & "<BR>"
		else
			Response.Write "Please sepcify a ID or path to download... <BR>"
		end if
	elseif request("TDCImagePath") <> "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"
		elseif request("DeliverablePath") = "" then 'display TDC path if "deliverable path is not set
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"

			Response.Write "<b>File Location being displayed is TDC path: </b>" & request("TDCImagePath") & "<BR>"	
		end if
	end if
  
	if request("Instr1") <> "" then
		Response.write request("Instr1")&"<BR>"
	end if
elseif request("Path2Location") <> "" then
  if request("Path3Location") = "" then 'no path3location
	if request("TDCImagePath") = "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		end if
	else
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"			
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"			
		elseif request("DeliverablePath") = "" then 'display TDC path if "deliverable path is not set
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed is TDC path: </b>" & request("TDCImagePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		end if
	end if 'TDCImagepath
  elseif request("Path3Location") <> "" then 'there is a path3location 
    if request("TDCImagePath") = "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		end if
	else
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"		
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
			Response.write "<b>TDC Path: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"		
		elseif request("DeliverablePath") = "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>File Location being displayed is TDC path: </b>" & request("TDCImagePath") & "<BR>"
			Response.write "<b>Also Available: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		end if
	end if 'TDCImagepath
  end if 'path3location
end if 'path2location

else 'if current user is asiapacific, default screen display the TDC path first
if request("Path2Location") = "" then
	if request("TDCImagePath") = "" then
		if request("DisplayError") <> "" then
			Response.Write request("DisplayError")
		elseif request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files for... </div><b>Deliverable: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b> "&request("DeliverablePath")&"<BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files for... </div>" & request("DeliverablePath") & "<BR>"
		else
			Response.Write "Please sepcify a ID or path to download... <BR>"
		end if
	elseif request("TDCImagePath") <> "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""file://" & request("DeliverablePath") & """>" & request("DeliverablePath") & "</a><BR>"
		elseif request("DeliverablePath") = "" then 'display TDC path if "deliverable path is not set
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"

			Response.Write "<b>TDC path: </b>" & request("TDCImagePath") & "<BR>"	
		end if
	end if
  
	if request("Instr1") <> "" then
		Response.write request("Instr1")&"<BR>"
	end if
elseif request("Path2Location") <> "" then
  if request("Path3Location") = "" then 'no path3location
	if request("TDCImagePath") = "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Houston File Location being displayed: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>Houston File Location being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		end if
	else 'TDCImagePath is available
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"			
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"			
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""file://" & request("DeliverablePath") & """>" & request("DeliverablePath") & "</a><BR>"
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		elseif request("DeliverablePath") = "" then 'display TDC path if "deliverable path is not set
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>TDC path being displayed: </b>" & request("TDCImagePath") & "<BR>"
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a><BR>"
		end if
	end if 'TDCImagepath
  elseif request("Path3Location") <> "" then 'there is a path3location 
    if request("TDCImagePath") = "" then
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>Houston loction file being displayed: </b>" & request("DeliverablePath") & "<BR>"
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		end if
	else 'TDCImagePath exist
		if request("DeliverableID") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
	
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"		
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""wizardframes.asp?ID=" & request("DeliverableID") & "&type=1"">" & request("DeliverableName") & "</a>" & "<BR><b>Path:</b>" & request("DeliverablePath") & "<BR>"	
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		elseif request("DeliverablePath") <> "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.write "<b>TDC Path being displayed: </b><a target=""_blank"" href=""file://" & request("TDCImagePath") & """>" & request("TDCImagePath") & "</a><BR>"		
			Response.Write "<b>Houston Path: </b><a target=""_blank"" href=""file://" & request("DeliverablePath") & """>" & request("DeliverablePath") & "</a><BR>"
			Response.write "<b>Also Available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		elseif request("DeliverablePath") = "" then
			Response.write "<div id=lblBrowse>Browsing Files, Please wait ... </div>"
		
			Response.Write "<b>TDC path being displayed: </b>" & request("TDCImagePath") & "<BR>"
			Response.write "<b>Deliverable available in Houston: </b><a target=""_blank"" href=""file://" & request("Path2Location") & """>" & request("Path2Description") & "</a>&nbsp;&nbsp"
			Response.write "<b>And  </b><a target=""_blank"" href=""file://" & request("Path3Location") & """>" & request("Path3Description") & "</a><BR>"
		end if
	end if 'TDCImagepath
  end if 'path3location
end if 'path2location
end if 'currentuser


if request("DisplayError") = "" then
	Response.Write "If you have any problems or questions, please contact "	 
	if request("Developer") <> "" then
		Response.Write " the developer <A href=""mailto:" & request("DeveloperEmail") & """>" & request("Developer") & "</A>"
		'Response.Write " or the <A HREF=""mailto:releaseteam@hp.com;prtreleaselab@hp.com;consnbreleaselab@hp.com"">RELEASEGROUP</A>."
	    if  ProgramOfficeServer > 0 then
			Response.Write " or <A HREF=""mailto:Schelli.Bettega@hp.com"">Schelli Bettega</A>"
		elseif PrinstallServer > 0 then
			Response.write " or <A HREF=""mailto:sal.vasi@hp.com"">Sal Vasi</A>"
		elseif SwtechServer > 0 then
			Response.write " or <A HREF=""mailto:danny.weaver@hp.com"">Danny Weaver</A>"    
		else
			if CurrentDomain = "asiapacific" or CurrentDomain = "excaliburweb" then
				Response.Write " or the <A HREF=""mailto:Marco.burgos@hp.com"">TDC Release Team</A>."
			else
				Response.Write " or the <A HREF=""mailto:psgsoftpaqsupport@hp.com;twn.pdc.nb-releaselab@hp.com"">Release Team</A>."
			end if
	    end if
	else
	    if  ProgramOfficeServer > 0 then
			Response.Write " <A HREF=""mailto:Schelli.Bettega@hp.com"">Schelli Bettega</A>"
		elseif PrinstallServer > 0 then
			Response.write " <A HREF=""mailto:sal.vasi@hp.com"">Sal Vasi</A>"
		elseif SwtechServer > 0 then
			Response.write " <A HREF=""mailto:danny.weaver@hp.com"">Danny Weaver</A>"    
		else
			if CurrentDomain = "asiapacific" or CurrentDomain = "excaliburweb" then
				Response.Write " the <A HREF=""mailto:Marco.burgos@hp.com"">TDC Release Team</A>."
			else
				Response.Write " the <A HREF=""mailto:releaseteam@hp.com;prtreleaselab@hp.com;twn.pdc.nb-releaselab@hp.com"">Release Team</A>."
			end if
		end if
	end if
end if
	
'Response.Flush
%>
	
</FONT>
</BODY>
</HTML>

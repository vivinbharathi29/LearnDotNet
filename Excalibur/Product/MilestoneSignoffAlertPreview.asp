<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
<TITLE>RTM - Alert Preview</TITLE>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function window_onload(){

}


//-->
</SCRIPT>
</HEAD>
<STYLE>
a:link, A:visited
{
    COLOR: blue
}
a:hover 
{
    COLOR: red
}

.EmbeddedTable TBODY TD{
	FONT-FAMILY: Verdana;
}
.EmbeddedTable TBODY TD{
	Font-Size: xx-small;
}

.AlertTable TD
{
    BORDER-COLOR: gray;
    BACKGROUND-COLOR: white;
	Font-Size: xx-small;
	FONT-FAMILY: Verdana;
}
.AlertHeader TD
{
    BACKGROUND-COLOR: gainsboro;
}
.AlertNone TD
{
    BACKGROUND-COLOR: #ebf5db;
}

input
{
    FONT-SIZE: 10pt;	
    FONT-FAMILY: Verdana;	
}
textarea
{
    FONT-SIZE: 10pt;	
    FONT-FAMILY: Verdana;	
}
.ImageTable TBODY TD{
	BORDER-TOP: gray thin solid;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}
.ImageTable TH{
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}

.imagerows TBODY TD{
	BORDER-TOP: none;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}

.imagerows THEAD TD{
	BORDER-TOP: none;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}
</STYLE>
<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<%
    dim cn
    dim rs
    
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    rs.open "spGetRTMAlert " & clng(request("ID")),cn
    if rs.eof and rs.bof then
        response.Write "Unable to find the requested information."
    else
        select case rs("ReportSectionID")
        case "1"
            response.write "<font size=2 face=verdana><b>Build Level Alerts"
        case "2"
            response.write "<font size=2 face=verdana><b>Distribution Alerts"
        case "3"
            response.write "<font size=2 face=verdana><b>Certification Alerts"
        case "4"
            response.write "<font size=2 face=verdana><b>Workflow Alerts"
        case "5"
            response.write "<font size=2 face=verdana><b>Availabiity Alerts"
        case "6"
            response.write "<font size=2 face=verdana><b>Developer Alerts"
        case "7"
            response.write "<font size=2 face=verdana><b>Root Deliverable Alerts"
        case "8"
            response.write "<font size=2 face=verdana><b>Primary OTS Alerts"
        end select
        response.Write " for " & rs("Product") 
        if trim(rs("Title") & "") <> "" then
            response.Write " - " & rs("Title")
        end if
        response.Write "</b><BR><BR></font>"
        
        response.write "<font size=1 face=verdana>These alerts were reviewed by " & longname(rs("UserName") & "") & " on " & formatdatetime(rs("LastUpdated"),vbshortdate  ) & "<BR><BR>"
        if trim(rs("Comments") & "") <> "" then
            response.write "<table cellpadding=3 bgcolor=#ebf5db border=1 width=""100%""><tr><td><font size=1 face=verdana>" & rs("Comments") & "</font></td></tr></table><BR>"
        end if

        response.Write  replace(replace(rs("AlertHTML") & "","search/ots/Report.asp","../search/ots/Report.asp"),"otsDetails.asp?txtFunction=1&txtNumbers=","../search/ots/Report.asp?txtReportSections=1&txtObservationID=")
    end if
    rs.close
    set rs = nothing
    cn.close
    set cn = nothing




	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
		
	end function	

%>

</BODY>
</HTML>



<%@ Language=VBScript %>
	
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value!="0")
		{
		window.returnValue = txtSuccess.value;
		window.close();
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
	dim strSuccess	

	strSuccess = ""

	if request("ProductID") = "" or request("ProductID") = "0" then
		Response.Write "Not enough information supplied to complete this action."
	else
		dim strSQL
		dim cn
	
		strSQL = "spAddOTSComponents " & clng(request("ProductID")) & "," & request("PDMID") & "," & request("SEPMID") & ",0," & request("PINPMID") & "," & request("PEID") & ",0," & request("SC")

		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.Recordset")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		cn.Execute strSQL
		if cn.Errors.count > 0 then
			strSuccess = "0"
		else
			strSuccess = "1"
		end if
		
		strOutput = ""
		if true then 'strSuccess = "1" then
			'CreateOutput
			
			dim blnShowCMTBoxes
			blnShowCMTBoxes = false

			
			
			strOutput = ""
			rs.Open "spListProductOTSComponents " & clng(request("ProductID")) & "",cn,adOpenStatic ' & clng(request("ProductID"))
			if not (rs.EOF and rs.BOF) then
					strOutput = "<Table style=""width:100%"" cellpadding=2 cellspacing=0>" 
					if blnShowCMTBoxes then
						strOutput = strOutput & "<TR><td nowrap><font face=verdana size=1><INPUT style=""WIDTH:16;Height:16"" type=""checkbox"" id=chkCMTAll name=chkCMTAll LANGUAGE=javascript onclick=""return chkCMTAll_onclick()""> <b>Source&nbsp;&nbsp;</b></font></td>"
					else
						strOutput = strOutput & "<TR><td><font face=verdana size=1><b>Source&nbsp;&nbsp;</b></font></td>"
					end if
					strOutput = strOutput & "<td><font face=verdana size=1><b>Err&nbsp;Type</b></font></td>"
					strOutput = strOutput & "<td><font face=verdana size=1><b>Category</b></font></td>"
					strOutput = strOutput & "<td><font face=verdana size=1><b>Component</b></font></td>"
					strOutput = strOutput & "<td><font face=verdana size=1><b>PM</b></font></td>"
					strOut = strOutput & "<td><font face=verdana size=1><b>Developer</b></font></td></TR>"

					do while not rs.EOF
						if rs("ID") = 0 then
							strOutput = strOutput & "<TR bgcolor=Lavender><td class=OTSComponentCell><INPUT style=""display:none;WIDTH:16;Height:16"" value=""" & rs("Partnumber") & """ type=""checkbox"" id=chkCMT name=chkCMT>CMT</td>"
						else
							strOutput = strOutput & "<TR bgcolor=white><td class=OTSComponentCell>Excalibur</td>"
						end if
						strOutput = strOutput & "<td class=OTSComponentCell>" & rs("ErrorType") & "</td>"
						strOutput = strOutput & "<td class=OTSComponentCell>" & rs("category") & "</td>"
						strOutput = strOutput & "<td class=OTSComponentCell>" & rs("Component") & "</td>"
						if rs("ID") = 0 then
							strOutput = strOutput & "<td class=OTSComponentCell>" & rs("PM") & "</td>"
							strOutput = strOutput & "<td class=OTSComponentCell>" & rs("Developer")& "</td></TR>"
						else
							strOutput = strOutput & "<td ID=OTSPM" & trim(rs("ID")) & " class=OTSComponentCell><a href=""javascript: EditOTSPM(" & rs("ID")& "," & rs("PMID") & ")"">" & longname(rs("PM")&"") & "</a></td>"
							strOutput = strOutput & "<td ID=OTSDeveloper" & trim(rs("ID")) & " class=OTSComponentCell><a href=""javascript: EditOTSDeveloper(" & rs("ID")& "," & rs("DeveloperID") & ")"">" & longname(rs("Developer")&"")& "</a></td></TR>"
						end if
						rs.MoveNext
					loop
				strOutput = strOutput & "</TABLE>"
			end if
			rs.close	
		end if
		cn.Close
		set rs = nothing
		set cn = nothing
	end if
	
	
	
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function	
	
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
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=server.HTMLEncode(strOutput)%>">
</BODY>
</HTML>

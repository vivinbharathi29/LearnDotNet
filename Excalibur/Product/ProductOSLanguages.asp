<%@ Language=VBScript %>

	<%
    if request("ReportFormat")= "Excel" then
		Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader "content-disposition","attachment; filename=observation.xls"
	elseif request("ReportFormat")= "Word" then
		Response.ContentType = "application/msword"
        Response.AddHeader "content-disposition","attachment; filename=observation.doc"
    else
      Response.Buffer = True
      Response.ExpiresAbsolute = Now() - 1
      Response.Expires = 0
      Response.CacheControl = "no-cache"
	end if		
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</STYLE>

</HEAD>


<BODY>
<%
    dim cn, rs, strSQL
    dim strLastOS
    dim strLanguageString
    dim strLang
    dim LangArray

    if request("ID") = "" then
        response.write "No product specified."
    else
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open
	    set rs = server.CreateObject("ADODB.recordset")

                strSQl = "Select distinct v.id as ProductID, v.DOTSName, case when coalesce(r.otherlanguage,'') = '' then r.OSLanguage else r.OSLanguage + ',' + r.OtherLanguage end as OSLanguages , o.Name as OSName " & _
                     "from productversion v with (NOLOCK), imagedefinitions d with (NOLOCK), images i with (NOLOCK), regions r with (NOLOCK), OSLookup o with (NOLOCK) " & _
                     "where v.ID = d.ProductVersionID " & _
                     "and d.ID = i.ImageDefinitionID " & _
                     "and i.RegionID = r.ID " & _
                     "and o.ID = d.OSID " & _
                     "and d.active=1 " & _
                     "and ProductVersionID in (" & request("ID") & ") " & _
                     " UNION " & _
                     "Select distinct v.id as ProductID, v.DOTSName, case when coalesce(r.otherlanguage,'') = '' then r.OSLanguage else r.OSLanguage + ',' + r.OtherLanguage end as OSLanguages , o.Name as OSName " & _
                     "from productversion v with (NOLOCK), productversion_ProductDrop pvpd with (NOLOCK), imagedefinitions d with (NOLOCK), images i with (NOLOCK), regions r with (NOLOCK), OSLookup o with (NOLOCK) " & _
                     "where v.ID = pvpd.ProductVersionID " & _
                     "and pvpd.ProductDropID=  d.ProductDropID  " & _
                     "and d.ID = i.ImageDefinitionID " & _
                     "and i.RegionID = r.ID  " & _
                     "and o.ID = d.OSID  " & _
                     "and d.active=1  " & _
                     "and v.ID in (" & request("ID") & ") " & _
                     "and v.Fusion = 1"


       ' strSQl = "Select distinct v.DOTSName, case when coalesce(r.otherlanguage,'') = '' then r.OSLanguage else r.OSLanguage + ',' + r.OtherLanguage end as OSLanguages , o.Name as OSName " & _
       '          "from productversion v with (NOLOCK), imagedefinitions d with (NOLOCK), images i with (NOLOCK), regions r with (NOLOCK), OSLookup o with (NOLOCK) " & _
       '          "where v.ID = d.ProductVersionID " & _
       '          "and d.ID = i.ImageDefinitionID " & _
       '          "and i.RegionID = r.ID " & _
       '          "and o.ID = d.OSID " & _
       '          "and d.active=1 " & _
       '          "and ProductVersionID = " & clng(request("ID")) & " " & _
       '          "order by o.Name"

        rs.open strSQL, cn
        if rs.eof and rs.bof then
            response.write "Unable to find the selected product."
        else
            strLastOS=""
            strLanguageString=""
            response.write "<table bgcolor=ivory border=1 cellpadding=2 cellspacing=0>"
            response.write "<h3>" & rs("DotsName") & "</h3>"
            do while not rs.eof
                if trim(lcase(rs("OSName") & "")) <> trim(lcase(strLastOS)) then
                    if strLanguageString <> "" then
                        response.Write "<tr><td>" & mid(strLanguageString,2) & "</td></tr>"
                    end if
                    response.write "<tr  style=""font-weight:bold"" bgcolor=beige><td>" & rs("OSName") & "</td></tr>"
                    strLanguageString = ""
                end if
                strLastOS = rs("OSName") & ""
                LangArray = split(rs("OSLanguages"),",")
                for each strLang in LangArray
                    if trim(strlang) <> "" then
                        if instr(strLanguageString,strLang) = 0 then
                            strLanguageString = strLanguageString & "," & strlang
                        end if
                    end if
                next
                rs.movenext
            loop
            rs.close
            if strLanguageString <> "" then
                response.Write "<tr><td>" & mid(strLanguageString,2) & "</td></tr>"
            end if
            response.write "</table>"
        end if

        set rs = nothing
        cn.Close
        set cn = nothing
    end if
%>
</BODY>
</HTML>





<%@ Language=VBScript %>

	<%
    response.ContentType="text/xml"
	%>
<%
    dim cn, rs, strSQL
    dim strLastOS
    dim strLanguageString
    dim strLang
    dim LangArray
    dim ProductArray
    dim blnProductListOK
    dim strProduct

    if request("ID") = "" then
        response.write "<rows><row><error>No product specified.</error></row></rows>"
    else
        blnProductListOK = true
        ProductArray = split(request("ID"),",")

        for each strProduct in ProductArray
            if not isnumeric(trim(strProduct)) then
                blnProductListOK = false
                exit for
            end if
        next

        if not blnProductListOK then
            response.write "<rows><row><error>Invalid product list specified.  Please supply a comma-separated list of product ID numbers.</error></row></rows>"
        else
	        set cn = server.CreateObject("ADODB.Connection")
	        cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	        cn.Open
	        set rs = server.CreateObject("ADODB.recordset")

            strSQl = "Select distinct v.id as ProductID, v.DOTSName as Product, case when coalesce(r.otherlanguage,'') = '' then r.OSLanguage else r.OSLanguage + ',' + r.OtherLanguage end as OSLanguages , o.Name as OSName " & _
                     "from productversion v with (NOLOCK), imagedefinitions d with (NOLOCK), images i with (NOLOCK), regions r with (NOLOCK), OSLookup o with (NOLOCK) " & _
                     "where v.ID = d.ProductVersionID " & _
                     "and d.ID = i.ImageDefinitionID " & _
                     "and i.RegionID = r.ID " & _
                     "and o.ID = d.OSID " & _
                     "and d.active=1 " & _
                     "and ProductVersionID in (" & request("ID") & ") " & _
                     " UNION " & _
                     "Select distinct v.id as ProductID, v.DOTSName as Product, case when coalesce(r.otherlanguage,'') = '' then r.OSLanguage else r.OSLanguage + ',' + r.OtherLanguage end as OSLanguages , o.Name as OSName " & _
                     "from productversion v with (NOLOCK), productversion_ProductDrop pvpd with (NOLOCK), imagedefinitions d with (NOLOCK), images i with (NOLOCK), regions r with (NOLOCK), OSLookup o with (NOLOCK) " & _
                     "where v.ID = pvpd.ProductVersionID " & _
                     "and pvpd.ProductDropID=  d.ProductDropID  " & _
                     "and d.ID = i.ImageDefinitionID " & _
                     "and i.RegionID = r.ID  " & _
                     "and o.ID = d.OSID  " & _
                     "and d.active=1  " & _
                     "and v.ID in (" & request("ID") & ") " & _
                     "and v.Fusion = 1"


            if request("Status") <> "" and isnumeric(request("Status") ) then
                strSQL = strSQL & " and d.statusid=" & clng(request("Status"))
            end if

            strSQl = strSQl & " order by v.Dotsname, o.Name"

            rs.open strSQL, cn
            if rs.eof and rs.bof then
                response.write "<rows><row><error>No rows match search criteria.</error></row></rows>"
            else
                strLastOS=""
                strLanguageString=""
                response.write "<rows>"
                do while not rs.eof
                    if trim(lcase(rs("OSName") & "")) <> trim(lcase(strLastOS)) then
                        if strLanguageString <> "" then
                            response.Write "<languages>" & mid(strLanguageString,2) & "</languages>"
                            strLanguageString = ""
                            response.write "</row>" 
                        end if
                        response.write "<row>" 
                        response.write "<productid>" & rs("ProductID") & "</productid>"
                        response.write "<product>" & rs("Product") & "</product>"
                        response.write "<os>" & rs("OSName") & "</os>"
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
                if strLanguageString <> "" then
                    response.Write "<languages>" & mid(strLanguageString,2) & "</languages>"
                    response.write "</row>"
                end if
                rs.close
                response.write "</rows>"

            end if

            set rs = nothing
            cn.Close
            set cn = nothing
        end if
    end if
%>





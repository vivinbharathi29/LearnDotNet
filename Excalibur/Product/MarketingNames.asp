<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<html>
<head>
<title>Excalibur - Marketing Names</title>
<meta name="VI60_DefaultClientScript" content="JavaScript" />

<link href="./style/wizard style.css" type="text/css" rel="stylesheet" />
<link href="./style/Excalibur.css" type="text/css" rel="stylesheet" />
<STYLE>
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}
</STYLE>
</head>
<body >

<font size=3 face="verdana"><b>Product Marketing Names</b></font><br /><br />

<%
		dim rs
		dim rs2
		dim cn
		dim cm
		dim p

		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.CommandTimeout =120
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")

        rs.open "Select id, dotsname from productversion where  productstatusid < 5 and typeid in (1,3) order by dotsname;",cn
        if not rs.eof then
            response.write "<TABLE bgcolor=""ivory"" border=1 bordercolor=""gainsboro"" cellspacing=0 cellpadding=2><TR>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Product</TD>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Long&nbsp;Name</TD>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Short&nbsp;Name</TD>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Logo&nbsp;Badge&nbsp;C&nbsp;Cover</TD>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Family&nbsp;Name</TD>"
            response.write "<TD nowrap style=""font-size:xx-small;font-weight:bold"">Brand&nbsp;Name</TD>"
            response.write "</tr>"
        else
            response.write "No products selected."
        end if
		do while not rs.eof
        response.flush
            dim strShortName
    	    dim strLogoBadge
			dim strPHWebFamily
			dim strBrandName
			dim strKMAT
				
			strLogoBadge = ""
			strShortName = ""
			strPHWebFamily = ""
			strBrandName = ""
			strKMAT = ""


                    strSeries = ""
				rs2.Open  "spListbrands4Product " & rs("ID") & ",1",cn,adOpenForwardOnly
				do while not rs2.EOF
					if trim(rs2("SeriesSummary") & "") <> "" then
						SeriesArray = split(rs2("SeriesSummary"),",")
						for i = 0 to ubound(SeriesArray)
							if trim(seriesArray(i)) <> "" then
								if rs2("StreetName") <> "" then
									strSeries =  strSeries & rs2("StreetName") & " " 
									strShortName= strShortName & rs2("StreetName2") & " " 
									strLogoBadge =  strLogobadge & rs2("StreetName3") & " " 
								    strBrandName = strBrandName & rs2("StreetName") & " "
								    strKMAT = strKMAT & rs2("KMAT") & " "
								end if
							
								if rs2("productversion") & "" <> "" then
									if lcase(rs2("productfamily")&"") = "davos"  and right(rs2("productversion") & "",3) = "1.0" then
										strPHWebFamily = strPHWebFamily & left(rs2("product"),len(rs2("product"))-1) & "X - " &  rs2("StreetName") & " " 
									elseif isnumeric(mid(rs2("productversion") & "",len(rs2("productversion") & ""),1)) then
										strPHWebFamily = strPHWebFamily & rs2("productfamily") & " " & rs2("RASSegment") & " " &  left(rs2("productversion"), len(rs2("productversion"))-1) & "X - " &  rs2("StreetName") & " " 
									else
										strPHWebFamily = strPHWebFamily & rs2("productfamily") & " " & rs2("RASSegment") & " " &  left(rs2("productversion"), len(rs2("productversion"))-2) & "X - " &  rs2("StreetName") & " " 
									end if
								end if
								strSeries = strSeries & seriesArray(i) 
                                if rs2("ShowSeriesNumberInShortName") then
								    strShortName= strShortName & seriesArray(i) 
                                else
								    strShortName= strShortName 
                                end if
								if rs2("ShowSeriesNumberInLogoBadge") then
                                    if rs2("SplitSeriesForLogoAndBrand") then
							            strLogoBadge = strLogobadge & val(seriesArray(i))
                                    else
								        strLogoBadge = strLogobadge & seriesArray(i)
                                    end if
								else
								    strLogoBadge = strLogobadge 
								end if
								if rs2("ShowSeriesNumberInBrandname") then
                                    if rs2("SplitSeriesForLogoAndBrand") then
       								    strBrandName = strBrandName & val(seriesArray(i))
                                    else
    								    strBrandName = strBrandName & seriesArray(i)
                                    end if
    							else
								    strBrandName = strBrandName 
								end if
								strPHWebFamily = strPHWebFamily & seriesArray(i)
							
								if trim(rs2("Suffix") & "") <> "" then
									strSeries = strSeries & " " & trim(rs2("Suffix") & "")
								end if
								strSeries = strSeries & "<BR>"
								strShortName= strShortName & "<BR>"
								strLogoBadge =  strLogobadge & "<BR>"
								strBrandName = strBrandName & "<BR>"
								strPHWebFamily = strPHWebFamily & "<BR>"
								strKMAT = strKMAT & "<BR>"
							end if
						next

					end if
					rs2.MoveNext
				loop
				rs2.Close

        				response.write  "<tr><TD nowrap valign=top style=""font-size:xx-small"">" & rs("DOTSName") & "</TD><TD nowrap valign=top style=""font-size:xx-small"">" & strSeries & "</TD><TD nowrap valign=top style=""font-size:xx-small"">" & strShortName & "</TD><TD nowrap valign=top style=""font-size:xx-small"">" & strLogoBadge & "</TD>"
                        response.write "<TD nowrap valign=top style=""font-size:xx-small"">" & strPHWebFamily & "</td>"
                        response.write "<TD nowrap valign=top style=""font-size:xx-small"">" & strBrandName & "</td>"
                        response.write "</TR>"

            rs.movenext
        loop
        if not (rs.bof and rs.eof) then
            response.write "</TABLE>"
        end if

    function Val(strText)
        dim strOutput
        dim i

        strOutput = ""
        for i = 1 to len(trim(strText))
            if isnumeric(mid(strText,i,1)) then
                strOutput = strOutput & mid(trim(strText),i,1)
            else
                exit for
            end if
        next
        Val = strOutput
    end function

	%>
<br />
<br />
<font face=verdana Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
</body>
</html>

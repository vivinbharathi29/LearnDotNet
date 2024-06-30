<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%
	dim cn
	dim rs
	dim strBrands
	dim BrandArray
	dim strBrand

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
		
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	rs.Open "Select ID, Brands from productversion with (NOLOCK)",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(rs("Brands") & "") <> "" then
			strBrands = rs("Brands")
			BrandArray = split(strBrands,",")
			for each strBrand in BrandArray
				if strBrand <> "" then
					set rs2 = server.CreateObject("ADODB.Recordset")
					rs2.Open "Select ID,Name from productlevel3 with (NOLOCK) where Name like '" & trim(strBrand) & "'",cn,adOpenForwardOnly
					if rs2.EOF and rs2.BOF then
						Response.Write "Not Found:" & rs("ID") & ":" & strBrand & "<BR>"
					else
						cn.Execute  "Insert Product_Productlevel3(ProductversionID,ProductLevel3ID) values(" & rs("ID") & "," & rs2("ID") & ")"
					end if
					set rs2=nothing
				end if
			next
		end if
		rs.MoveNext
	loop
	rs.Close	


	set rs = nothing
	cn.Close
	set cn=nothing

%>

</BODY>
</HTML>

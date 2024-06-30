<script runat="server" language="vbscript">

    dim strFavs, FavCount, EmployeeID, returnValue
    strFavs = request.QueryString("Favorites")
    FavCount = request.QueryString("FavCount")
    EmployeeID = request.QueryString("CurrentUserID")
    
    dim cn 
	dim cm 
	dim i

	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spUpdateFavorites"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
	p.value =EmployeeID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FavCount", 2,  &H0001)
	p.value =FavCount
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@Favorites", 201, &H0001, 2147483647)
	p.value = strFavs
	cm.Parameters.Append p


	cm.Execute rowschanged

	if rowschanged = 1 then
		returnValue = "1"
		cn.committrans
	else
		returnValue = "0"
		cn.rollbacktrans
	end if


	set cm = nothing
	cn.close	
	set cn = nothing

    response.Write returnValue
</script>

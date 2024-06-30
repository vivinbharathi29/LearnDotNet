<% 
    if left(lcase(Session("LoggedInUser")),9)<>"auth\" and left(lcase(Session("LoggedInUser")),12)<>"auth\" and left(lcase(Session("LoggedInUser")),5)<>"auth\" and (Instr(Session("LoggedInUser"), "@hp.com") = 0) then
			Response.Redirect "/Excalibur/NoAccess.asp?Level=0"
   end if
%>
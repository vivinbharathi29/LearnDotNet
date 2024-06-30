<%@ Language=VBScript %>
<!-- #include file = "includes/noaccess.inc" -->
<html>
  <head>
    <title>test</title>
    <script ID="clientEventHandlersJS" LANGUAGE="javascript">
      <!--

      function hello()
      {
        var iStart = -1;
        var iStop = -1;
        var doit = "yes";
        var url = String(window.location);

        iStart = url.indexOf("?id=");
        if (iStart < 0)
        {
          iStart = url.indexOf("&id=");
        }
        if (iStart >= 0)
        {
          iStop = url.indexOf("&", iStart);
        }
        if (doit != "")
        {
          if (iStart < 0)
          {
            url = url + "?id=1";
          }
          else
          {
            if (iStop < 0)
            {
              url = url.substring(0, iStart + 1) + "id=1";
            }
            else
            {
              url = url.substring(0, iStart + 1) + "id=1" + url.substr(iStop);
            }
          }
          //alert(url);
        }
        else
        {
        }
      }

      function setUrlKey(u, k, v)
      {
        var arKeys;
        var arUrl;
        var arKeyValue;
        var i;
        var url = String(oldurl);
        var oldurl = String(u);
        var key = String(k);
        var value = String(v);
        var bFound;

        // check to be sure we were given an url
        if (oldurl.length)
        {
          // separate location from parameters
          arUrl = oldurl.split("?");
          url = arUrl[0];
          // do we have any parameters, and we were given a key?
          if ((arUrl.length > 1) && (key.length))
          {
            // split key/value pairs
            arKeys = arUrl[1].split("&");
            bFound = false;
            for (i = 0; (!bFound) && (i < arKeys.length); i++)
            {
              arKeyValue = arKeys[i].split("=");
              if (arKeyValue[0] == key)
              {
                // found the key, either replace the value, or remove the key
                bFound = true;
                if (value.length)
                {
                  // replace the value
                  arKeys[i] = arKeyValue[0] + "=" + value;
                }
                else
                {
                  // remove key/value pair
                  //arKeys = arKeys.splice(i, 1);
                  arKeys.splice(i, 1);
                }
              }
            }
            if ((! bFound) && (value.length))
            {
              // didn't find the key, just add it to the array of key/value pairs
              arKeys.push(key + "=" + value);
            }
            if (arKeys.length)
            {
              // add the first key/value pair special case
              url = url + "?" + arKeys[0];
            }
            // append the rest of the key/value pairs
            for (i = 1; i < arKeys.length; i++)
            {
              url = url + "&" + arKeys[i];
            }
          }
          else if (key.length && value.length)
          {
            // no parameters there yet, just append
            url = url + "?" + key + "=" + value;
          }
        }
        return url;
      }
      //-->
    </script>
  </head>
  <STYLE>
    TEXTAREA
    {
      FONT-WEIGHT: normal;
      FONT-SIZE: x-small;
      FONT-FAMILY: Verdana;
    }
    A:visited
    {
      COLOR: blue
    }
    A:hover
    {
      COLOR: red
    }

    TD.HeaderButton
    {
      FONT-SIZE: xx-small;
      FONT-FAMILY: Verdana;
      FONT-WEIGHT: bold;
      COLOR: White;
    }
  </STYLE>
  <body bgcolor="ivory" LANGUAGE="javascript" onload="return hello()">
<%
function TemplateStrings(ID, RequestType, NewName, strProdFamId, strProdVerId, strAdd, strRemove, EmployeeId)

  on error resume next

  dim bContinue
  dim cm
  dim cn
  dim drArray
  dim i
  dim newId
  dim rs
  dim strResult

  select case RequestType

    case 1 'Rename
      strResult = ""
      set cn = server.createobject("ADODB.Connection")
      set cm = server.CreateObject("ADODB.Command")

      cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
      cn.open

      cn.begintrans

      set cm = server.CreateObject("ADODB.Command")
      cm.ActiveConnection = cn

      cm.CommandText = "up_ct_Rename"
      cm.CommandType =  adCmdStoredProc

      set p = cm.CreateParameter("@ID", adInteger,  adParamInput)
      p.value = ID
      cm.Parameters.Append p

      set p = cm.CreateParameter("@NewName", adVarChar, adParamInput, 80)
      p.value = left(NewName,80)
      cm.Parameters.Append p

      set p = cm.CreateParameter("@EmployeeID", adInteger,  adParamInput)
      p.value = EmployeeID
      cm.Parameters.Append p

      cm.Execute rowschanged

      if rowschanged = 1 then
        strResult = "1"
        cn.committrans
      else
        cn.rollbacktrans
      end if

      set cm = nothing
      set cn = nothing
      TemplateStrings = strResult

    case 2 'Delete
      strResult = ""
      set cn = server.createobject("ADODB.Connection")
      set cm = server.CreateObject("ADODB.Command")

      cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
      cn.open

      cn.begintrans

      set cm = server.CreateObject("ADODB.Command")
      cm.ActiveConnection = cn

      cm.CommandText = "up_ct_Delete"
      cm.CommandType =  adCmdStoredProc

      set p = cm.CreateParameter("@ID", adInteger,  adParamInput)
      p.value = ID
      cm.Parameters.Append p

      cm.Execute rowschanged

      if rowschanged = 1 then
        strResult = "1"
        cn.committrans
      else
        cn.rollbacktrans
      end if

      set cm = nothing
      set cn = nothing
      TemplateStrings = strResult

    case 3 'Update
      strResult = ""
      set cn = server.createobject("ADODB.Connection")
      set cm = server.CreateObject("ADODB.Command")

      cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
      cn.open
      cn.begintrans

      bContinue = true
      drArray = split(strAdd, ",")
      for i = lbound(drArray) to ubound(drArray)
        if (bContinue) then
          set cm = server.CreateObject("ADODB.Command")
          cm.ActiveConnection = cn
          cm.CommandText = "up_ct_AddLink"
          cm.CommandType =  adCmdStoredProc

          set p = cm.CreateParameter("@id", adInteger,  adParamInput)
          p.value = ID
          cm.Parameters.Append p

          set p = cm.CreateParameter("@drid", adInteger,  adParamInput)
          p.value = drArray(i)
          cm.Parameters.Append p

          cm.Execute rowschanged

          bContinue = bContinue and (rowschanged > 0)
          set cm = nothing
        else
          break
        end if
      next
      drArray = nothing

      if (bContinue) then
        drArray = split(strRemove, ",")
        for i = lbound(drArray) to ubound(drArray)
          if (bContinue) then
            set cm = server.CreateObject("ADODB.Command")
            cm.ActiveConnection = cn
            cm.CommandText = "up_ct_RemoveLink"
            cm.CommandType =  adCmdStoredProc

            set p = cm.CreateParameter("@id", adInteger,  adParamInput)
            p.value = ID
            cm.Parameters.Append p

            set p = cm.CreateParameter("@drid", adInteger,  adParamInput)
            p.value = drArray(i)
            cm.Parameters.Append p

            cm.Execute rowschanged
            bContinue = bContinue and (rowschanged > 0)
            set cm = nothing
          else
            break
          end if
        next
        drArray = nothing
      end if

      if (bContinue) then
        set cm = server.CreateObject("ADODB.Command")
        cm.ActiveConnection = cn
        cm.CommandText = "up_ct_Update"
        cm.CommandType =  adCmdStoredProc

        set p = cm.CreateParameter("@ID", adInteger,  adParamInput)
        p.value = ID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@EmployeeID", adInteger,  adParamInput)
        p.value = EmployeeID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductFamilyID", adInteger,  adParamInput)
        p.value = strProdFamId
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductVersionID", adInteger,  adParamInput)
        p.value = strProdVerId
        cm.Parameters.Append p

        cm.Execute rowschanged

response.write "<br>up_ct_update " & ID & ", " & EmployeeID & ", " & strProdFamId & ", " & strProdVerId &_
  " (" & rowschanged & ")"
        bContinue = bContinue and (rowschanged > 0)
        if (bContinue) then
          strResult = "1"
          cn.committrans
        else
          cn.rollbacktrans
        end if
      end if

      set cm = nothing
      set cn = nothing
      TemplateStrings = strResult

    case 4 'Add
      strResult = ""
      set cn = server.createobject("ADODB.Connection")
      set cm = server.CreateObject("ADODB.Command")

      cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
      cn.open
      cn.begintrans

      bContinue = true
      if (bContinue) then
        set cm = server.CreateObject("ADODB.Command")
        cm.ActiveConnection = cn
        cm.CommandText = "up_ct_Add"
        cm.CommandType =  adParamReturnValue

        set p = cm.CreateParameter("@Name", adVarChar, adParamInput, 80)
        p.value = left(NewName,80)
        cm.Parameters.Append p

        set p = cm.CreateParameter("@EmployeeID", adInteger,  adParamInput)
        p.value = EmployeeID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductFamilyID", adInteger,  adParamInput)
        p.value = strProdFamId
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductVersionID", adInteger,  adParamInput)
        p.value = strProdVerId
        cm.Parameters.Append p

      	set p = cm.CreateParameter("@NewID", adInteger, adParamOutput)
      	cm.Parameters.Append p

        cm.Execute rowschanged

        bContinue = bContinue and (rowschanged > 0)
        if bContinue then
          strResult = cm("@NewID")
          newId = strResult
        else
          strResult = ""
        end if
        set cm = nothing
      end if

      drArray = split(strAdd, ",")
      for i = lbound(drArray) to ubound(drArray)
        if (bContinue) then
          set cm = server.CreateObject("ADODB.Command")
          cm.ActiveConnection = cn
          cm.CommandText = "up_ct_AddLink"
          cm.CommandType =  adCmdStoredProc

          set p = cm.CreateParameter("@id", adInteger,  adParamInput)
          p.value = newId
          cm.Parameters.Append p

          set p = cm.CreateParameter("@drid", adInteger,  adParamInput)
          p.value = drArray(i)
          cm.Parameters.Append p

          cm.Execute rowschanged

          bContinue = bContinue and (rowschanged > 0)
          set cm = nothing
        else
          break
        end if
      next
      if (bContinue) then
        cn.committrans
      else
        strResult = ""
        cn.rollbacktrans
      end if

      drArray = nothing
      set cm = nothing
      set cn = nothing
      TemplateStrings = strResult
  end select
end function

'Add
'strResult = TemplateStrings("", "4", "My Template", "147", "0", "", "", "695")
'Delete
'strResult = TemplateStrings("2", "2", "", "", "", "", "", "")
'Update
'strResult = TemplateStrings("9", "3", "", "-1", "355", "", "", "695")
'strResult = TemplateStrings("9", "3", "", "-1", "0", "", "", "695")
'response.write "<br>strResult='" & strResult & "'<br>"
%>
<script language="javascript">
  var url = window.location;
  url = setUrlKey(url, "dr", "1,2,3");
  document.write(url + "<br>");
  url = setUrlKey(url, "id", "abc");
  document.write(url + "<br>");
  url = setUrlKey(url, "dr", "1,4,5");
  document.write(url + "<br>");
  url = setUrlKey(url, "", "");
  document.write(url + "<br>");
  url = setUrlKey(url, "id", "xyz");
  document.write(url + "<br>");
  url = setUrlKey(url, "dr", "1,2,3");
  document.write(url + "<br>");
  url = setUrlKey(url, "dr", "");
  document.write(url + "<br>");
  url = setUrlKey(url, "id", "");
  document.write(url + "<br>");
</script>
  </body>
</html>

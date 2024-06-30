<!--#include file="../_ScriptLibrary/jsrsServer.inc"-->
<% jsrsDispatch( "TemplateStrings" ) %>

<script runat="server" language="vbscript">

function TemplateStrings(ID, RequestType, NewName, strTestCatId, strProdFamId, strProdVerId, strAdd, strRemove, EmployeeId)

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

      cm.CommandText = "up_tmt_ct_Rename"
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

      cm.CommandText = "up_tmt_ct_Delete"
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
          cm.CommandText = "up_tmt_ct_AddLink"
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
            cm.CommandText = "up_tmt_ct_RemoveLink"
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
        cm.CommandText = "up_tmt_ct_Update"
        cm.CommandType =  adCmdStoredProc

        set p = cm.CreateParameter("@ID", adInteger,  adParamInput)
        p.value = ID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@EmployeeID", adInteger,  adParamInput)
        p.value = EmployeeID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@TestCategoryID", adInteger,  adParamInput)
        p.value = strTestCatId
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductFamilyID", adInteger,  adParamInput)
        p.value = strProdFamId
        cm.Parameters.Append p

        set p = cm.CreateParameter("@ProductVersionID", adInteger,  adParamInput)
        p.value = strProdVerId
        cm.Parameters.Append p

        cm.Execute rowschanged

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
        cm.CommandText = "up_tmt_ct_Add"
        cm.CommandType =  adParamReturnValue

        set p = cm.CreateParameter("@Name", adVarChar, adParamInput, 80)
        p.value = left(NewName,80)
        cm.Parameters.Append p

        set p = cm.CreateParameter("@EmployeeID", adInteger,  adParamInput)
        p.value = EmployeeID
        cm.Parameters.Append p

        set p = cm.CreateParameter("@TestCategoryID", adInteger,  adParamInput)
        p.value = strTestCatId
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
          cm.CommandText = "up_tmt_ct_AddLink"
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

</script>

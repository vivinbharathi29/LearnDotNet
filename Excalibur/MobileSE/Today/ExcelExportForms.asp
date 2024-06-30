<SCRIPT LANGUAGE=VBScript RUNAT=Server>
'Process of x-www-form-urlencoded POST data
'Using BinaryRead, v1.00
'2001 Antonin Foller, PSTRUH Software, http://www.pstruh.cz
Function GetForm
  'Dictionary which will store source fields.
  Dim FormFields
  Set FormFields = CreateObject("Scripting.Dictionary")
  'If there are some POST source data
  If Request.Totalbytes>0 And _
    Request.ServerVariables("HTTP_CONTENT_TYPE") = _
    "application/x-www-form-urlencoded" Then

    'Read the data
    Dim SourceData
    'SourceData = Request.BinaryRead(Request.Totalbytes) 
    
    '********************************************************
    
    dim BytesLeft 
    dim TotalSize
    dim CurrentBytes 

    TotalSize = Request.TotalBytes             
    ChunkSize = 64*1024                        
    If ChunkSize > TotalSize Then              
        ChunkSize = TotalSize 
    End If
    
    BytesLeft = TotalSize                      
    SourceData = ""
    Do While BytesLeft > 0                     
    
        If BytesLeft < ChunkSize Then          
            ChunkSize = BytesLeft              
        End If
        
        SourceData = sourcedata & RSBinaryToString(Request.BinaryRead(ChunkSize))
        BytesLeft = BytesLeft - ChunkSize                                

    Loop    
    
    
    
    
    '*******************************************************

    'Convert source binary data To a string
    'SourceData = RSBinaryToString(SourceData)

    'Form fields are separated by "&"
    SourceData = split(SourceData, "&")
    Dim Field, FieldName, FieldContents
  
    For Each Field In SourceData
      'Field name And contents is separated by "="
      Field = split(Field, "=")
      FieldName = URLDecode(Field(0))
      FieldContents = URLDecode(Field(1))
      'Add field To the dictionary
      FormFields.Add FieldName, FieldContents
    Next
  end if'Request.Totalbytes>0
  Set GetForm = FormFields
End Function

Function URLDecode(ByVal What)
'URL decode Function
'2001 Antonin Foller, PSTRUH Software, http://www.pstruh.cz
  Dim Pos, pPos

  'replace + To Space
  What = Replace(What, "+", " ")

  on error resume Next
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  If err = 0 Then 'URLDecode using ADODB.Stream, If possible
    on error goto 0
    Stream.Type = 2 'String
    Stream.Open

    'replace all %XX To character
    Pos = InStr(1, What, "%")
    pPos = 1
    Do While Pos > 0
      Stream.WriteText Mid(What, pPos, Pos - pPos) + _
        Chr(CLng("&H" & Mid(What, Pos + 1, 2)))
      pPos = Pos + 3
      Pos = InStr(pPos, What, "%")
    Loop
    Stream.WriteText Mid(What, pPos)

    'Read the text stream
    Stream.Position = 0
    URLDecode = Stream.ReadText

    'Free resources
    Stream.Close
  Else 'URL decode using string concentation
    on error goto 0
    'UfUf, this is a little slow method. 
    'Do Not use it For data length over 100k
    Pos = InStr(1, What, "%")
    Do While Pos>0 
      What = Left(What, Pos-1) + _
        Chr(Clng("&H" & Mid(What, Pos+1, 2))) + _
        Mid(What, Pos+3)
      Pos = InStr(Pos+1, What, "%")
    Loop
    URLDecode = What
  End If
End Function


Function RSBinaryToString(Binary)
  'Antonin Foller, http://www.pstruh.cz
  'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY)
  'to a string (BSTR) using ADO recordset
  
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function
</SCRIPT>


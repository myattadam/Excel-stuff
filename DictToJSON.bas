Option Explicit

Sub saveJSON(filename As String, entity As Variant)
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fs.OpenTextFile(ActiveWorkbook.Path & "\" & filename, ForWriting, True)
    ts.Write toJSON(entity)
    ts.Close
End Sub

Function toJSON(entity As Variant) As String
    
    Dim index As Long
    Dim s As String
    
    If IsArray(entity) Then
        s = s & "["

        For index = LBound(entity) To UBound(entity)
            s = s & toJSON(entity(index))
            If index < UBound(entity) Then s = s & ","
        Next

        s = s & "]"
    
    Else
    
        Select Case TypeName(entity)
            Case "Empty"
                s = s & "null"
            
            Case "Integer", "Long", "Single", "Double"
                s = s & entity
                
            Case "String"
                s = s & """" & entity & """"
                
            Case "Dictionary"
                s = s & "{"
                
                Dim keylist As Variant
                keylist = entity.Keys
        
                For index = LBound(keylist) To UBound(keylist)
                    s = s & """" & keylist(index) & """:"
                    s = s & toJSON(entity(keylist(index)))
                    If index < UBound(keylist) Then s = s & ","
                Next
                
                
                s = s & "}"
            
        End Select
    End If
    
    toJSON = s
    
End Function

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
                
            Case "Boolean"
                s = s & """" & entity & """"
                
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
                     
            Case "Collection"
                s = s & "["
                
                For index = 1 To entity.count
                    s = s & toJSON(entity(index))
                    If index < entity.count Then s = s & ","
                Next
                
                s = s & "]"
            
            Case Else
                s = s & """" & TypeName(entity) & """"
            
        End Select
    End If
    
    toJSON = s
    
End Function

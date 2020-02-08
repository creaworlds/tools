Public Function PasswordMaker(length As Integer)
    Dim result As String
    Dim counter As Integer
    
    Dim passphrase As String
    passphrase = "abcdefghijklmnopqrstuvwxyz1234567890"
    
    Dim lenPass As Integer
    lenPass = Len(passphrase)
    
    result = ""
    
    Do While Len(result) < length
        Math.Randomize
        
        Dim rand
        rand = CInt(Math.Rnd * lenPass)
        
        If rand > 0 And rand <= lenPass Then
            result = result & Mid(passphrase, rand, 1)
        End If
    Loop
    
    PasswordMaker = result
End Function


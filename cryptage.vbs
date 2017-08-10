Sub TestChiffrement()

    Dim K

    Dim M

    Dim C

    Dim D

 
    K = "teseatidnestunpetitcachotier"

 
    M = InputBox ("Choissit le message de crypter?")     ' M cest le message normal

 
    C = ""

    For i = 1 To Len(M)

        C = C & Chr(Asc(Mid(K, i, 1)) Xor Asc(Mid(M, i, 1)))

    Next

    MsgBox "Chiffré: " & C
    Inputbox "Le code chiffré","Parle By ABOAT",C      ' C cest le code chiffré

 
    D = ""

    For i = 1 To Len(C)

        D = D & Chr(Asc(Mid(K, i, 1)) Xor Asc(Mid(C, i, 1)))

    Next

 
    Inputbox "Le code Déchiffré","Parle By ABOAT",D   ' D le code Déchiffré

End Sub

Call TestChiffrement()
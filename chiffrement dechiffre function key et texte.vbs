Function Chiffrement(Message, Clef)

    Dim i
    Dim C 
    Message = InputBox ("Ecrit le texte")
    Clef = InputBox ("Ecrit la clef")
 
    For i = 1 To Len(Message)

        C = C & Chr(Asc(Mid(Clef, i Mod (Len(Message) + 1), 1)) Xor Asc(Mid(Message, i, 1)))

    Next

    Chiffrement = C
    MsgBox Chiffrement

End Function

 
Function Dechiffrement(MsgCode, Clef)
    MsgCode = InputBox ("Ecrit le texte")
    Clef = InputBox ("Ecrit la clef")

    Dechiffrement = Chiffrement(MsgCode, Clef)
    MsgBox Dechiffrement

End Function


Call Chiffrement(Message, Clef)
Call Dechiffrement(MsgCode, Clef)  

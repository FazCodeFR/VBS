
Menu = InputBox ("1 = encoder" & vbCr & vbCr & "2 = Decoder")  
If Menu = 1 Then Call Encode ()
If Menu = 2 Then Call Decode ()

 
 
Sub Encode ()
strText = InputBox ("Le texte a encoder : ")
lngSeed = InputBox ("Quel est le code a mettre pour l'encodage : ")
MsgBox "Le texte :" & vbCr & strText & vbCr & "Avec comme code : " & vbCr & lngSeed & vbCr & "Deviens : " & vbCr & vbCr & Scramble (strText, lngSeed)
InputBox "","",Scramble (strText, lngSeed)
End Sub 


Sub Decode () 
strText = InputBox ("Le texte a decoder : ")
lngSeed = InputBox ("Quel est le code a mettre : ")
MsgBox "L'encode:" & vbCr & strText & vbCr & "Avec comme code : " & vbCr & lngSeed & vbCr & "Deviens en claire : " & vbCr & vbCr & Scramble (strText, -lngSeed)
End sub


 '---- Scramble ----'
 Function Scramble (strText, lngSeed)
     Dim L,intRand,bytASC
 
     '---- Force seeded random mode 
     Rnd(-1)
 
     '---- Set (positive) seed 
     Randomize ABS(lngSeed)
 
     '---- Scan through string
     For L = 1 To Len(strText)
 
         '---- Get ASC of char
         bytASC=Asc(Mid(strText, L))
 
         '---- Fix for quotes (tilde to quote)
         If bytASC=126 then bytASC=34
 
         '---- Add a random value from -80 to 80, encode\decode is decided by the seed's sign
         intRand = bytASC + ((Int(Rnd(1) * 160) - 80) * SGN(lngSeed)) 
 
         '---- Cycle char between 32 and 125 (with carry)
         If intRand <= 31 Then 
             intRand = 125 - (31 - intRand)
         ElseIf intRand >= 126 Then
             intRand = 32 + (intRand - 126)
         End If
 
         '---- Fix for quotes (quote to tilde)
         If intRand=34 then intRand=126
 
         '---- Output string
         Scramble = Scramble & Chr(intRand)
     Next
 End Function
 
 
 
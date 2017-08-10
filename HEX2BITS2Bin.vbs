'Juste deux petite fonctions très très utile pour crypté :
'La conversion valeur => binaire...
Call CrypteTest()

Function ValToBinString
  For i = 0 To 7
    c = Vl Mod 2
    If c > 0 Then
      c = 1
      Vl = (Vl - 1) / 2
    Else
      Vl = Vl / 2
    End If
    ValToBinString = CStr(c) & ValToBinString  
    
  Next
End Function

'...et la conversion binaire => valeur
Function BinToVal
  For i = 1 To 8
    BinToVal = BinToVal + (2 ^ (8 - i)) * Val(Mid(Sbit, i, 1))
  Next
End Function

'Bordélisation :
Sub GenAlea
  ReDim TblDeSortie(128)
  'génère une suite pseudosinusoidale pour éviter les redondance linéaire
  For i = 1 To 128
    TblDeSortie(i) = CByte(Int(Sin(i) * 127) + 128)
  Next
End Sub
'exemple bordéliseur du tutorial :
Sub CrypteTest()
  Dim TableAlea()
  GenAlea TableAlea()
  TestString = Space(6)
  For i = 1 To Len(TestString)
    Calc = Asc(Mid(TestString, i, 1)) + 3 + TableAlea(i)
    If Calc > 255 Then Calc = Calc - 255
    StrCrypte = StrCrypte & Chr(Calc)
  Next
  MsgBox TestString 
  MsgBox StrCrypte
End Sub

'Plusieurs Passes, en fonction de la clé
Sub CrypteTest2
  Dim TableAlea()
  GenAlea TableAlea()
  a = 0
  TestString = Space(6)
  For i = 1 To Len(TestString)
    For k = 0 to Asc(Mid(Cle,i,1))
      a = a + 1
      If a > Ubound(TableAlea) Then a = 1
      Calc = Asc(Mid(TestString, i, 1)) Xor 3 Xor TableAlea(a)
    Next
    StrCrypte = StrCrypte & Chr(Calc)
  Next
End Sub

'le BitShift
Private Sub BitShift
'décalage de bit dans 1 byte
Dim eBit(8)
Dim i

    'conversion octet vers bits (les bits sont dans un tableau)
    eBit(1) = Abs(CBool(myByte And 128))
    eBit(2) = Abs(CBool(myByte And 64))
    eBit(3) = Abs(CBool(myByte And 32))
    eBit(4) = Abs(CBool(myByte And 16))
    eBit(5) = Abs(CBool(myByte And 8))
    eBit(6) = Abs(CBool(myByte And 4))
    eBit(7) = Abs(CBool(myByte And 2))
    eBit(8) = Abs(CBool(myByte And 1))
    
    'décale "vers a gauche" ou "vers la droite" selon la valeur de Decalage
    If Decalage > 0 Then
        For i = 1 To Decalage
            r = eBit(8): eBit(8) = eBit(7): eBit(7) = eBit(6): eBit(6) = eBit(5)
            eBit(5) = eBit(4): eBit(4) = eBit(3): eBit(3) = eBit(2): eBit(2) = eBit(1): eBit(1) = r
        Next
    ElseIf Decalage < 0 Then
        For i = 1 To Abs(Decalage)
            r = eBit(1): eBit(1) = eBit(2): eBit(2) = eBit(3): eBit(3) = eBit(4): eBit(4) = eBit(5)
            eBit(5) = eBit(6): eBit(6) = eBit(7): eBit(7) = eBit(8): eBit(8) = r
        Next
    End If
        
    'conversion bits vers octet
    myByte = (2 ^ 7) * eBit(1) + (2 ^ 6) * eBit(2) + (2 ^ 5) * eBit(3) + (2 ^ 4) * eBit(4) + _
             (2 ^ 3) * eBit(5) + (2 ^ 2) * eBit(6) + (2 ^ 1) * eBit(7) + (2 ^ 0) * eBit(8)

End Sub

'CheckSum
Function XorCheck
Dim Mediat
Mediat = Val("&h" & Left(Hex(Len(TxZ)) & "0", 2) & "&")
For i = 1 To Len(TxZ)
    Mediat = i + Mediat + Mediat Xor Asc(Mid(TxZ, i, 1))
Next
XorCheck = Mediat
End Function

Call ValToBinString ()

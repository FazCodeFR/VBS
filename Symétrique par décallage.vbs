Set FSO = CreateObject("Scripting.FileSystemObject")
Set AG = WScript.Arguments

For I = 0 to AG.Count -1
 Set T = FSO.GetFile(AG(I))
 Do
  QUE = Inputbox("    Voici un programme de Cryptage/Décryptage symétrique à algorithme de substitution par décallage."&VbCrLf&VbCrLf&"Que voulez-vous faire ?"&VbcrLf&" 1) Crypter"&VbcrLf&" 2) Décrypter"&VbcrLf&VbcrLf&"Taper 1 ou 2","Cryptage/Décryptage de "& Chr(34) & T.Name & Chr(34))
  If QUE = "1" Then Q1
  If QUE = "2" Then Q2
  If QUE = Empty Then
   M = MsgBox ("Voulez-vous quittez ce programme ?",32+4,"Cryptage/Décryptage symétrique...")
   If M = VbYes Then WScript.Quit
  End If
 Loop Until QUE <> Empty And (QUE = "1" Or QUE = "2")
 If I = AG.Count -1 Then WScript.Quit
Next

MsgBox "Il faut faire un glisser-déposer sur ce programme avec le fichier choisi.",48+0,"Attention"

Sub Q1
KEY = InputBox ("2) Saisissez une clé pour le cryptage de vos données.","Cryptage symétrique de "& Chr(34) & T.Name & Chr(34))
If KEY = Empty Then MsgBox "Pas de clé, pas de cryptage.",48+0,"Fin" : WScript.Quit
CRYPT_DECRYPT "chiffré",1,KEY
End Sub

Sub Q2
KEY = InputBox ("2) Saisissez une clé pour le décryptage de vos données.","Décryptage symétrique de "& Chr(34) & T.Name & Chr(34))
If KEY = Empty Then MsgBox "Pas de clé, pas de décryptage.",48+0,"Fin" : WScript.Quit
CRYPT_DECRYPT "déchiffré",-1,KEY
End Sub

Function CRYPT_DECRYPT(DECH,S,KE)
K = ""
X = 1
For A = 1 To Len(KE)
 K = K & Asc(Mid(KE,X,1))
 X = X + 1
Next

Set CODE = FSO.OpenTextFile(AG(I),1)
CO = CODE.ReadAll
CODE.Close

For F = 1 To 3
 Y = 1
 CF = ""
 For V = 1 To -Int(-Len(CO)/Len(K))
  X = 1
  For G = 1 To Len(K)
   K2 = CInt(Mid(K,X,1))
   C = Asc(Mid(CO,Y,1))
   W = C + K2 * S
   If W > 255 Then W = W - 255
   If W = 0 Then W = 255
   If W < 0 Then W = 255 + W
   CF = CF & Chr(W)
   X = X + 1
   Y = Y + 1
   If Y > Len(CO) Then Exit For
  Next
 Next
 CO = CF
 K = Len(K) & strReverse(K) & Mid(K,1,1)
Next

ND = T.ParentFolder.Path & "\" & FSO.GetBaseName(T) & "(" & DECH & ")" & "." & FSO.GetExtensionName(T)
INFO = MsgBox ("3) Fichier " & DECH & " avec la clé : " &VbCrLf& "-->" & KE &VbCrLf&VbCrLf& "Voici le début du fichier " & DECH &" :"&VbCrLf& "-->" & Mid(CF,1,50)&VbCrLf&VbCrLf&"Voulez-vous enregister ce fichier avec le nom prévu par défaut :"&VbCrLf& "--> ...(" & DECH & ").txt     (cliquez sur NON pour en choisir un autre).",4,"Opération terminée")
If INFO = VbYes Then
 N = ND
Else
 Do
  N = InputBox ("Tapez son URL complète, sauf si vous voulez qu'il se trouve dans le même répertoire que ce programme. Dans ce cas, il suffit simplement de saisir son NOM.","Choix du nom pour le fichier " & DECH)
  If N = Empty Then N = ND : MsgBox "Nom par défaut attribué au fichier",64+0,"Info" : Exit Do
  V = InStr(StrReverse(N),"\")
  F = Mid(N,1,Len(N)-V)
  If FSO.FolderExists(F) Or V = 0 Then Exit Do
  MsgBox "URL invalide",16+0,"Erreur"
 Loop
End If

Set CODE = FSO.CreateTextFile(N,True)
CODE.Write CO
CODE.Close
End Function
Option Explicit

Public minuscules   'Les caracteres minuscules autorisés
Public majuscules   'Les caracteres majuscules autorisés
Public chiffres     'Les chiffres autorisés

'
'Procedure qui initialise les caracteres autorises
'
Sub init()
	minuscules="abcdefghijklmnopqrstuvwxyz"
	majuscules=UCase(minuscules)
	chiffres="0123456789"
End sub

'
'Fonction qui permet de generer un mot de passe de nbChar caracteres
'
Function GenererMp(nbChar)
	Dim i      'compteur de boucle
	Dim pMin   'la probabilite d'avoir une minuscule
	Dim pMaj   'la probabilite d'avoir une majuscule
	Dim pNum   'la probabilite d'avoir un chiffre
	Dim total  'le nombre total de caracteres
	Dim res    'le resultat de la fonction
	
	res=""
	total=Len(minuscules)+Len(majuscules)+Len(chiffres)
	pMaj = (Len(majuscules)/total)
	pMin = (Len(minuscules)/total)
	pNum = (Len(chiffres)/total)
	
	For i=1 To nbChar
		Randomize
		If rnd<pMin Then
			res = res & SelectRandomChar(minuscules)
		ElseIf rnd<pMin+pMaj Then
			res = res & SelectRandomChar(majuscules)
		ElseIf rnd<pMin+pMaj+pNum Then
			res = res & SelectRandomChar(chiffres)
		End If		
	Next
	GenererMp=res
End Function

Function SelectRandomChar(caracteres)
	dim alea
	Randomize
	alea = Int(Len(caracteres)*rnd+1)
	SelectRandomChar = Mid(caracteres, alea, 1)
End Function

Sub lanceur()

	On Error Resume Next
                  Const ForWriting = 8
	Dim nbChar 'Nombre de caracteres voulus par l'utilisateur
	Dim fin    'Choix final du l'utilisateur
                 Dim fso, f    
	fin = ""
	init()
	nbChar=20
                 salut = MsgBox ( GenererMP(nbChar) )
                 Set fso = CreateObject("Scripting.FileSystemObject") 
                 Set f = fso.OpenTextFile("C:\a.txt", ForWriting,true) 
                 f.Write vbnewline
                 f.write("' " & GenererMP(nbChar))
End Sub


Call lanceur()
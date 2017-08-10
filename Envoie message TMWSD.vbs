On Error Resume Next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.Toolbar=0
IE.StatusBar=0
IE.Height=560
IE.Width=1000
IE.Top=0
IE.Left=0
IE.Resizable=0
IE.navigate "https://www.thismessagewillselfdestruct.com/" 

Texte = Inputbox ("Quel est le texte a mettre ?","Le texte : ","Salut YB c'est ABOAT !")
MDP = Inputbox ("Quel est le mdp souhaité ?", "Le mot de passe : ")

While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 1000
Set Helem = IE.Document.All.Item("message_body") 'message[body] Marche aussi
Helem.Value=Texte

Set Helem = IE.Document.All.Item("message_password")
Helem.Value=MDP
WScript.Sleep 1000
Set Helem = IE.Document.All.Item("commit")
Helem.click


While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 2000

' IE.document.all.namedItem("message_message_urls_attributes_0_id").Value 'donne bien la valeur  du bouton
' IE.Document.getElementByID("message_message_urls_attributes_0_id").Value 'donne bien la valeur  du bouton
' IE.Document.All("message_message_urls_attributes_0_id").Value 'donne bien la valeur  du bouton
' IE.document.all.namedItem("message[message_urls_attributes][0][id]").Value 'donne bien la valeur du bouton
'IE.Document.All("message[message_urls_attributes][0][id]").Value 'donne bien la valeur du bouton


Lien1 = IE.document.all.namedItem("message_url_" & IE.Document.getElementByID("message_message_urls_attributes_0_id").Value).outertext
If Len(Lien1) < 2 Then 'Moin de 2 caractere
	Lien1 = IE.document.all.namedItem("message_url_" & IE.Document.All("message_message_urls_attributes_0_id").Value).outertext
		If Len(Lien1) < 2 Then 'Moin de 2 caractere
			Lien1 = IE.document.all.namedItem("message_url_" & IE.document.all.namedItem("message[message_urls_attributes][0][id]").Value).outertext
		End If 
End If


InputBox "Le lien généré est : " & Lien1,"Lien généré : ",Lien1
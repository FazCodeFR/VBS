'Synchronisation de l'horloge de l'ordinateur avec un serveur SNTP
'Auteur : Brughes
'Vous pouvez écouter/télécharger ma musique en open source : http://soundcloud.com/cyberflaneur ou http://www.jamendo.com/fr/artist/Brughes

Option Explicit

Dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run "net time /setsntp:ntp.imag.fr", 0, True
'ou
'WshShell.Run "net time /setsntp:fr.pool.ntp.org", 0, True

Set WshShell = Nothing

WScript.Quit

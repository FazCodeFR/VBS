'le plus bas (level 1) autorise les macros et l'exécution d'objets OLE sans avertissement.

Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security\Level", 1, "REG_DWORD"
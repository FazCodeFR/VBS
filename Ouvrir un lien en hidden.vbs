URL = "https://www.youtube.com/embed/tLDRN9HafZc?rel=0;showinfo=0;controls=0;iv_load_policy=3;autoplay=1;"
Set ie = CreateObject("InternetExplorer.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
ie.Navigate (URL) 
ie.Visible=false
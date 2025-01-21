Dim objShell, objIE
Set objShell = CreateObject("WScript.Shell")
Set objIE = CreateObject("InternetExplorer.Application")

' Configure the IE window
With objIE
    .Visible = True
    .AddressBar = True
    .StatusBar = True
    .ToolBar = True
    .MenuBar = True
    .Resizable = True
    .Width = 400
    .Height = 300
    .Top = 200
    .Left = 300
    .Navigate "google.com"
End With
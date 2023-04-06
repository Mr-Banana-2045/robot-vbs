Dim speech
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim currentTime
Set name = CreateObject("WScript.Network")

msg = inputbox("i am robot speaker","bot")
if msg = "hello" Then
   lo = "salam dooste man"
   set speech=CreateObject("sapi.spvoice")
   speech.speak lo
elseif msg = "help" Then
    up=MsgBox("hello - how are you - thanks - info - on google - info ip - time - who am i",10,"input commands")
elseif msg = "how are you" Then
    mo = "man khoobam too cheetoory"
    set speech=CreateObject("sapi.spvoice")
   speech.speak mo
elseif msg = "thanks" Then
     sho = "khahesh mycoonam"
     set speech=CreateObject("sapi.spvoice")
     speech.speak sho
elseif msg = "info" Then
     shoo = "man yek robot sookhaangoo hastaam man javab shoomaro beh tooree sootty miidam"
     set speech=CreateObject("sapi.spvoice")
     speech.speak shoo
elseif msg = "on google" Then
     shooo = "google shooma amadast"
     set speech=CreateObject("sapi.spvoice")
     speech.speak shooo
     WshShell.Run "https://google.com"
elseif msg = "info ip" Then
     Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
     Set colAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
     For Each objAdapter in colAdapters
     If Not IsNull(objAdapter.IPAddress) Then
     For Each strIPAddress in objAdapter.IPAddress
     shooo = "ip shooma "& strIPAddress
     set speech=CreateObject("sapi.spvoice")
     speech.speak shooo
     Next
     End If
     Next
elseif msg = "time" Then
     currentTime = Time
     sh = "zoman mahaly "&currentTime
     set speech=CreateObject("sapi.spvoice")
     speech.speak sh
elseif msg = "who am i" Then
     nom = name.ComputerName
     shw = "nam shooma hast "&nom
     set speech=CreateObject("sapi.spvoice")
     speech.speak shw
elseif msg = "on cmd" Then
     WshShell.Run "cmd.exe"
     shp = "cmd shooma amadast"
     set speech=CreateObject("sapi.spvoice")
     speech.speak shp
else
     op=msgbox("bye",10,"off bot")
end if
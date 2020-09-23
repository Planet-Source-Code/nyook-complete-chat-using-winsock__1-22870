Attribute VB_Name = "AllgFunktionen"
Public Sub FilterClientList(Teil As String) 'diese prozedur liest aus "Teil" die informationen über die Clients aus, und aktualisiert die clientinformationen
  DeleteClientInformations 'die Clientdaten werden gelöscht

  anzClients = 0
  
  While Not Teil = ""
    anzClients = anzClients + 1
    Client(anzClients - 1).FontColor = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
    Client(anzClients - 1).IP = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
      If InStr(1, Teil, "/") = 0 Then
        Client(anzClients - 1).NickName = Teil
        GoTo Weiter
      End If
    Client(anzClients - 1).NickName = Left(Teil, InStr(1, Teil, "/") - 1)
    Teil = Right(Teil, Len(Teil) - InStr(1, Teil, "/"))
  Wend
Weiter:
End Sub

Public Sub DeleteClientInformations()  'diese Prozedur löscht die clientinformationen
  For i = 0 To (anzClients - 1)
    Client(i).FontColor = ""
    Client(i).IP = ""
    Client(i).NickName = ""
    Client(i).LastMessage = 0
    Client(i).anzFloodMsgs = 0
  Next i
  anzClients = 0 'es sind keine Clients mehr vorhanden
End Sub

Public Function FormNickName(NickName As String) 'Diese Prozedur entfernt eventuelle Leerzeichen links und rechts vom Nicknamen und ersetzt Leerzeichen bzw. / im Nicknamen durch _ bzw. \
  NickName = Trim(NickName) 'Eventuelle Leerzeichen links und rechts entfernen
  
Anfang: 'Hier werden Leerzeichen durch _ ersetzt
  If InStr(1, NickName, " ") = 0 Then GoTo Weiter 'Wenn keine Leerzeichen mehr im Namen vorkommen, dann soll der Name auf / geprüft werden
  NickName = Left(NickName, InStr(1, NickName, " ") - 1) + "_" + Right(NickName, Len(NickName) - InStr(1, NickName, " ")) 'das erste leerzeichen von Links durch _ ersetzen
  GoTo Anfang
  
Weiter: 'Hier werden / durch \ ersetzt
  If InStr(1, NickName, "/") = 0 Then 'wenn keine / mehr im Nicknamen vorkommen
    FormNickName = NickName
    Exit Function
  End If
  NickName = Left(NickName, InStr(1, NickName, "/") - 1) + "\" + Right(NickName, Len(NickName) - InStr(1, NickName, "/")) 'den ersten / von Links durch \ ersetzen
  GoTo Weiter

End Function

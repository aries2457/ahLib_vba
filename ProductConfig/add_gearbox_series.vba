Public Sub add_gearbox_series(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_series"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_series"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Baureihen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_series " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    DoCmd.SetWarnings False
    
    'Bestehende Einträge für Baureihen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Baureihen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_series ( `name`, `properties`, `note` ) " _
         & "SELECT T_gearbox_series.`name`, T_gearbox_series.`properties`, T_gearbox_series.`note` " _
         & "FROM T_gearbox_series ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub

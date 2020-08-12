Public Sub add_classification_societies(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_classification_societies"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_classification_societies"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Classification Societies ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_classification_societies " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
    'Einträge für societies neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_classification_societies ( `name` ) " _
         & "SELECT T_gearbox_classification_societies.`name` " _
         & "FROM T_gearbox_classification_societies;"
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
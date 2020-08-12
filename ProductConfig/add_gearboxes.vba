Public Sub add_gearboxes(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearboxes"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearboxes"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Getriebetypen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearboxes " & vbLf _
         & "Voraussetzungen: gearbox_series gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Getriebetypen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
   'Einträge für Getriebetypen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearboxes ( series, `size`, design, e_motor, comment ) " _
         & "SELECT T_gearboxes.series, T_gearboxes.`size`, T_gearboxes.design, T_gearboxes.e_motor, T_gearboxes.comment " _
         & "FROM T_gearboxes ;"
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
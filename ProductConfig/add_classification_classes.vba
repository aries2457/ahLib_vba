Public Sub add_classification_classes(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_classification_classes"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "gearbox_classification_classes"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für DutyClasses ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): CurrentDB.T_gearbox_classification_classes " & vbLf _
         & "Voraussetzungen: gearbox_classification_societies gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Getriebetypen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    
   'Einträge für Getriebetypen neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO gearbox_classification_classes ( `name`, classification_society ) " _
         & "SELECT T_gearbox_classification_classes.`name`, T_gearbox_classification_classes.classification_society " _
         & "FROM T_gearbox_classification_classes ;"
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
Public Sub add_options(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_options"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preiskomponenten"
    sTab1 = "gearbox_options"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Optionen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): T_gearbox_options, " & sTab0 & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Datensätze löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Optionen aus Preiskomponenten neu einfügen
    sSQL = "INSERT INTO " & sTab1 & " ( `name` ) " _
         & "SELECT " & sTab0 & ".KPID " _
         & "FROM " & sTab0 & " " _
         & "WHERE " & sTab0 & ".KPID NOT LIKE '0*' " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    'Zusätzliche Einträge für SAP-Import (XML-Datei) aus lokaler Optionstabelle einfügen
    sTab0 = "T_gearbox_options"
    sSQL = "INSERT INTO " & sTab1 & " ( `name` ) " _
         & "SELECT " & sTab0 & ".( `name` ) " _
         & "FROM " & sTab0 & " ;" 
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
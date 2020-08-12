Public Sub add_option_localization(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_option_localization"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preiskomponenten"
    sTab1 = "gearbox_option_localization"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Sprachen für Getriebeoptionen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sTab0 & vbLf _
         & "Voraussetzungen: gearbox_options " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für Sprachen löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Sprache 'en-US' neu einfügen
    sSQL = "INSERT INTO gearbox_option_localization ( option_name, `language`, localization ) " _
         & "SELECT " & sTab0 & ".KPID, 'en-US' AS lang, " & sTab0 & ".Bez_en " _
         & "FROM " & sTab0 & " INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.`name` " _
         & "ORDER BY " & sTab0 & ".KPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    'Einträge für Sprache 'de-DE' neu einfügen
    sSQL = "INSERT INTO gearbox_option_localization ( option_name, `language`, localization ) " _
         & "SELECT " & sTab0 & ".KPID, 'de-DE' AS lang, " & sTab0 & ".Bez_de " _
         & "FROM " & sTab0 & " INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.`name` " _
         & "ORDER BY " & sTab0 & ".KPID ;"
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
Public Sub add_option_prices(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_option_prices"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTab0 As String
    
    sTab0 = "T_Preise"
    sTab1 = "gearbox_option_prices"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Preise für Optionen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): " & sTab0 & vbLf _
         & "Voraussetzungen: gearbox_options " & vbLf _
         & "Nachbearbeitung: --." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für OptionsPreise neu einfügen (aktuell auf Standard-Typen begrenzt)
    sSQL = "INSERT INTO " & sTab1 & " ( series, `size`, `option`, price ) " _
         & "SELECT " & sTab0 & ".BAR, " & sTab0 & ".BGR, " & sTab0 & ".KPID, " & sTab0 & ".Preis1 " _
         & "FROM (T_StdTypen INNER JOIN " & sTab0 & " ON (T_StdTypen.BAR = " & sTab0 & ".BAR) " _
         & "AND (T_StdTypen.BGR = " & sTab0 & ".BGR) AND (T_StdTypen.BAF = " & sTab0 & ".BAF)) " _
         & "INNER JOIN gearbox_options ON " & sTab0 & ".KPID = gearbox_options.name " _
         & "WHERE (((" & sTab0 & ".Preis1) Is Not Null) And ((" & sTab0 & ".MOT) Is Null)) " _
         & "ORDER BY " & sTab0 & ".BAR, " & sTab0 & ".BGR, " & sTab0 & ".BAF, " & sTab0 & ".KPID ; "
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
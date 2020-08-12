Public Sub add_attribute_mappings(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_attribute_mappings"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sTbl(2) As String
    
    sTab1 = "gearbox_attribute_mappings"
    sTbl(0) = "T_Abmessungen_x"

    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt Mappings für neue Bemaßungen ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE." & sTbl(0) & vbLf _
         & "Voraussetzungen: gearbox_attribute_descriptions & ...mappings gefüllt " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Abmessungen von Standard-Getriebeypen neu einfügen (Quelle ProductGuideS)
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, attribute_key, gearbox_e_motor ) " _
         & "SELECT T_StdTypen.BAR, T_StdTypen.BGR, T_StdTypen.BAF, " & sTbl(0) & ".ATNAM, 0 AS MOT " _
         & "FROM " & sTbl(0) & " INNER JOIN T_StdTypen ON " & sTbl(0) & ".TYPID = T_StdTypen.TYPID ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    'Einträge für Abmessungen von L-Getriebeypen neu einfügen (Quelle ProductGuideS)
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, attribute_key, gearbox_e_motor ) " _
         & "SELECT T_StdTypen_L.BAR, T_StdTypen_L.BGR2, T_StdTypen_L.BAF, " & sTbl(0) & ".ATNAM, 0 AS MOT " _
         & "FROM " & sTbl(0) & " INNER JOIN T_StdTypen_L ON " & sTbl(0) & ".TYPID = T_StdTypen_L.TYPID ;"
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
Public Sub add_gearbox_descriptions(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_descriptions"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sTab1 As String
    
    sTab1 = "gearbox_descriptions"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Bemaßungen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Abmessungen_x" & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Neue Kurzbeschreibungen für Getriebe hinzufügen
    sSQL = "INSERT INTO " & sTab1 & " ( id, `language`, description ) " _
         & "SELECT T_gearbox_descriptions.id, T_gearbox_descriptions.lang, T_gearbox_descriptions.note " _
         & "FROM T_gearbox_descriptions " _
         & "ORDER BY T_gearbox_descriptions.id, T_gearbox_descriptions.lang ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL
        
 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: " & sModName & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        Debug.Print sMsg
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
    End With
    Resume ExitHere
End Sub
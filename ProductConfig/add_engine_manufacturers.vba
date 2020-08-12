Public Sub add_engine_manufacturers(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_engine_manufacturers"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String
    
    sTab1 = "engine_manufacturers"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Motorenhersteller ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_MotHersteller " & vbLf _
         & "Voraussetzungen: -- " & vbLf _
         & "Nachbearbeitung: ACHTUNG! Abhängige Tabellen müssen neu gefüllt werden." & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Einträge für Motorenhersteller neu einfügen (Quelle currentDB)
    sSQL = "INSERT INTO engine_manufacturers ( `name`, description ) " _
         & "SELECT T_MotHersteller.Maker, T_MotHersteller.LName " _
         & "FROM T_MotHersteller " _
         & "WHERE (((T_MotHersteller.Maker) Not Like '_*')) ;"
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
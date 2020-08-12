Public Sub add_engines(Optional bClean As Boolean = True)
    On Error GoTo ErrHandler
    Const sProcName As String = "add_engines"
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sClass As String, sFeld1 As String
    
    sTab1 = "engines"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Dieselmotoren ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): ProductGuideS-BE.T_Motoren " & vbLf _
         & "Voraussetzungen: engine_manufacturers gefüllt " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge in Zieltabelle löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Leistungmomente für Motoren neu einfügen
    sSQL = "INSERT INTO engines ( manufacturer, Type, rating_category, rating, Speed, Torque, power, Series, `application`, `Note` ) " _
         & "SELECT T_Motoren.Maker, T_Motoren.Type, T_Motoren.RatingCat, T_Motoren.Rating, T_Motoren.Speed, T_Motoren.Torque, T_Motoren.Rating, T_Motoren.Series, T_Motoren.Appl, T_Motoren.Note " _
         & "FROM T_Motoren " _
         & "WHERE (((T_Motoren.Maker) Not Like '_*')) " _
         & "ORDER BY T_Motoren.Maker, T_Motoren.Type ;"
    Debug.Print sSQL
    DoCmd.RunSQL sSQL

 ExitHere:
    Exit Sub

 ErrHandler:
    With Err
        sMsg = "Object: Modul" & vbCrLf _
             & "Procedure: " & sProcName & vbCrLf _
             & "Error: " & .Number & vbCrLf & .Description
        MsgBox sMsg, vbCritical + vbOKOnly, .Source
        Debug.Print sMsg
    End With
    Resume ExitHere
    
End Sub
Public Sub add_gearbox_description_mappings(Optional bClean As Boolean = True)
    ' Purpose:  Add description mappings for gearbox series
    ' Author:   Andreas Herrel
    ' Date:     2018-10-15; updated: 2018-10-17
    ' Inputs:   gearboxes; gearbox_descriptions; T_gearbox_descriptions
    ' Output:   gearbox_description_mappings
    ' Requires: MySQL database online

    On Error GoTo ErrHandler
    Const sProcName As String = "add_gearbox_description_mappings"
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sTab() As String
    
    sTab(1) = "gearbox_description_mappings"
    sTab(2) = "T_gearbox_descriptions"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Beschreibungen für Getriebe ein. " & vbLf _
         & "Zieltabelle: " & sTab(1) & vbLf _
         & "Quelle(n): gearboxes; gearbox_descriptions; T_gearbox_descriptions" & vbLf _
         & "Voraussetzungen: Quellen vorhanden " & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge für societies löschen
    If bClean Then
        sSQL = "DELETE " & sTab(1) & ".* FROM " & sTab(1) & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Neue Kurzbeschreibungen für Getriebetypen hinzufügen
    sSQL = "INSERT INTO " & sTab(1) & " ( gearbox_series, gearbox_size, gearbox_design, " _
         & "gearbox_e_motor, description_id, description_language ) " _
         & "SELECT gearboxes.series, gearboxes.size, gearboxes.design, gearboxes.e_motor, " _
         & "gearbox_descriptions.id, gearbox_descriptions.language " _
         & "FROM (gearbox_descriptions INNER JOIN " & sTab(2) & " ON " _
         & "(gearbox_descriptions.language = " & sTab(2) & ".lang) AND " _
         & "(gearbox_descriptions.id = " & sTab(2) & ".id)) INNER JOIN gearboxes " _
         & "ON " & sTab(2) & ".series = gearboxes.series " _
         & "ORDER BY gearboxes.series, gearboxes.size, gearboxes.design ;"
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
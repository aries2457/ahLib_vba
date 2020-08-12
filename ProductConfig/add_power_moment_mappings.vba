Public Sub add_power_moment_mappings(Optional bClean As Boolean = True)
    ' Purpose:  This procedure deletes existing records in the target table (gearbox_power_moment_mappings) and adds new mappings.
    ' Author:   Andreas Herrel
    ' FirstCreated: 2018-12-17
    ' LastUpdated:  2018-12-17
    ' Inputs:   gearboxes, gearbox_power_moments
    ' Output:   gearbox_power_moment_mappings
    ' Requirements: input tables
    ' Post processing: --

    On Error GoTo ErrHandler
    Const sProcName As String = "add_power_moment_mappings"
    Const RefLevel As Byte = 2
    Dim db1 As DAO.Database
    Dim rs1 As DAO.Recordset
    Dim sKrit As String, sMsg As String, sAns As String, sSQL As String, sFltr As String
    Dim sTab1 As String, sClass As String, sFeld1 As String
    
    sTab1 = "gearbox_power_moment_mappings"
    sMsg = "Diese Prozedur löscht vorhandene Einträge in der Zieltabelle " & vbLf _
         & "und fügt neue Daten für Classification Societies ein. " & vbLf _
         & "Zieltabelle: " & sTab1 & vbLf _
         & "Quelle(n): gearboxes, gearbox_power_moments " & vbLf _
         & "Voraussetzungen: gearboxes + gearbox_power_moments gefüllt" & vbLf _
         & "Nachbearbeitung: --" & vbLf & vbLf _
         & "Prozedur fortsetzen?"
    
    If MsgBox(sMsg, vbQuestion + vbOKCancel, sProcName) = vbCancel Then GoTo ExitHere
    
    'Bestehende Einträge löschen
    If bClean Then
        sSQL = "DELETE " & sTab1 & ".* FROM " & sTab1 & " ;"
        DoCmd.RunSQL sSQL
    End If
    
    'Mappings für Leistungmomente neu einfügen
    DoCmd.SetWarnings False
    
    sSQL = "INSERT INTO " & sTab1 & " ( gearbox_series, gearbox_size, gearbox_design, gearbox_e_motor, ratio, class ) " _
         & "SELECT gearboxes.series, gearboxes.size, gearboxes.design, gearboxes.e_motor, gearbox_power_moments.ratio, gearbox_power_moments.class " _
         & "FROM gearboxes INNER JOIN gearbox_power_moments ON (gearboxes.size = gearbox_power_moments.gearbox_size) " _
         & "AND (gearboxes.series = gearbox_power_moments.gearbox_series);"

    Debug.Print sSQL
    DoCmd.RunSQL sSQL
    
    DoCmd.SetWarnings True
    sSQL = "SELECT " & sTab1 & ".* FROM " & sTab1 & " ;"
    Set rs1 = CurrentDb.OpenRecordset(sSQL, dbOpenSnapshot)
    With rs1
        .MoveLast
        sMsg = "Procedure completed normally. " & vbCrLf & CStr(Nz(.RecordCount)) & " records inserted into " & sTab1
        .Close
    End With
    Set rs1 = Nothing
    MsgBox sMsg, vbInformation, sProcName

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